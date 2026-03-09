from fastapi import APIRouter
from agents.excel_agent import ExcelAgent
from engines.query_engine import QueryEngine
from engines.dashboard_engine import DashboardEngine
from models.schemas import QueryRequest
from database.database import SessionLocal, QueryHistory
import json
import numpy as np
import pandas as pd
import traceback


router = APIRouter()

# ✅ CREATE ENGINE INSTANCES
query_engine = QueryEngine()
dashboard_engine = DashboardEngine()

# ✅ INJECT INTO EXCEL AGENT
agent = ExcelAgent(query_engine)

# =====================================
# JSON SAFE CONVERTER
# =====================================

def make_json_safe(data):

    if isinstance(data, pd.DataFrame):
        return make_json_safe(data.to_dict(orient='records'))

    if isinstance(data, pd.Series):
        return make_json_safe(data.tolist())

    if isinstance(data, dict):
        return {k: make_json_safe(v) for k, v in data.items()}

    if isinstance(data, list):
        return [make_json_safe(i) for i in data]

    if isinstance(data, np.integer):
        return int(data)

    if isinstance(data, np.floating):
        return float(data)

    if isinstance(data, pd.Timestamp):
        return str(data)

    return data


# =====================================
# ANALYZE ENDPOINT
# =====================================

GREETINGS = [
    "hello", "hi", "hey", "good morning", "good afternoon", "good evening",
    "howdy", "what's up", "sup", "greetings", "hola", "how are you",
    "who are you", "what are you", "what can you do", "help me", "help",
    "thanks", "thank you", "ok", "okay", "bye", "goodbye", "nice"
]

@router.post("/analyze")
async def analyze(request: QueryRequest):

    try:
        q = request.question.strip().lower().rstrip("!?.,:;")

        # Handle greetings and casual messages with a friendly reply
        if any(q == g or q.startswith(g) for g in GREETINGS):
            friendly_replies = {
                "hello": "👋 Hello! I'm Webenoid AI. Ask me anything about your Excel data — like totals, charts, counts, or details!",
                "hi": "👋 Hi there! Ready to analyze your Excel data. What would you like to know?",
                "hey": "👋 Hey! I'm here to help you explore your spreadsheet. Ask me a question about your data!",
                "who are you": " I'm Webenoid AI — your intelligent Excel assistant. I can analyze your spreadsheet data, create charts, count rows, find totals, and much more!",
                "what are you": " I'm Webenoid AI — your intelligent Excel assistant. I can analyze your spreadsheet data, create charts, count rows, find totals, and much more!",
                "what can you do": "💡 I can help you with:\n• Counts (how many rows match a condition)\n• Totals, averages, max, min\n• Charts (bar, pie, line)\n• Filtering and listing records\n• Full table details\n\nJust ask me a question about your data!",
                "help": "💡 Try asking things like:\n• 'How many students got A+?'\n• 'What is the total sales?'\n• 'Show me a chart of marks by subject'\n• 'Who has the highest score?'\n• 'List all students with grade B'",
                "thanks": "😊 You're welcome! Let me know if you have more questions about your data.",
                "thank you": "😊 You're welcome! Let me know if you have more questions about your data.",
            }
            # Find best match
            reply = next((v for k, v in friendly_replies.items() if q.startswith(k)), 
                         "👋 I'm Webenoid AI! Ask me anything about your Excel data — I can count, sum, chart, filter, and more.")
            return {"success": True, "operation": "message", "text": reply}

        result = agent.run(request.question, request.data)
        
        # ✅ SAVE TO DATABASE — All 5 fields
        try:
            db = SessionLocal()

            # 1️⃣ User Prompt → already in request.question
            # 2️⃣ AI Response text (readable answer)
            safe_result = make_json_safe(result)
            ai_response_text = safe_result.get("text") or safe_result.get("answer") or safe_result.get("message") or json.dumps(safe_result)

            print(f"DEBUG: Operation: {safe_result.get('operation')}, Type: {safe_result.get('type')}")
            if safe_result.get("operation") == "conversation":
                print(f"DEBUG: Conversational Response: {ai_response_text}")

            # 3️⃣ Chart Type — extracted from the result if it contains chart/graph data
            chart_type = None
            if safe_result.get("type") == "chart":
                chart_type = safe_result.get("chartType") or safe_result.get("chart_type") or "chart"

            history = QueryHistory(
                question      = request.question,            # 1️⃣ User Prompt
                ai_response   = ai_response_text,            # 2️⃣ AI Response
                chart_type    = chart_type,                  # 3️⃣ Chart Type
                # created_at is auto-set by the model        # 4️⃣ Time Asked
                user_name     = request.user_name,           # 5️⃣ User Info (name)
                user_email    = request.user_email,          # 5️⃣ User Info (email)
                response_type = safe_result.get("type", "message"),
            )
            db.add(history)
            db.commit()
            db.close()
            print("💾 Saved query to DB — all 5 fields stored.")
        except Exception as db_err:
            print(f"❌ Database error: {db_err}")

        print("FINAL RESPONSE:", result)
        return result

    except Exception as e:
        print("🔥 FULL ERROR TRACE:")
        traceback.print_exc()
        return {
            "success": False,
            "type": "error",
            "message": str(e)
        }



# =====================================
# DYNAMIC DASHBOARD ENDPOINT
# =====================================

@router.post("/dashboard")
async def create_dashboard(request: QueryRequest):

    try:
        df = agent.combine_sheets(request.data)

        if df is None or df.empty:
            return {"success": False, "error": "No data found in the spreadsheet."}

        # 🔥 REMOVE DUPLICATE COLUMNS (CRITICAL)
        df = df.loc[:, ~df.columns.duplicated()]

        # Remove fully empty columns
        df = df.dropna(axis=1, how="all")

        # =====================================
        # AUTO DATE DETECTION (SAFE)
        # =====================================

        date_columns = []

        for col in df.columns:

            if not isinstance(col, str):
                continue

            series = df[col]

            # Skip if duplicate caused DataFrame
            if isinstance(series, pd.DataFrame):
                continue

            # Skip numeric columns
            if pd.api.types.is_numeric_dtype(series):
                continue

            if pd.api.types.is_object_dtype(series):

                sample = series.dropna().astype(str).head(10)

                if sample.str.contains(r"[-/]").any():

                    converted = pd.to_datetime(
                        series,
                        errors="coerce"
                    )

                    if converted.notna().sum() > len(df) * 0.6:
                        df[col] = converted
                        date_columns.append(col)

        # =====================================
        # NUMERIC DETECTION
        # =====================================

        numeric_columns = []

        for col in df.columns:
            series = df[col]

            if isinstance(series, pd.DataFrame):
                continue

            if pd.api.types.is_numeric_dtype(series) and col not in date_columns:
                numeric_columns.append(col)

        # =====================================
        # CATEGORICAL DETECTION
        # =====================================

        categorical_columns = [
            col for col in df.columns
            if col not in numeric_columns and col not in date_columns
        ]

        if not numeric_columns and not categorical_columns:
            return {"success": False, "error": "Could not detect any meaningful columns for a dashboard."}

        # =====================================
        # DYNAMIC DASHBOARD VIA LLM ENGINE
        # =====================================

        print(f"📊 Building dynamic dashboard | Numeric: {numeric_columns} | Categorical: {categorical_columns} | Dates: {date_columns}")

        dashboard_data = dashboard_engine.build(df, numeric_columns, categorical_columns, date_columns)

        return {
            "success": True,
            "dashboard": make_json_safe(dashboard_data)
        }

    except Exception as e:
        print("🔥 DASHBOARD ERROR TRACE:")
        traceback.print_exc()
        return {
            "success": False,
            "error": str(e)
        }