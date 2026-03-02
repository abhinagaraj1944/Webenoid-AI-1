from fastapi import APIRouter
from agents.excel_agent import ExcelAgent
from engines.query_engine import QueryEngine   # ✅ ADD THIS
from models.schemas import QueryRequest
import numpy as np
import pandas as pd
import traceback


router = APIRouter()

# ✅ CREATE QUERY ENGINE INSTANCE
query_engine = QueryEngine()

# ✅ INJECT INTO EXCEL AGENT
agent = ExcelAgent(query_engine)

# =====================================
# JSON SAFE CONVERTER
# =====================================

def make_json_safe(data):

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

@router.post("/analyze")
async def analyze(request: QueryRequest):

    try:
        result = agent.run(request.question, request.data)
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
# UNIVERSAL DASHBOARD ENDPOINT
# =====================================

@router.post("/dashboard")
async def create_dashboard(request: QueryRequest):

    try:
        df = agent.combine_sheets(request.data)

        if df is None or df.empty:
            return {"success": False, "error": "No data found"}

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

        # =====================================
        # BASIC KPIs
        # =====================================

        total_rows = len(df)

        primary_numeric = numeric_columns[0] if numeric_columns else None

        total_numeric_sum = 0
        if primary_numeric:
            total_numeric_sum = pd.to_numeric(
                df[primary_numeric], errors="coerce"
            ).sum()

        # =====================================
        # TOP CATEGORY GROUPING
        # =====================================

        grouped_data = []

        if categorical_columns and primary_numeric:

            primary_category = categorical_columns[0]

            grouped_data = (
                df.groupby(primary_category)[primary_numeric]
                .sum()
                .sort_values(ascending=False)
                .head(5)
                .reset_index()
                .to_dict(orient="records")
            )

        # =====================================
        # TREND ANALYSIS
        # =====================================

        trend_data = []

        if date_columns and primary_numeric:

            date_col = date_columns[0]

            trend_data = (
                df.groupby(df[date_col].dt.year)[primary_numeric]
                .sum()
                .reset_index()
                .to_dict(orient="records")
            )

        dashboard_data = {
            "total_rows": total_rows,
            "primary_numeric_column": primary_numeric,
            "total_numeric_sum": float(total_numeric_sum),
            "top_grouped_data": grouped_data,
            "trend_data": trend_data
        }

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