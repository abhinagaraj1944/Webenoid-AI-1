import os
import json
import pandas as pd
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()


class AIEngine:

    def __init__(self):
        self.client = OpenAI(
            api_key=os.getenv("OPENROUTER_API_KEY"),
            base_url=os.getenv("OPENROUTER_BASE_URL")
        )

    # =====================================================
    # MAKE DATAFRAME JSON SAFE
    # =====================================================
    def make_json_safe_df(self, df):

        safe_df = df.copy()

        for col in safe_df.columns:

            if str(safe_df[col].dtype).startswith("datetime"):
                safe_df[col] = safe_df[col].astype(str)

            safe_df[col] = safe_df[col].apply(
                lambda x: str(x) if isinstance(x, pd.Timestamp) else x
            )

        return safe_df

    # =====================================================
    # SMART SAMPLE
    # =====================================================
    def build_sample(self, df):

        head_part = df.head(3)

        random_part = df.sample(
            min(5, len(df)),
            random_state=42
        )

        sample_df = pd.concat([head_part, random_part]).drop_duplicates()

        return sample_df

    # =====================================================
    # BUILD CATEGORICAL VALUES (🔥 IMPORTANT FIX)
    # =====================================================
    def build_categorical_values(self, df):

        categorical_values = {}

        for col in df.columns:

            unique_count = df[col].nunique()

            # Only small categorical columns
            if 1 < unique_count <= 20:
                categorical_values[col] = (
                    df[col]
                    .dropna()
                    .unique()
                    .tolist()
                )

        return categorical_values

    # =====================================================
    # GENERATE QUERY PLAN
    # =====================================================
    def generate_query_plan(self, question, df, context):

        try:
            columns = list(df.columns)

            safe_df = self.make_json_safe_df(df)

            sample_df = self.build_sample(safe_df)

            categorical_values = self.build_categorical_values(safe_df)

            # 🔥 SEND FULL CONTEXT TO AI
            schema_info = {
                "columns": columns,
                "row_count": len(df),
                "sample_data": sample_df.to_dict(orient="records"),
                "categorical_values": categorical_values,
                "conversation_history": context.get("history", []),
                "last_query_plan": context.get("last_query_plan"),
                "last_filters": context.get("last_filters")
            }

            system_prompt = """
You are an advanced Excel Data Analytics AI Planner.

Return ONLY valid JSON.
No explanation.
No markdown.

==================================================
STRICT RULES
==================================================

- Never invent column names.
- Use column names EXACTLY as provided.
- If filtering by value, use values from categorical_values when available.

==================================================
FOLLOW-UP RULES
==================================================

If:
- question refers to previous result
- does not mention new filter value
- like "who are they", "show them", "their details"

Then:
- reuse_context = true
- operation = "details"
- filters = []

==================================================
INTENT RULES
==================================================

- how many → count
- total/sum → sum
- average → average
- highest → max
- lowest → min
- unique/list → list
- top → top
- chart/graph → chart

==================================================
JSON FORMAT
==================================================

{
  "operation": null,
  "target_column": null,
  "group_by": null,
  "filters": [],
  "limit": null,
  "reuse_context": false,
  "chart_type": "ColumnClustered | Pie | Line | BarClustered | Area | XYScatter"
}
"""

            response = self.client.chat.completions.create(
                model="openai/gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {
                        "role": "user",
                        "content": f"DATA:\n{json.dumps(schema_info)}\n\nQUESTION:\n{question}"
                    }
                ],
                temperature=0,
                max_tokens=600
            )

            content = response.choices[0].message.content.strip()

            if content.startswith("```"):
                content = content.split("```")[1]
                content = content.replace("json", "").strip()

            plan = json.loads(content)

            if "filters" in plan:
                plan["filters"] = [
                    f for f in plan["filters"]
                    if f.get("column") and f.get("value") is not None
                ]

            return plan

        except Exception as e:
            print("AI Planning Error:", e)

            return {
                "operation": "details",
                "target_column": None,
                "group_by": None,
                "filters": [],
                "limit": None,
                "reuse_context": False,
                "chart_type": None
            }