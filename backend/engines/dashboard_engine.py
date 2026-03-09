import os
import json
import pandas as pd
import numpy as np
from openai import OpenAI
from dotenv import load_dotenv
import traceback

load_dotenv()


class DashboardEngine:
    """
    Uses an LLM to intelligently select the most meaningful KPIs,
    category columns, and numeric columns from any dataset before
    building the dashboard — instead of blindly picking [0].
    """

    def __init__(self):
        self.client = OpenAI(
            api_key=os.getenv("OPENROUTER_API_KEY"),
            base_url=os.getenv("OPENROUTER_BASE_URL")
        )

    def make_json_safe(self, data):
        if isinstance(data, pd.DataFrame):
            return self.make_json_safe(data.to_dict(orient='records'))
        if isinstance(data, pd.Series):
            return self.make_json_safe(data.tolist())
        if isinstance(data, dict):
            return {k: self.make_json_safe(v) for k, v in data.items()}
        if isinstance(data, list):
            return [self.make_json_safe(i) for i in data]
        if isinstance(data, np.integer):
            return int(data)
        if isinstance(data, np.floating):
            return float(data)
        if isinstance(data, pd.Timestamp):
            return str(data)
        try:
            if pd.isna(data) is True:
                return None
        except:
            pass
        return data

    def select_dashboard_columns(self, df: pd.DataFrame, numeric_columns: list, categorical_columns: list, date_columns: list) -> dict:
        """
        Ask the LLM to pick the best KPI columns and chart column for this specific dataset.
        Returns a dict like:
        {
            "primary_numeric": "TotalPrice",
            "secondary_numerics": ["Quantity"],
            "primary_category": "Product",
            "secondary_category": "Region",
            "date_column": "OrderDate",
            "kpi_titles": {
                "total_rows": "Total Orders",
                "primary_numeric": "Total Revenue",
                "secondary_numeric": "Total Units Sold"
            }
        }
        """
        schema_summary = {
            "columns": list(df.columns),
            "numeric_columns": numeric_columns,
            "categorical_columns": categorical_columns,
            "date_columns": date_columns,
            "sample": self.make_json_safe(df.head(3).to_dict(orient="records"))
        }

        prompt = f"""You are a business intelligence expert. 
Given this spreadsheet schema, identify the most meaningful columns for a business dashboard.

Schema:
{json.dumps(schema_summary, indent=2)}

Respond ONLY with a valid JSON object (no markdown, no explanation) with these exact keys:
{{
  "primary_numeric": "<most important numeric column, e.g. Revenue, TotalPrice, Salary>",
  "secondary_numerics": ["<2nd important numeric>", "<3rd important numeric>"],
  "primary_category": "<most meaningful category for grouping, e.g. Product, Department, Region>",
  "secondary_category": "<2nd most useful category column, or null if none>",
  "date_column": "<date column name, or null if none>",
  "kpi_titles": {{
    "total_rows": "<human-friendly label for row count, e.g. 'Total Orders', 'Total Employees'>",
    "primary_numeric": "<human-friendly label for the primary numeric sum, e.g. 'Total Revenue', 'Total Salary'>",
    "secondary_numeric": "<human-friendly label for secondary numeric sum, or null>"
  }}
}}

Rules:
- Choose columns that make business sense (not IDs, not row numbers).
- primary_numeric must be from: {numeric_columns}
- primary_category must be from: {categorical_columns}
- date_column must be from: {date_columns} (or null)
- secondary_numerics may be an empty list if there's only 1 numeric column.
- secondary_category may be null if there's only 1 categorical column.
"""

        try:
            response = self.client.chat.completions.create(
                model="openai/gpt-4o-mini",
                messages=[
                    {"role": "user", "content": prompt}
                ],
                temperature=0,
                max_tokens=500
            )
            raw = response.choices[0].message.content.strip()

            # Strip markdown if present
            import re
            code_match = re.search(r'```(?:json)?(.*?)```', raw, re.DOTALL)
            if code_match:
                raw = code_match.group(1).strip()

            selection = json.loads(raw)
            print("📊 LLM Dashboard Selection:", selection)
            return selection

        except Exception as e:
            print("⚠️ LLM column selection failed, falling back to defaults:", str(e))
            traceback.print_exc()
            # Fallback to first available columns
            return {
                "primary_numeric": numeric_columns[0] if numeric_columns else None,
                "secondary_numerics": numeric_columns[1:3] if len(numeric_columns) > 1 else [],
                "primary_category": categorical_columns[0] if categorical_columns else None,
                "secondary_category": categorical_columns[1] if len(categorical_columns) > 1 else None,
                "date_column": date_columns[0] if date_columns else None,
                "kpi_titles": {
                    "total_rows": "Total Records",
                    "primary_numeric": f"Total {numeric_columns[0]}" if numeric_columns else "Total",
                    "secondary_numeric": f"Total {numeric_columns[1]}" if len(numeric_columns) > 1 else None
                }
            }

    def build(self, df: pd.DataFrame, numeric_columns: list, categorical_columns: list, date_columns: list) -> dict:
        """
        Build a fully dynamic dashboard payload for any dataset.
        """

        # 1. Use LLM to intelligently select columns
        selection = self.select_dashboard_columns(df, numeric_columns, categorical_columns, date_columns)

        primary_numeric = selection.get("primary_numeric") or (numeric_columns[0] if numeric_columns else None)
        secondary_numerics = selection.get("secondary_numerics", [])
        primary_category = selection.get("primary_category") or (categorical_columns[0] if categorical_columns else None)
        secondary_category = selection.get("secondary_category")
        date_column = selection.get("date_column") or (date_columns[0] if date_columns else None)
        kpi_titles = selection.get("kpi_titles", {})

        total_rows = len(df)

        # 2. Compute KPIs
        kpis = [
            {
                "title": kpi_titles.get("total_rows", "Total Records"),
                "value": total_rows,
                "is_percentage": False
            }
        ]

        if primary_numeric:
            primary_series = pd.to_numeric(df[primary_numeric], errors="coerce")
            primary_sum = float(primary_series.sum())
            primary_avg = float(primary_series.mean())

            kpis.append({
                "title": kpi_titles.get("primary_numeric", f"Total {primary_numeric}"),
                "value": primary_sum,
                "is_percentage": False
            })
            kpis.append({
                "title": f"Avg {primary_numeric}",
                "value": round(primary_avg, 2),
                "is_percentage": False
            })

        for sec_col in secondary_numerics[:1]:  # max 1 secondary KPI
            if sec_col in df.columns:
                sec_series = pd.to_numeric(df[sec_col], errors="coerce")
                sec_sum = float(sec_series.sum())
                kpis.append({
                    "title": kpi_titles.get("secondary_numeric", f"Total {sec_col}"),
                    "value": sec_sum,
                    "is_percentage": False
                })

        # 3. Primary grouping chart (Bar)
        top_grouped_data = []
        top_grouped_chart = None
        if primary_category and primary_numeric:
            grouped = (
                df.groupby(primary_category)[primary_numeric]
                .sum()
                .sort_values(ascending=False)
                .head(8)
                .reset_index()
            )
            top_grouped_data = self.make_json_safe(grouped.to_dict(orient="records"))
            top_grouped_chart = {
                "title": f"{primary_numeric} by {primary_category}",
                "chart_type": "ColumnClustered",
                "category_column": primary_category,
                "value_column": primary_numeric,
                "data": top_grouped_data
            }

        # 4. Secondary grouping chart (Pie) for distribution
        secondary_grouped_chart = None
        if secondary_category and primary_numeric:
            sec_grouped = (
                df.groupby(secondary_category)[primary_numeric]
                .sum()
                .sort_values(ascending=False)
                .head(6)
                .reset_index()
            )
            secondary_grouped_chart = {
                "title": f"{primary_numeric} by {secondary_category}",
                "chart_type": "Pie",
                "category_column": secondary_category,
                "value_column": primary_numeric,
                "data": self.make_json_safe(sec_grouped.to_dict(orient="records"))
            }
        elif primary_category and primary_numeric and not secondary_category:
            # Fallback: use primary_category for pie if no secondary
            secondary_grouped_chart = {
                "title": f"Distribution by {primary_category}",
                "chart_type": "Pie",
                "category_column": primary_category,
                "value_column": primary_numeric,
                "data": top_grouped_data
            }

        # 5. Trend line chart (if date column available)
        trend_chart = None
        if date_column and primary_numeric and pd.api.types.is_datetime64_any_dtype(df[date_column]):
            trend = (
                df.groupby(df[date_column].dt.year)[primary_numeric]
                .sum()
                .reset_index()
            )
            trend.columns = ["Year", primary_numeric]
            trend_chart = {
                "title": f"{primary_numeric} Trend by Year",
                "chart_type": "Line",
                "category_column": "Year",
                "value_column": primary_numeric,
                "data": self.make_json_safe(trend.to_dict(orient="records"))
            }

        return {
            "kpis": kpis,
            "primary_chart": top_grouped_chart,
            "secondary_chart": secondary_grouped_chart,
            "trend_chart": trend_chart,
            "meta": {
                "primary_numeric": primary_numeric,
                "primary_category": primary_category,
                "secondary_category": secondary_category,
                "date_column": date_column
            }
        }
