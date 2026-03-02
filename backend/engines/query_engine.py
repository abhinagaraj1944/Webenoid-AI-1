import pandas as pd
import difflib
import re
import numpy as np


class QueryEngine:

    def __init__(self):
        self.last_working_df = None

    # =====================================================
    # JSON SAFE
    # =====================================================
    def make_json_safe(self, data):

        if isinstance(data, dict):
            return {k: self.make_json_safe(v) for k, v in data.items()}

        if isinstance(data, list):
            return [self.make_json_safe(i) for i in data]

        if isinstance(data, np.integer):
            return int(data)

        if isinstance(data, np.floating):
            return float(data)

        if isinstance(data, np.bool_):
            return bool(data)

        if pd.isna(data):
            return None

        return data

    # =====================================================
    # COLUMN TYPES
    # =====================================================
    def get_numeric_columns(self, df):
        return [
            col for col in df.columns
            if pd.api.types.is_numeric_dtype(df[col])
        ]

    def get_text_columns(self, df):
        return [
            col for col in df.columns
            if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col])
        ]

    # =====================================================
    # FUZZY MATCH COLUMN
    # =====================================================
    def fuzzy_match_column(self, question, columns):

        if not columns:
            return None

        question = question.lower()

        for col in columns:
            if col.lower() in question:
                return col

        tokens = re.findall(r"[a-z0-9]+", question)

        for col in columns:
            col_clean = col.lower().replace(" ", "")
            for token in tokens:
                if token in col_clean:
                    return col

        best_match = None
        highest_score = 0

        for col in columns:
            score = difflib.SequenceMatcher(
                None, col.lower(), question
            ).ratio()

            if score > highest_score:
                highest_score = score
                best_match = col

        if highest_score > 0.3:
            return best_match

        return None

    # =====================================================
    # SMART FILTER (FULL + FIXED)
    # =====================================================
    def smart_filter(self, df, question):

        filtered_df = df.copy().reset_index(drop=True)
        question = question.lower().strip()

        # ==========================================
        # FOLLOW-UP CONTEXT
        # ==========================================
        follow_words = ["them", "those", "their", "that", "above"]
        if any(word in question for word in follow_words):
            if self.last_working_df is not None:
                return self.last_working_df.copy().reset_index(drop=True)

        # ==========================================
        # SAFE Department Detection (FIXED)
        # ==========================================
        if "Department" in filtered_df.columns:

            dept_series = (
                filtered_df["Department"]
                .dropna()
                .astype(str)
                .str.lower()
                .str.strip()
            )

            for dept in dept_series.unique():
                if isinstance(dept, str) and dept:
                    pattern = r"\b" + re.escape(dept) + r"\b"
                    if re.search(pattern, question):
                        mask = (
                            filtered_df["Department"]
                            .astype(str)
                            .str.lower()
                            .str.strip() == dept
                        )
                        filtered_df = filtered_df[mask].reset_index(drop=True)
                        break

        # ==========================================
        # AUTO CONVERT DATE COLUMNS
        # ==========================================
        for col in filtered_df.columns:
            if "date" in col.lower():
                filtered_df[col] = pd.to_datetime(filtered_df[col], errors='coerce')

        # ==========================================
        # Numeric & Date Operators
        # ==========================================
        numeric_cols = self.get_numeric_columns(filtered_df)

        # Greater Than / After
        for m in re.finditer(r'(greater than|more than|above|over|>|after|older than)\s*(\d+(?:\.\d+)?)', question):
            val = float(m.group(2))
            col = self.fuzzy_match_column(question, numeric_cols)

            if 1900 <= val <= 2100:
                for c in filtered_df.columns:
                    if pd.api.types.is_datetime64_any_dtype(filtered_df[c]):
                        filtered_df = filtered_df[filtered_df[c].dt.year > val].reset_index(drop=True)
                        break
            elif col:
                filtered_df = filtered_df[
                    pd.to_numeric(filtered_df[col], errors='coerce') > val
                ].reset_index(drop=True)

        # Less Than / Before
        for m in re.finditer(r'(less than|under|below|<|before|younger than)\s*(\d+(?:\.\d+)?)', question):
            val = float(m.group(2))
            col = self.fuzzy_match_column(question, numeric_cols)

            if 1900 <= val <= 2100:
                for c in filtered_df.columns:
                    if pd.api.types.is_datetime64_any_dtype(filtered_df[c]):
                        filtered_df = filtered_df[filtered_df[c].dt.year < val].reset_index(drop=True)
                        break
            elif col:
                filtered_df = filtered_df[
                    pd.to_numeric(filtered_df[col], errors='coerce') < val
                ].reset_index(drop=True)

        # ==========================================
        # TOKEN FILTERING (UNCHANGED)
        # ==========================================
        stopwords = {
            "how", "many", "are", "is", "the", "in", "of", "who", "what", "which",
            "show", "me", "working", "work", "create", "bar", "chart", "pie",
            "line", "area", "graph", "count", "total", "department", "departments",
            "employees", "employee", "with", "and", "under", "over", "above", "below",
            "than", "greater", "less", "more", "before", "after", "joined", "have"
        }

        clean_q = re.sub(r'(greater than|more than|above|over|>|after|older than)\s*\d+', '', question)
        clean_q = re.sub(r'(less than|under|below|<|before|younger than)\s*\d+', '', clean_q)

        tokens = [
            t for t in re.findall(r"[a-z0-9]+", clean_q)
            if t not in stopwords and len(t) > 1
        ]

        text_cols = self.get_text_columns(filtered_df)
        other_cols = [c for c in filtered_df.columns if c not in text_cols]
        other_cols = sorted(other_cols, key=lambda c: filtered_df[c].nunique())
        search_cols = text_cols + other_cols

        for token in tokens:
            token_matched = False

            for col in search_cols:
                series = filtered_df[col].astype(str).str.lower().str.strip()

                if token in series.values:
                    filtered_df = filtered_df[series == token].reset_index(drop=True)
                    token_matched = True
                    break

            if not token_matched:
                for col in search_cols:
                    series_str = filtered_df[col].astype(str).str.lower()
                    mask = series_str.str.contains(rf"\b{re.escape(token)}\b", na=False)
                    if mask.any():
                        filtered_df = filtered_df[mask].reset_index(drop=True)
                        break

        print("FILTERED ROWS:", len(filtered_df))
        return filtered_df

    # =====================================================
    # AUTO GROUP COLUMN
    # =====================================================
    def auto_group_column(self, df):
        text_cols = self.get_text_columns(df)
        if not text_cols:
            return None
        return max(text_cols, key=lambda c: df[c].nunique())

    # =====================================================
    # EXECUTE (FULL ORIGINAL LOGIC KEPT)
    # =====================================================
    def execute(self, plan, df, question):

        if df is None or (isinstance(df, pd.DataFrame) and df.empty):
            return {"success": False, "error": "No data available"}

        question = question.lower().strip()
        operation = plan.get("operation")

        df.columns = df.columns.str.strip()
        df = df.loc[:, ~df.columns.duplicated()]

        df = self.smart_filter(df, question)

        numeric_cols = self.get_numeric_columns(df)
        text_cols = self.get_text_columns(df)

        numeric_col = self.fuzzy_match_column(question, numeric_cols)
        group_col = self.fuzzy_match_column(question, text_cols)

        # DETAILS
        if "who" in question or "detail" in question or "show" in question:
            self.last_working_df = df.copy()
            return {
                "success": True,
                "operation": "details",
                "data": self.make_json_safe(df.to_dict(orient="records"))
            }

        # COUNT
        if operation == "count":
            self.last_working_df = df.copy()
            return {
                "success": True,
                "operation": "count",
                "row_count": int(len(df))
            }

        # LIST
        if operation == "list":
            column = group_col or self.auto_group_column(df)
            if not column:
                return {"success": False, "error": "No column found to list"}

            self.last_working_df = df.copy()
            values = df[column].dropna().unique().tolist()

            return {
                "success": True,
                "operation": "list",
                "column": column,
                "values": self.make_json_safe(values)
            }

        # AGGREGATION
        if operation == "aggregation":
            if not numeric_col:
                return {"success": False, "error": "Numeric column not detected"}

            agg_type = plan.get("aggregation_type")
            df[numeric_col] = pd.to_numeric(df[numeric_col], errors="coerce")

            if agg_type == "sum":
                value = df[numeric_col].sum()
            elif agg_type == "mean":
                value = df[numeric_col].mean()
            elif agg_type == "max":
                value = df[numeric_col].max()
            elif agg_type == "min":
                value = df[numeric_col].min()
            else:
                return {"success": False, "error": "Invalid aggregation type"}

            self.last_working_df = df.copy()
            return {
                "success": True,
                "operation": agg_type,
                "column": numeric_col,
                "value": round(float(value), 2)
            }

        # CHART
        if operation == "chart":

            chart_type = plan.get("chart_type", "bar")
            category_col = group_col or self.auto_group_column(df)

            if not category_col:
                return {"success": False, "error": "No category column found"}

            if numeric_col:
                df[numeric_col] = pd.to_numeric(df[numeric_col], errors="coerce")
                grouped = df.groupby(category_col)[numeric_col].sum().reset_index()
                value_column = numeric_col
            else:
                grouped = df.groupby(category_col).size().reset_index(name="Count")
                value_column = "Count"

            if chart_type == "pie":
                total = grouped[value_column].sum()
                grouped["Percentage"] = grouped[value_column] / total * 100

            self.last_working_df = df.copy()

            return {
                "success": True,
                "type": "chart",
                "chart_type": chart_type,
                "category_column": category_col,
                "value_columns": [value_column],
                "data": self.make_json_safe(grouped.to_dict(orient="records"))
            }

        # DEFAULT
        self.last_working_df = df.copy()

        return {
            "success": True,
            "operation": "details",
            "data": self.make_json_safe(df.head(100).to_dict(orient="records"))
        }