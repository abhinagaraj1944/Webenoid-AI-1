import pandas as pd
from fuzzywuzzy import process


class AggregationEngine:

    def find_best_column(self, df, keyword):

        if df.empty:
            return None

        columns = list(df.columns)

        # Fuzzy match
        match = process.extractOne(keyword, columns)

        if not match:
            return None

        best_match, score = match

        if score > 60:
            return best_match

        return None


    def execute(self, intent, df):

        if df.empty:
            return {"error": "No data found in any sheet"}

        keyword = intent.get("column")
        operation = intent.get("operation")

        column = self.find_best_column(df, keyword)

        if not column:
            return {"error": "No matching column found"}

        series = df[column]

        # Try numeric conversion
        numeric_series = pd.to_numeric(series, errors="coerce")
        is_numeric = numeric_series.notna().sum() > 0

        # 🔥 LIST OPERATION
        if operation == "list":
            unique_values = series.dropna().unique().tolist()
            return unique_values

        # 🔥 COUNT OPERATION
        if operation == "count":
            return int(series.dropna().count())

        # 🔥 If text column but operation is sum → auto convert to list
        if not is_numeric and operation in ["sum", "average"]:
            unique_values = series.dropna().unique().tolist()
            return unique_values

        # Use numeric series for math
        series = numeric_series

        # 🔥 SUM
        if operation == "sum":
            result = series.sum()

        # 🔥 AVERAGE
        elif operation == "average":
            result = series.mean()

        else:
            return {"error": "Unsupported operation"}

        return {
            "column_used": column,
            "operation": operation,
            "value": float(result)
        }
