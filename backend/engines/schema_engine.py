from engines.column_profiler import ColumnProfiler
import pandas as pd


class SchemaEngine:

    def __init__(self):
        self.profiler = ColumnProfiler()

    def build_schema(self, df):

        raw_schema = self.profiler.profile(df)

        entity = raw_schema.get("entity")
        measure = raw_schema.get("measure")
        category_columns = raw_schema.get("category_columns", [])
        date_columns = raw_schema.get("date_columns", [])

        # Clean empty category columns
        category_columns = [
            col for col in category_columns
            if col and str(col).strip()
        ]

        # Detect numeric columns dynamically
        numeric_columns = [
            col for col in df.columns
            if pd.api.types.is_numeric_dtype(df[col])
        ]

        # Detect date columns dynamically (extra safety)
        for col in df.columns:
            if col not in date_columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    date_columns.append(col)

        # Remove duplicates
        date_columns = list(set(date_columns))

        # Fallback entity detection
        if not entity and category_columns:
            entity = category_columns[0]

        # Ensure measure is numeric
        if measure and measure not in numeric_columns:
            measure = None

        return {
            "entity": entity,
            "measure": measure,
            "all_columns": list(df.columns),
            "numeric_columns": numeric_columns,
            "date_columns": date_columns,
            "category_columns": category_columns
        }