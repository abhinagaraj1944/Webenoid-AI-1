import pandas as pd


class ColumnProfiler:

    def profile(self, df):

        schema = {
            "entity": None,
            "measure": None,
            "category_columns": [],
            "date_columns": [],
            "id_columns": []
        }

        if df is None or df.empty:
            return schema

        # 🔥 REMOVE DUPLICATE COLUMNS (CRITICAL FIX)
        df = df.loc[:, ~df.columns.duplicated()]

        total_rows = len(df)

        for col in df.columns:

            # Skip invalid column names
            if not isinstance(col, str) or not col.strip():
                continue

            try:
                series = df[col]

                # 🔥 If duplicate column somehow returns DataFrame → skip
                if isinstance(series, pd.DataFrame):
                    continue

                col_lower = col.lower()
                unique_count = series.nunique(dropna=True)
                unique_ratio = unique_count / total_rows if total_rows > 0 else 0

                # -------------------------
                # DATE COLUMN
                # -------------------------
                if pd.api.types.is_datetime64_any_dtype(series):
                    schema["date_columns"].append(col)
                    continue

                # -------------------------
                # NUMERIC COLUMN (MEASURE)
                # -------------------------
                if pd.api.types.is_numeric_dtype(series):
                    if schema["measure"] is None:
                        schema["measure"] = col
                    continue

                # -------------------------
                # TEXT COLUMN
                # -------------------------
                if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):

                    # Skip obvious non-group columns
                    if any(x in col_lower for x in ["id", "email", "phone"]):
                        schema["id_columns"].append(col)
                        continue

                    # ENTITY (high uniqueness)
                    if unique_ratio > 0.85:
                        if schema["entity"] is None:
                            schema["entity"] = col
                        else:
                            schema["id_columns"].append(col)
                        continue

                    # CATEGORY (moderate uniqueness)
                    if 1 < unique_count < total_rows * 0.8:
                        schema["category_columns"].append(col)
                        continue

                    schema["id_columns"].append(col)

            except Exception:
                continue

        return schema