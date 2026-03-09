import pandas as pd


class DataCleaner:

    def clean(self, df):

        if df is None or df.empty:
            return df

        # 🔥 REMOVE DUPLICATE COLUMNS (CRITICAL FIX)
        df = df.loc[:, ~df.columns.duplicated()]

        # Strip column names and remove newlines
        df.columns = [
            str(col).replace("\n", " ").replace("\r", " ").strip()
            for col in df.columns
        ]

        for col in df.columns:

            if not isinstance(col, str):
                continue

            try:
                series = df[col]

                # 🔥 If duplicate still returns DataFrame → skip
                if isinstance(series, pd.DataFrame):
                    continue

                # -----------------------
                # OBJECT COLUMN CLEANING
                # -----------------------
                if pd.api.types.is_object_dtype(series):

                    df[col] = (
                        series
                        .astype(str)
                        .str.strip()
                        .replace({"nan": None, "": None})
                    )

                # -----------------------
                # NUMERIC CLEANING
                # -----------------------
                elif pd.api.types.is_numeric_dtype(series):

                    df[col] = pd.to_numeric(
                        series,
                        errors="coerce"
                    )

            except Exception:
                continue

        return df