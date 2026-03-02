import pandas as pd


class ExcelEngine:

    def combine_sheets(self, excel_data):

        dataframes = []

        for sheet_name, rows in excel_data.items():

            if not rows:
                continue

            df = pd.DataFrame(rows)

            # ✅ Normalize column names
            df.columns = df.columns.str.lower().str.strip().str.replace(" ", "_")

            for col in df.columns:

                # ==============================
                # 🔥 SMART DATE CONVERSION
                # ==============================
                if "date" in col:

                    # Try normal datetime conversion (for string dates)
                    converted = pd.to_datetime(df[col], errors="coerce")

                    # If most values failed → try Excel serial number conversion
                    if converted.isna().sum() > len(df) * 0.5:
                        converted = pd.to_datetime(
                            pd.to_numeric(df[col], errors="coerce"),
                            origin="1899-12-30",
                            unit="D",
                            errors="coerce"
                        )

                    df[col] = converted

                # ==============================
                # 🔥 NUMERIC CONVERSION
                # ==============================
                if any(keyword in col for keyword in ["salary", "bonus", "age", "rating", "experience"]):
                    df[col] = pd.to_numeric(df[col], errors="coerce")

            dataframes.append(df)

        if not dataframes:
            return pd.DataFrame()

        final_df = pd.concat(dataframes, ignore_index=True)

        return final_df
