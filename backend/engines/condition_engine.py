import pandas as pd


class ConditionEngine:

    def apply_conditions(self, df, conditions):

        if not conditions:
            return df

        filtered_df = df.copy()

        for condition in conditions:

            column = condition.get("column")
            operator = condition.get("operator")
            value = condition.get("value")

            # Skip invalid columns
            if column not in filtered_df.columns:
                continue

            # Handle numeric columns
            if pd.api.types.is_numeric_dtype(filtered_df[column]):
                filtered_df = self._apply_numeric(
                    filtered_df, column, operator, value
                )

            # Handle datetime columns
            elif pd.api.types.is_datetime64_any_dtype(filtered_df[column]):
                filtered_df = self._apply_date(
                    filtered_df, column, operator, value
                )

            # Handle string/category columns
            else:
                filtered_df = self._apply_string(
                    filtered_df, column, operator, value
                )

        return filtered_df.reset_index(drop=True)

    # -----------------------------------------------------
    # NUMERIC FILTER
    # -----------------------------------------------------
    def _apply_numeric(self, df, column, operator, value):

        df[column] = pd.to_numeric(df[column], errors="coerce")

        if operator == "=":
            return df[df[column] == float(value)]

        if operator == ">":
            return df[df[column] > float(value)]

        if operator == "<":
            return df[df[column] < float(value)]

        if operator == ">=":
            return df[df[column] >= float(value)]

        if operator == "<=":
            return df[df[column] <= float(value)]

        if operator == "between" and isinstance(value, list):
            return df[
                (df[column] >= float(value[0])) &
                (df[column] <= float(value[1]))
            ]

        return df

    # -----------------------------------------------------
    # DATE FILTER
    # -----------------------------------------------------
    def _apply_date(self, df, column, operator, value):

        df[column] = pd.to_datetime(df[column], errors="coerce")

        compare_value = pd.to_datetime(value, errors="coerce")

        if operator == "=":
            return df[df[column] == compare_value]

        if operator == ">":
            return df[df[column] > compare_value]

        if operator == "<":
            return df[df[column] < compare_value]

        if operator == ">=":
            return df[df[column] >= compare_value]

        if operator == "<=":
            return df[df[column] <= compare_value]

        if operator == "between" and isinstance(value, list):
            start = pd.to_datetime(value[0], errors="coerce")
            end = pd.to_datetime(value[1], errors="coerce")
            return df[(df[column] >= start) & (df[column] <= end)]

        return df

    # -----------------------------------------------------
    # STRING FILTER
    # -----------------------------------------------------
    def _apply_string(self, df, column, operator, value):

        df[column] = df[column].astype(str)

        if operator == "=":
            return df[
                df[column].str.lower() == str(value).lower()
            ]

        return df