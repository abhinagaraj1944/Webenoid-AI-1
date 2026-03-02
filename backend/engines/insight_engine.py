class InsightEngine:

    def generate_insight(self, result):

        if not result or not result.get("success"):
            return None

        operation = result.get("operation")

        # --------------------------------------------
        # GROUP AGGREGATION INSIGHT
        # --------------------------------------------
        if operation and operation.endswith("_by_group"):

            data = result.get("data", [])
            if not data:
                return None

            first_row = data[0]
            group_col = list(first_row.keys())[0]
            value_col = list(first_row.keys())[1]

            top = data[0]
            bottom = data[-1]

            return (
                f"{top[group_col]} has the highest {value_col} "
                f"with {round(top[value_col], 2)}. "
                f"{bottom[group_col]} has the lowest at "
                f"{round(bottom[value_col], 2)}."
            )

        # --------------------------------------------
        # GROWTH INSIGHT
        # --------------------------------------------
        if operation == "growth":

            data = result.get("data", [])
            if not data:
                return None

            max_growth = max(
                data,
                key=lambda x: x.get("growth_percent", -999)
            )

            return (
                f"Highest growth observed in "
                f"{max_growth.get(list(max_growth.keys())[0])} "
                f"with {round(max_growth.get('growth_percent', 0), 2)}% increase."
            )

        # --------------------------------------------
        # COMPARISON INSIGHT
        # --------------------------------------------
        if operation == "comparison":

            data = result.get("data", [])
            if len(data) == 2:

                first = data[0]
                second = data[1]

                metric = list(first.keys())[1]

                winner = first if first[metric] > second[metric] else second

                return (
                    f"{winner[list(winner.keys())[0]]} performs better "
                    f"in {metric} with value {round(winner[metric], 2)}."
                )

        # --------------------------------------------
        # SIMPLE AGGREGATION
        # --------------------------------------------
        if operation in ["sum", "average", "max", "min"]:

            data = result.get("data", {})
            if data:

                key = list(data.keys())[0]
                value = data[key]

                return f"The {operation} of {key} is {round(value, 2)}."

        return None