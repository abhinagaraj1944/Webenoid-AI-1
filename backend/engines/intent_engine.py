import re


class IntentEngine:

    def normalize(self, text):
        # Keep + and numbers
        return re.sub(r'[^a-z0-9 +]', '', text.lower())

    def detect_intent(self, question, schema=None):

        print("INTENT ENGINE RUNNING")
        print("QUESTION RECEIVED:", question)

        question_clean = self.normalize(question)

        plan = {
            "operation": "details",
            "aggregation_type": None,
            "chart_type": None
        }

        # =========================
        # COUNT
        # =========================
        if any(phrase in question_clean for phrase in [
            "how many",
            "count",
            "number of"
        ]):
            plan["operation"] = "count"
            return plan

        # =========================
        # CHART
        # =========================
        if any(word in question_clean for word in [
            "chart", "graph", "bar", "column", "pie"
        ]):
            plan["operation"] = "chart"

            if "pie" in question_clean:
                plan["chart_type"] = "pie"
            else:
                plan["chart_type"] = "bar"

            return plan

        # =========================
        # AGGREGATIONS
        # =========================
        if any(word in question_clean for word in ["total", "sum"]):
            plan["operation"] = "aggregation"
            plan["aggregation_type"] = "sum"
            return plan

        if any(word in question_clean for word in ["average", "avg"]):
            plan["operation"] = "aggregation"
            plan["aggregation_type"] = "mean"
            return plan

        if any(word in question_clean for word in ["highest", "maximum", "top"]):
            plan["operation"] = "aggregation"
            plan["aggregation_type"] = "max"
            return plan

        if any(word in question_clean for word in ["lowest", "minimum"]):
            plan["operation"] = "aggregation"
            plan["aggregation_type"] = "min"
            return plan

        # =========================
        # LIST
        # =========================
        if question_clean.startswith("what are") or question_clean.startswith("list"):
            plan["operation"] = "list"
            return plan

        return plan