import re


class IntentEngine:

    def normalize(self, text):
        return re.sub(r'[^a-z0-9 ]', '', text.lower())

    def detect_intent(self, question, schema=None):

        question_clean = self.normalize(question)

        plan = {
            "operation": "details",
            "aggregation_type": None,
            "conditions": [],
            "reuse_context": False
        }

        # =====================================
        # COUNT
        # =====================================
        if "how many" in question_clean or "count" in question_clean:
            plan["operation"] = "count"
            return plan

        # =====================================
        # PIE CHART / PERCENTAGE
        # =====================================
        if "pie" in question_clean or "percent" in question_clean:
            plan["operation"] = "chart"
            plan["chart_type"] = "pie"
            return plan

        # =====================================
        # BAR / COLUMN CHART
        # =====================================
        if "bar" in question_clean or "column" in question_clean:
            plan["operation"] = "chart"
            plan["chart_type"] = "bar"
            return plan

        # =====================================
        # GENERAL CHART 
        # =====================================
        if "chart" in question_clean or "graph" in question_clean:
            plan["operation"] = "chart"
            plan["chart_type"] = "bar"
            return plan

        # =====================================
        # LIST / WHO
        # =====================================
        if question_clean.startswith("what are") or question_clean.startswith("list") or question_clean.startswith("who"):
            plan["operation"] = "list"
            return plan

        # =====================================
        # DETAILS
        # =====================================
        if "detail" in question_clean:
            plan["operation"] = "details"
            return plan

        # =====================================
        # TOTAL / SUM
        # =====================================
        if "total" in question_clean or "sum" in question_clean:
            plan["operation"] = "aggregation"
            plan["aggregation_type"] = "sum"
            return plan

        # =====================================
        # AVERAGE
        # =====================================
        if "average" in question_clean or "avg" in question_clean:
            plan["operation"] = "aggregation"
            plan["aggregation_type"] = "mean"
            return plan

        # =====================================
        # GENERAL CHART 
        # =====================================
        if "chart" in question_clean or "graph" in question_clean:
            plan["operation"] = "chart"
            plan["chart_type"] = "bar"
            return plan

        # =====================================
        # LIST
        # =====================================
        if question_clean.startswith("what are") or question_clean.startswith("list"):
            plan["operation"] = "list"
            return plan

        return plan