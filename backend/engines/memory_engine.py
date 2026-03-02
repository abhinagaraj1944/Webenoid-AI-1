class MemoryEngine:

    def __init__(self):
        self.history = []
        self.last_query_plan = None
        self.last_filters = None

    def add_interaction(self, question, query_plan):
        self.history.append({
            "question": question,
            "query_plan": query_plan
        })

        self.last_query_plan = query_plan
        self.last_filters = query_plan.get("filters", [])

    def get_context(self):
        return {
            "history": self.history[-5:],
            "last_query_plan": self.last_query_plan,
            "last_filters": self.last_filters
        }