import os
import json
from openai import OpenAI


class IntentEngine:

    def __init__(self):
        self.client = OpenAI(
            api_key=os.getenv("OPENROUTER_API_KEY"),
            base_url="https://openrouter.ai/api/v1"
        )

    def parse_intent(self, question, schema):

        prompt = f"""
You are an Excel analytics AI.

Return ONLY valid JSON.
Do NOT explain anything.

Schema:
All Columns: {schema.get("all_columns")}
Numeric Columns: {schema.get("numeric_columns")}
Date Columns: {schema.get("date_columns")}
Category Columns: {schema.get("category_columns")}

Rules:
- Supported operations: count, list, aggregation, chart, table
- Operators allowed: =, >, <, >=, <=, between
- Convert date phrases like "after 2022" to "2022-01-01"
- If no aggregation, set it to null
- If no conditions, return empty list

Return format:
{{
  "operation": "",
  "conditions": [],
  "aggregation": null,
  "group_by": null,
  "top_n": null
}}

User Question:
{question}
"""

        response = self.client.chat.completions.create(
            model="openai/gpt-4o-mini",   # Recommended stable model
            messages=[
                {"role": "user", "content": prompt}
            ],
            temperature=0
        )

        content = response.choices[0].message.content.strip()

        return json.loads(content)