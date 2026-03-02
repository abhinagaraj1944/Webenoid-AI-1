import os
import json
import pandas as pd
import numpy as np
from openai import OpenAI
from dotenv import load_dotenv
import traceback

load_dotenv()

class PythonEngine:
    def __init__(self):
        self.client = OpenAI(
            api_key=os.getenv("OPENROUTER_API_KEY"),
            base_url=os.getenv("OPENROUTER_BASE_URL")
        )
        self.chat_history = []

    def make_json_safe(self, data):
        if isinstance(data, dict):
            return {k: self.make_json_safe(v) for k, v in data.items()}
        if isinstance(data, list):
            return [self.make_json_safe(i) for i in data]
        if isinstance(data, np.integer):
            return int(data)
        if isinstance(data, np.floating):
            return float(data)
        if isinstance(data, pd.Timestamp):
            return str(data)
        if pd.isna(data):
            return None
        return data

    def run_dynamic_query(self, question, dfs_dict):
        try:
            # 1. Build schema info for all sheets
            schema_info = ""
            for sheet_name, df in dfs_dict.items():
                schema_info += f"Sheet Name: '{sheet_name}'\n"
                schema_info += f"Columns: {list(df.columns)}\n"
                
                # Try to get some samples safely
                sample_data = df.head(3).to_dict(orient="records")
                safe_sample = self.make_json_safe(sample_data)
                schema_info += f"Sample Data: {json.dumps(safe_sample)}\n\n"

            # Update and format history
            self.chat_history.append(f"User: {question}")
            if len(self.chat_history) > 6:
                self.chat_history = self.chat_history[-6:]
            history_str = "\n".join(self.chat_history[:-1]) # Don't include the current question in history block

            system_prompt = f"""
You are an expert Python data analyst. You are provided with a dictionary of pandas DataFrames named `dfs`, where keys are sheet names and values are DataFrames.
Your task is to write a Python script using pandas to answer the user's question accurately.

The data schemas are:
{schema_info}

RECENT CONVERSATION HISTORY (Use this to understand follow-up contexts, like "who are they?" or "how about for X?"):
{history_str}

STRICT RULES:
1. Return ONLY valid Python code. NO markdown formatting, NO explanations, NO backticks (```python ... ```). Start your code immediately.
2. The user's question must be answered by analyzing the DataFrames in the `dfs` dictionary.
3. If the user asks a follow-up question referencing previous data (e.g. "who are they?", "what is their average salary?"), you MUST look at the conversational history to identify what filters were previously applied, and explicitly re-apply those exact same pandas filters in your new script to answer the current question.
4. You MUST assign your final answer to a variable named `result`.
5. The `result` variable MUST follow this JSON-serializable dictionary format to match the frontend, depending on the operation:

- For count:
  result = {{"operation": "count", "row_count": numeric_count, "is_percentage": True/False, "title": "Optional descriptive title"}}
- For a chart/graph:
  result = {{"type": "chart", "chart_type": "bar" or "pie" or "line", "category_column": "x_col_name", "value_columns": ["y_col_name"], "data": [{{"x_col_name": "a", "y_col_name": 10}}], "is_percentage": True if the user asked for percentages to be shown else False }}
  (If the user asks for counts AND percentages in a chart, just provide the raw numeric counts in a single `y_col_name` and set `is_percentage: True`. Do NOT provide a second percentage column; the frontend will use the raw counts to natively graph and display both.)
- For a single numerical aggregation (sum, max, min, average, percentage calculation):
  result = {{"operation": "sum", "data": {{"value": numeric_value}}, "is_percentage": True/False, "title": "Optional descriptive title"}}
- For listing items (e.g. unique names, departments):
  result = {{"operation": "list", "values": ["item1", "item2"]}}
- For general details or a filtered table:
  result = {{"operation": "details", "data": list_of_dicts}} # Use df.to_dict(orient='records') here

Do NOT print the result. ONLY assign it to the variable `result`.
Assume pandas is imported as pd and numpy as np. They are already available.
Important: Make sure your pandas code handles potential string/numeric data type conversions if needed.
"""

            print("--- ASKING AI TO WRITE PANDAS CODE ---")
            
            # 2. Call OpenAI to get the script
            response = self.client.chat.completions.create(
                model="openai/gpt-4o-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"User Question: {question}"}
                ],
                temperature=0,
                max_tokens=2000
            )

            script = response.choices[0].message.content.strip()

            # Clean up the script just in case there's markdown
            if script.startswith("```python"):
                script = script[9:]
            elif script.startswith("```"):
                script = script[3:]
            if script.endswith("```"):
                script = script[:-3]
            script = script.strip()
            
            print("--- GENERATED SCRIPT ---")
            print(script)
            print("------------------------")

            # 3. Execute the script safely
            local_vars = {"dfs": dfs_dict, "pd": pd, "np": np}
            exec(script, {}, local_vars)
            
            raw_result = local_vars.get("result", {})
            
            safe_result = self.make_json_safe(raw_result)
            safe_result["success"] = True
            
            # Save the successful script logic to history so follow-ups know the filters
            self.chat_history.append(f"AI Script Executed:\n{script}")
            
            return safe_result

        except Exception as e:
            traceback.print_exc()
            return {"success": False, "type": "error", "message": f"AI Execution Error: {str(e)}"}
