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
        if isinstance(data, pd.DataFrame):
            return self.make_json_safe(data.to_dict(orient='records'))
        if isinstance(data, pd.Series):
            return self.make_json_safe(data.tolist())
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
        
        # Safe check for scalar nulls
        try:
            if pd.isna(data) is True:
                return None
        except:
            pass
            
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
3. NEVER return hardcoded or static results. You MUST ALWAYS write actual pandas code that queries the DataFrame to produce the result dynamically. For example, NEVER write `result = {{"operation": "list", "values": []}}` — instead, write pandas code that filters the DataFrame and extracts the values.
4. FOLLOW-UP QUERIES (CRITICAL — READ CAREFULLY):
   - If the user asks a follow-up question like "who are they?", "what are their names?", "list them", "tell me about them", etc., you MUST:
     a) Look at the CONVERSATION HISTORY above to find the PREVIOUS script that was executed.
     b) Copy the EXACT SAME pandas filter/condition from that previous script.
     c) Apply that filter to the DataFrame to get the matching rows.
     d) Identify the name/identifier column (look for columns like 'Name', 'Student Name', 'Employee Name', 'Student', or the first string-type column).
     e) Extract the names from those filtered rows.
   - "who are they?", "give me their names", "list them", "tell me names" → return ONLY the names using the "list" operation.
   - "give me their details", "show full details", "show their info", "tell me about them", "give details" → return ALL columns using the "details" operation with the full table.

   EXAMPLE: If the previous script was:
     result = {{"operation": "count", "row_count": dfs['Student Marks'][dfs['Student Marks']['Grade'] == 'B+'].shape[0], ...}}
   And the user now asks "who are they?", your script MUST be:
     filtered = dfs['Student Marks'][dfs['Student Marks']['Grade'] == 'B+']
     name_col = [c for c in filtered.columns if 'name' in c.lower()]
     col = name_col[0] if name_col else filtered.columns[0]
     result = {{"operation": "list", "values": filtered[col].tolist()}}

5. You MUST assign your final answer to a variable named `result`.
6. The `result` variable MUST follow this JSON-serializable dictionary format to match the frontend, depending on the operation:

- For count:
  result = {{"operation": "count", "row_count": numeric_count, "is_percentage": True/False, "title": "Optional descriptive title"}}
- For a chart/graph:
  result = {{"type": "chart", "chart_type": "bar" or "pie" or "line" or "area" or "doughnut" or "scatter" or "stacked" or "funnel" or "waterfall", "category_column": "x_col_name", "value_columns": ["y_col_name"], "data": [{{"x_col_name": "a", "y_col_name": 10}}], "is_percentage": True ONLY if the user explicitly asked for a percentage-based chart (e.g. "show percentages", "in percent", "distribution in %"), else False }}
  CRITICAL RULE: ALWAYS use the RAW original metric values (e.g. actual Quantity, actual TotalPrice) in the 'data' array — NEVER pre-compute percentages yourself.
  - The 'value_columns' key should ALWAYS be the original numeric column name (e.g. "Quantity", "TotalPrice"), NOT a pre-computed "Percentage" column.
  - When is_percentage is True, the frontend will automatically calculate percentages from the raw values and handle label formatting.
  - NEVER create a computed percentage column like total_sales['Percentage'] = ... and use it as the value. Always use the direct aggregation result.
- For a single numerical aggregation (sum, max, min, average, percentage calculation):
  result = {{"operation": "sum", "data": {{"value": numeric_value, "label": "Name/identifier of the entity (e.g. student name) if applicable, else None"}}, "is_percentage": True/False, "title": "Optional descriptive title"}}
  IMPORTANT: For max/min operations (e.g. "highest marks", "lowest salary"), you MUST find the row with that max/min value using .idxmax()/.idxmin() and include the entity's name/identifier in the "label" field. For sum/average where no single entity applies, set "label" to None.
- For listing names only (e.g. "who are they?", "give me their names"):
  result = {{"operation": "list", "values": ["Name1", "Name2", ...]}}
  You MUST populate this list by querying the DataFrame. NEVER return an empty list without filtering first.
- For full details or a filtered table (e.g. "give me their details", "show full info"):
  result = {{"operation": "details", "data": filtered_df.to_dict(orient='records')}}
- For comparison questions (e.g. "Compare X vs Y", "Compare Store A and Store B sales", "difference between Laptop and Phone"):
  result = {{"operation": "comparison", "title": "Descriptive comparison title", "metric": "The numeric column being compared (e.g. Sales, Price, Count)", "items": [{{"name": "Item A name", "value": numeric_value_A}}, {{"name": "Item B name", "value": numeric_value_B}}], "difference": abs(value_A - value_B), "difference_percent": round(abs(value_A - value_B) / min(value_A, value_B) * 100, 2) if min(value_A, value_B) > 0 else 0, "higher": "Name of the item with higher value"}}
  You MUST calculate ALL values dynamically from the DataFrame. Support comparing by sum, count, average, or any aggregation the user implies. If comparing more than 2 items, include all items in the "items" array and set "difference" and "higher" based on the top 2.
- For questions asking for an Excel formula or function (e.g. "what is the formula of ADD?", "formula for average sales", "how to calculate sum"):
  result = {{"operation": "formula", "formula": "The exact Excel formula (e.g. =SUM(A1:A10))", "explanation": "Brief explanation of how the formula works"}}
  NEVER try to write pandas code for this. Just assign the dictionary above to `result`. Provide a realistic example formula for the requested operation.
- For greetings, thank you, conversational messages, or questions NOT related to data analysis (e.g. "Hi", "Hello", "Thank you", "How are you?", "What can you do?"):
  result = {{"operation": "conversation", "message": "Your friendly response here"}}
  For greetings: respond warmly and mention you can help analyze their Excel data.
  For thank you: respond politely.
  For "what can you do": explain your capabilities (count, sum, filter, chart, list, details, comparison, etc.).
  NEVER try to write pandas code for conversational messages.

Do NOT print the result. ONLY assign it to the variable `result`.
Assume pandas is imported as pd and numpy as np. They are already available.
Important: Make sure your pandas code handles potential string/numeric data type conversions if needed.
7. ROBUST COLUMN MATCHING: If the user's question mentions a column that isn't an exact match (e.g., "Dept" vs "Department"), use pandas' `.str.contains()` or simple list comprehensions to find the intended column. DO NOT simply state that the data is missing if a similar column name exists.
REMINDER: NEVER return hardcoded/static values. ALWAYS write pandas code that dynamically queries the data.
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

            import re
            script = response.choices[0].message.content.strip()

            # Robust Regex: Extract ONLY the code inside the first markdown block if present.
            # This prevents crashes if the AI adds "Sure! Here is the code: ..." or other text.
            code_match = re.search(r'```(?:python)?(.*?)```', script, re.DOTALL)
            if code_match:
                script = code_match.group(1).strip()
            else:
                # Fallback: strip backticks if Regex didn't catch the block
                script = script.replace('```python', '').replace('```', '').strip()
            
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
            return {"success": False, "type": "error", "message": "I'm sorry, I'm having trouble analyzing your request. Could you try rephrasing your question?"}
