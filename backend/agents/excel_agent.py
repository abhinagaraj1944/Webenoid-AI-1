import pandas as pd
from engines.intent_engine import IntentEngine
from engines.query_engine import QueryEngine
from engines.schema_engine import SchemaEngine
from engines.data_cleaner import DataCleaner
from engines.python_engine import PythonEngine

class ExcelAgent:

    def __init__(self, query_engine=None):
        self.intent_engine = IntentEngine()
        self.query_engine = query_engine or QueryEngine()
        self.schema_engine = SchemaEngine()
        self.cleaner = DataCleaner()
        self.python_engine = PythonEngine()

    # ============================================
    # COMBINE MULTIPLE SHEETS (SAFE VERSION)
    # ============================================
    def combine_sheets(self, data):

        if not data:
            return None

        df = None

        # Case 1: list of rows
        if isinstance(data, list):
            try:
                df = pd.DataFrame(data)
            except Exception:
                return None

        # Case 2: dict of sheets
        elif isinstance(data, dict):

            dfs = []

            for key, value in data.items():
                if isinstance(value, list):
                    try:
                        dfs.append(pd.DataFrame(value))
                    except Exception:
                        continue

            if dfs:
                df = pd.concat(dfs, ignore_index=True)

        if df is None:
            return None

        # 🔥 REMOVE DUPLICATE COLUMNS (CRITICAL FIX)
        df = df.loc[:, ~df.columns.duplicated()]

        return df

    # ============================================
    # MAIN RUN
    # ============================================
    # MAIN RUN
    # ============================================
    def run(self, question, data):

        if not data:
            return {
                "success": False,
                "type": "error",
                "message": "No data found"
            }

        dfs_dict = {}

        # Case 1: list of rows (single sheet)
        if isinstance(data, list):
            try:
                df = pd.DataFrame(data)
                if not df.empty:
                    dfs_dict["Sheet1"] = self.cleaner.clean(df)
            except Exception:
                pass

        # Case 2: dict of sheets
        elif isinstance(data, dict):
            for sheet_name, sheet_data in data.items():
                if isinstance(sheet_data, list):
                    try:
                        df = pd.DataFrame(sheet_data)
                        if not df.empty:
                            dfs_dict[sheet_name] = self.cleaner.clean(df)
                    except Exception:
                        pass

        if not dfs_dict:
            return {
                "success": False,
                "type": "error",
                "message": "Could not parse Excel data"
            }
            
        for name, df in dfs_dict.items():
            print(f"DEBUG: Sheet '{name}' | Columns: {list(df.columns)}")

        print("Using Python Engine for Dynamic Querying on Sheets:", list(dfs_dict.keys()))

        # Execute query using LLM Python execution
        result = self.python_engine.run_dynamic_query(question, dfs_dict)

        return result