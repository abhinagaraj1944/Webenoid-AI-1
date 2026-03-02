from pydantic import BaseModel
from typing import Dict, Any

class QueryRequest(BaseModel):
    question: str
    data: Dict[str, Any]
