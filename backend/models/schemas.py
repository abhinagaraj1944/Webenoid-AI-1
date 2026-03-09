from pydantic import BaseModel
from typing import Dict, Any, Optional

class QueryRequest(BaseModel):
    question: str
    data: Dict[str, Any]
    user_name: Optional[str] = None   # e.g. "Abhishek"
    user_email: Optional[str] = None  # e.g. "abhishek@example.com"

class SignUpRequest(BaseModel):
    name: str
    email: str
    phone: str
    password: str

class LoginRequest(BaseModel):
    email: str
    password: str
