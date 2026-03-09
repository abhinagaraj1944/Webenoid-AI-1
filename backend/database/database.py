from sqlalchemy import Column, Integer, String, Text, DateTime, create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import datetime
import os
from dotenv import load_dotenv

load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL")

# Create engine and session factory
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class QueryHistory(Base):
    __tablename__ = "query_history"

    id            = Column(Integer, primary_key=True, index=True)

    # 1️⃣ User Prompt
    question      = Column(Text, nullable=False)

    # 2️⃣ AI Response (full text / answer)
    ai_response   = Column(Text)

    # 3️⃣ Chart Type (e.g. "bar", "pie", "line" or None)
    chart_type    = Column(String(50))

    # 4️⃣ Time Asked
    created_at    = Column(DateTime, default=datetime.datetime.utcnow)

    # 5️⃣ User Info
    user_name     = Column(String(100))
    user_email    = Column(String(200))

    # Extra metadata
    response_type = Column(String(50))  # e.g. "count", "chart", "table", "message"

class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), nullable=False)
    email = Column(String(200), unique=True, index=True, nullable=False)
    phone = Column(String(20), nullable=False)
    hashed_password = Column(String(200), nullable=False)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

# Create tables in the database (safe: won't overwrite existing tables)
def init_db():
    Base.metadata.create_all(bind=engine)

if __name__ == "__main__":
    init_db()
    print("✅ Database initialized!")
