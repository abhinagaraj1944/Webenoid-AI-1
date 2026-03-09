"""
Run this once to create the webenoid_ai_db database automatically.
Usage: python create_db.py
"""
import psycopg2
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT
import os
from dotenv import load_dotenv

load_dotenv()

# Connect to the default 'postgres' database first
conn = psycopg2.connect(
    host="localhost",
    port=5432,
    user="postgres",
    password="Happy@1944"  # Your actual password
)
conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)

cursor = conn.cursor()

# Check if database already exists
cursor.execute("SELECT 1 FROM pg_database WHERE datname = 'webenoid_ai_db'")
exists = cursor.fetchone()

if not exists:
    cursor.execute("CREATE DATABASE webenoid_ai_db")
    print("✅ Database 'webenoid_ai_db' created successfully!")
else:
    print("ℹ️ Database 'webenoid_ai_db' already exists.")

cursor.close()
conn.close()

print("🎉 Done! Now restart your uvicorn server.")
