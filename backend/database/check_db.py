import psycopg2

conn = psycopg2.connect(
    host="localhost",
    port=5432,
    user="postgres",
    password="Happy@1944"
)
cur = conn.cursor()
cur.execute("SELECT datname FROM pg_database WHERE datistemplate = false;")
dbs = cur.fetchall()
print("Databases found in PostgreSQL:")
for db in dbs:
    print(f"  - {db[0]}")
conn.close()
