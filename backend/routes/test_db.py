import mysql.connector

print("🔌 Trying manual connection...")
try:
    conn = mysql.connector.connect(
        host="localhost",
        user="root",
        password="Nav@1234",
        database="sales_data",
        port=3306,
        connection_timeout=5
    )
    print("✅ Connected successfully!")
    conn.close()
except Exception as e:
    print(f"❌ Failed: {e}")
