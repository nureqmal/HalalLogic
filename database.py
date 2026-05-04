import sqlite3

def get_connection():
    return sqlite3.connect("ihcs.db", check_same_thread=False)

def init_db():
    conn = get_connection()
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        password TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS company (
        user_id INTEGER,
        name TEXT,
        address TEXT,
        contact TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS policy (
        user_id INTEGER,
        content TEXT
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS materials (
        user_id INTEGER,
        name TEXT,
        supplier TEXT,
        halal_cert TEXT,
        expiry TEXT
    )
    """)

    conn.commit()
    conn.close()
