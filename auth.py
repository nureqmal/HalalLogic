import bcrypt
from database import get_connection

def register(username, password):
    conn = get_connection()
    c = conn.cursor()

    hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt())

    try:
        c.execute("INSERT INTO users (username, password) VALUES (?, ?)", (username, hashed))
        conn.commit()
        return True
    except:
        return False

def login(username, password):
    conn = get_connection()
    c = conn.cursor()

    c.execute("SELECT id, password FROM users WHERE username=?", (username,))
    user = c.fetchone()

    if user:
        if bcrypt.checkpw(password.encode(), user[1]):
            return user[0]
    return None
