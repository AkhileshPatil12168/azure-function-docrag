import sqlite3

def get_db_conn():
    conn = sqlite3.connect("app.db")
    conn.row_factory = sqlite3.Row
    return conn