# db.py - SQLite Database initialization

import sqlite3
import os

# Create safe writable path
APP_FOLDER = os.path.join(os.getenv('APPDATA'), 'HypeProduction')
os.makedirs(APP_FOLDER, exist_ok=True)

DB_NAME = os.path.join(APP_FOLDER, 'database.db')

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        article TEXT, card TEXT, color TEXT, size TEXT,
        qty INTEGER, component TEXT, print_opt TEXT, date TEXT
    )''')
    conn.commit()
    conn.close()
