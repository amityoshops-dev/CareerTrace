import sqlite3
from .config import DB_PATH, DATA_DIR

def get_conn():
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS applications (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company TEXT NOT NULL,
        role TEXT NOT NULL,
        jd TEXT NOT NULL,
        source TEXT,
        recruiter TEXT,
        recruiter_email TEXT,
        status TEXT NOT NULL DEFAULT 'DISCOVERED',
        match_summary TEXT,
        ats_keywords TEXT,
        missing_keywords TEXT,
        tailored_summary TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    CREATE INDEX IF NOT EXISTS idx_app_status ON applications(status);
    """)
    conn.commit()
    conn.close()
