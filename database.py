import sqlite3
import config
from contextlib import contextmanager

def init_db():
    if not config.DB_PATH.exists():
        conn = sqlite3.connect(config.DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
                       CREATE TABLE IF NOT EXISTS reports
                       (
                           id
                           INTEGER
                           PRIMARY
                           KEY
                           AUTOINCREMENT,
                           specimen_id
                           TEXT
                           UNIQUE,
                           report_data
                           TEXT,
                           created_at
                           TIMESTAMP
                           DEFAULT
                           CURRENT_TIMESTAMP
                       )
                       ''')

        conn.commit()
        conn.close()

def get_db():
    # check_same_thread=False: 비동기/동기 혼용 시 스레드 에러 방지 옵션
    conn = sqlite3.connect(config.DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()
