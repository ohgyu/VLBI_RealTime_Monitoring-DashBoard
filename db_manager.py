import sqlite3
import os

DB_PATH = r"D:\GUI\VLBI_DB.db"
IMAGE_PATH = r"D:\GUI\image"


def get_connection(readonly: bool = False):
    if readonly:
        # 읽기 전용 → PRAGMA 절대 건드리면 안 됨
        conn = sqlite3.connect(
            f"file:{DB_PATH}?mode=ro",
            uri=True,
            timeout=5.0,
            check_same_thread=False
        )
        return conn

    # 쓰기 모드
    conn = sqlite3.connect(
        DB_PATH,
        timeout=5.0,
        check_same_thread=False
    )

    # WAL 설정
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA busy_timeout=5000;")
    return conn


def fetch_event_logs(limit=10):
    conn = get_connection(readonly=True)  # 🔥 이렇게
    cursor = conn.cursor()
    try:
        cursor.execute(f"""
            SELECT datetime, message
            FROM Event
            ORDER BY datetime DESC
            LIMIT {limit}
        """)
        rows = cursor.fetchall()
    except Exception as e:
        print("DB 오류:", e)
        rows = []
    finally:
        conn.close()
    return rows
