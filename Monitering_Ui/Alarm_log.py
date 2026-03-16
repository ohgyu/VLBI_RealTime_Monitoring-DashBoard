import sqlite3

class AlarmLogger:
    def __init__(self, db_path):
        self.db_path = db_path
        self._init_table()

    def _init_table(self):
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS Alarm_log (
                datetime TEXT NOT NULL,
                device TEXT NOT NULL ,
                alarm_level TEXT,
                message TEXT 
            )
        """)
        conn.commit()
        conn.close()

    def log(self, log_time, device, level, message):
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO Alarm_log (datetime, device, alarm_level, message)
            VALUES (?, ?, ?, ?)
        """, (log_time, device, level, message))
        conn.commit()
        conn.close()
