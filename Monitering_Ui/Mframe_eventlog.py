# Monitering_Ui/Mframe_eventlog.py
from PyQt6.QtWidgets import QFrame, QVBoxLayout, QLabel, QScrollArea, QWidget
from PyQt6.QtCore import Qt, QTimer
from db_manager import get_connection

class FrameEventLog(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background-color:#0F172A; border-radius:10px;")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(5)

        title = QLabel("EVENT LOG")
        title.setStyleSheet("color:#f97316; font-size:18pt; font-weight:bold;")
        main_layout.addWidget(title)

        # Scroll Area
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("border:none; background-color:#0F172A;")

        self.content = QWidget()
        self.scroll_layout = QVBoxLayout(self.content)
        self.scroll_layout.setContentsMargins(5, 5, 5, 5)
        self.scroll_layout.setSpacing(6)

        self.scroll.setWidget(self.content)
        main_layout.addWidget(self.scroll, stretch=1)

        # Timer for auto-refresh (10 sec)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.reload_logs)
        self.timer.start(10_000)

        # 초기에 로딩
        self.reload_logs()

    # -------------------------------------------------------------
    #  PARSING HISTORY LOAD (LOCK FREE)
    # -------------------------------------------------------------
    def reload_logs(self):
        # 기존 위젯 제거
        for i in reversed(range(self.scroll_layout.count())):
            item = self.scroll_layout.itemAt(i)
            w = item.widget()
            if w is not None:
                w.deleteLater()

        rows = []
        try:
            conn = get_connection(readonly=True)
            cur = conn.cursor()

            # 컬럼: Log_name, Parsed_at
            cur.execute("""
                SELECT Log_name, Parsed_at
                FROM _Parsing_history_
                ORDER BY Parsed_at DESC
                LIMIT 20
            """)

            rows = cur.fetchall()

        except Exception as e:
            err = QLabel(f"Parsing History DB 오류: {e}")
            err.setStyleSheet("color:#fecaca;")
            self.scroll_layout.addWidget(err)
            return

        finally:
            try:
                conn.close()
            except:
                pass

        # 출력
        for logname, parsed_at in rows:
            lbl = QLabel(f"[{parsed_at}]  {logname}")
            lbl.setStyleSheet("""
                QLabel {
                    color:#FCA5A5;
                    font-size:11pt;
                }
            """)
            lbl.setWordWrap(True)
            self.scroll_layout.addWidget(lbl)

        self.scroll_layout.addStretch()
