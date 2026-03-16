import sys
import os
import sqlite3
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout
from PyQt6.QtCore import QTimer, QObject, pyqtSignal, QRunnable, QThreadPool
from datetime import datetime, UTC

from Monitering_Ui.Mframe_top import FrameTop
from Monitering_Ui.Mframe_summary import FrameSummary
from Monitering_Ui.Mframe_left import MFrameLeft
from Monitering_Ui.Mframe_footer import FrameFooter

from db_manager import get_connection

import traceback
def exception_hook(exctype, value, tb):
    print(''.join(traceback.format_exception(exctype, value, tb)))
    sys.exit(1)

if __name__ == "__main__":
    sys.excepthook = exception_hook

# ==================================================================
# Worker Thread for Database Queries
# ==================================================================
class WorkerSignals(QObject):
    """실행 중인 작업자 스레드에서 사용할 수 있는 신호를 정의"""
    finished = pyqtSignal()
    error = pyqtSignal(str)
    result = pyqtSignal(dict)

class StatusWorker(QRunnable):
    """데이터베이스에서 장치 상태를 가져오는 작업자 스레드"""
    def __init__(self):
        super().__init__()
        self.signals = WorkerSignals()

    def run(self):
        try:
            conn = get_connection(readonly=True)
            cursor = conn.cursor()

            status_data = {"PCAL": None, "FlatMirror": None, "CalChop": None}

            # 1. Fetch latest PCAL and FlatMirror that ARE NOT NULL
            query_pcal_fm = """
                        SELECT PCAL, FlatMirror
                        FROM Calibration 
                        WHERE band = '43ghz' 
                          AND PCAL IS NOT NULL 
                          AND FlatMirror IS NOT NULL
                        ORDER BY datetime DESC 
                        LIMIT 1
                    """
            cursor.execute(query_pcal_fm)
            res_pcal_fm = cursor.fetchone()
            if res_pcal_fm:
                status_data["PCAL"] = res_pcal_fm[0]
                status_data["FlatMirror"] = res_pcal_fm[1]

            # 2. Fetch the latest CalChop that IS NOT NULL
            query_calchop = """
                        SELECT CalChop
                        FROM Calibration 
                        WHERE band = '43ghz' 
                          AND CalChop IS NOT NULL
                        ORDER BY datetime DESC 
                        LIMIT 1
                    """
            cursor.execute(query_calchop)
            res_calchop = cursor.fetchone()
            if res_calchop:
                status_data["CalChop"] = res_calchop[0]

            conn.close()
            self.signals.result.emit(status_data)

        except Exception as e:
            print(f"Worker Error: {e}")
            self.signals.error.emit(str(e))
        finally:
            self.signals.finished.emit()


# ==================================================================
# Main Window
# ==================================================================
class MonitoringWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VLBI Real-Time Monitoring")
        self.setStyleSheet("background-color:#0F172A;")

        self.threadpool = QThreadPool()

        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        # UI Frames
        self.frame_top = FrameTop(parent=self)
        self.frame_summary = FrameSummary(parent=self)
        self.frame_left = MFrameLeft()
        self.frame_footer = FrameFooter(parent=self)

        main_layout.addWidget(self.frame_top, stretch=1)
        main_layout.addWidget(self.frame_summary, stretch=1)
        center_layout = QHBoxLayout()
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(5)
        center_layout.addWidget(self.frame_left, stretch=1)
        main_layout.addLayout(center_layout, stretch=8)
        main_layout.addWidget(self.frame_footer)

        # Connections
        QTimer.singleShot(0, lambda: setattr(self.frame_left, "summary", self.frame_summary))
        QTimer.singleShot(10, lambda: self.frame_left.update_all_thresholds())

        # Main Timer
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.on_timer_tick)
        self.timer.start(30_000)
        self.on_timer_tick()

    # ==================================================================
    # 통신 상태 체크 - (UTC DeprecationWarning 제거)
    # ==================================================================
    def check_connection_status(self):
        try:
            conn = get_connection(readonly=True)
            cur = conn.cursor()
            cur.execute("SELECT Parsed_at FROM _Parsing_history_ ORDER BY Parsed_at DESC LIMIT 1")
            row = cur.fetchone()
            conn.close()
            if not row: return False
            last_dt = datetime.fromisoformat(row[0]).replace(tzinfo=UTC)
            now_utc = datetime.now(UTC)
            return (now_utc - last_dt).total_seconds() < 60
        except Exception:
            return False

    def update_summary_status(self):
        """데이터베이스 쿼리를 백그라운드 스레드에서 실행"""
        worker = StatusWorker()
        worker.signals.result.connect(self.frame_summary.update_device_status)
        worker.signals.error.connect(self.handle_worker_error)
        self.threadpool.start(worker)

    def handle_worker_error(self, error_msg):
        """작업자 스레드에서 발생하는 오류를 처리"""
        print(f"ERROR from background worker: {error_msg}")

    def on_timer_tick(self):
        # Update data-intensive parts
        self.frame_left.update_all_thresholds()
        self.frame_left.refresh_expanded()
        self.update_summary_status() # Trigger background worker

        # U부하가 적은 부분을 직접 업데이트
        ok = self.check_connection_status()
        self.frame_top.set_comm_status(ok)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MonitoringWindow()
    QTimer.singleShot(0, win.showMaximized)
    sys.exit(app.exec())
