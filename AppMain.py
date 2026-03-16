import sys
import logging
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QTimer

from DashBoardMain import MainWindow
from MoniteringMain import MonitoringWindow

class AppController:
    def __init__(self):
        self.dashboard = MainWindow()
        self.monitoring = MonitoringWindow()

    def show(self):
        self.dashboard.showMaximized()
        QTimer.singleShot(0, self.monitoring.showMaximized)

def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    app = QApplication(sys.argv)
    controller = AppController()
    controller.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()