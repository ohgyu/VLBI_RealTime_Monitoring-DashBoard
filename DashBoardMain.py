# main.py
import sys
import logging
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QSplitter
from PyQt6.QtCore import Qt

from DashBoard_Ui.frame_top import FrameTop
from DashBoard_Ui.frame_left import FrameLeft
from DashBoard_Ui.frame_center import FrameCenter



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VLBI Dashboard")

        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Top
        self.frame_top = FrameTop(parent=self)
        main_layout.addWidget(self.frame_top, stretch=1)

        # Left & Center
        self.frame_left = FrameLeft()
        self.frame_center = FrameCenter()

        self.frame_left.set_frame_center(self.frame_center)
        self.frame_center.frame_left = self.frame_left



        # ★ PyQt6 전용 QSplitter
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self.frame_left)
        splitter.addWidget(self.frame_center)

        # 초기 폭
        splitter.setSizes([260, 1200])

        # ★ 왼쪽 패널 고정 설정 (여기가 핵심)
        splitter.setCollapsible(0, False)  # 왼쪽 접기 방지
        splitter.setStretchFactor(0, 0)  # 왼쪽은 고정
        splitter.setStretchFactor(1, 1)  # 오른쪽만 늘어남
        splitter.setHandleWidth(0)  # 드래그 손잡이 제거

        main_layout.addWidget(splitter, stretch=9)

        self.frame_top.frame_center = self.frame_center
        self.frame_left.item_selected.connect(self.on_item_selected)

    def on_item_selected(self, group_name, item_name, is_parent):
        if is_parent:
            return  # 부모 클릭은 UI에서만 처리 (FrameLeft에서 끝)
        else:
            self.frame_center.show_child_graph(group_name, item_name)
            self.frame_left.refresh_child_selection()


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,  # 개발 중이면 DEBUG
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    app = QApplication(sys.argv)
    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec())
