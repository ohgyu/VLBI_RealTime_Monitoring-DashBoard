# DashBoard_Ui/frame_top.py
import os
from PyQt6.QtWidgets import QLabel, QPushButton, QFrame, QHBoxLayout, QVBoxLayout
from PyQt6.QtCore import Qt, QTimer, QDateTime
from PyQt6.QtGui import QPixmap
from db_manager import IMAGE_PATH



class FrameTop(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background-color: rgb(15,23,42); border: none;")

        # 향후 필요 시 center 접근용 (지금은 직접 사용하지 않음)
        self.frame_center = None

        top_layout = QHBoxLayout(self)
        top_layout.setContentsMargins(10, 5, 10, 5)
        top_layout.setSpacing(10)

        # 안테나 사진
        logo_path = os.path.join(IMAGE_PATH, "logo.png")
        pixmap = QPixmap(logo_path)

        if not pixmap.isNull():
            pixmap = pixmap.scaled(
                250, 250,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
        label_image = QLabel()
        label_image.setPixmap(pixmap)
        label_image.setFixedSize(250, 120)
        top_layout.addWidget(label_image, alignment=Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

        # 제목 영역
        title_layout = QVBoxLayout()
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(2)

        label_title = QLabel("측지 VLBI 통계 대시보드")
        label_title.setStyleSheet("color:#38bdf8; font-family: 'Malgun Gothic'; font-weight:bold; font-size:35pt;")
        label_subtitle = QLabel("Geodesic VLBI Statistics Dashboard")
        label_subtitle.setStyleSheet("color:#94a3b8; font-size:16pt;")
        title_layout.addWidget(label_title)
        title_layout.addWidget(label_subtitle)
        title_layout.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

        top_layout.addLayout(title_layout, stretch=1)

        # 시간 레이블 (KST / UTC)
        self.time_label = QLabel()
        self.time_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.time_label.setStyleSheet("color:white; font-size:12pt; font-weight:bold;")

        time_layout = QVBoxLayout()
        time_layout.setContentsMargins(0, 0, 0, 0)
        time_layout.addWidget(self.time_label, alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        top_layout.addLayout(time_layout)

        # 타이머
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        self.update_time()

    def update_time(self):
        now_utc = QDateTime.currentDateTimeUtc()
        now_kst = now_utc.addSecs(9 * 3600)

        kst_str = now_kst.toString("yyyy-MM-dd hh:mm:ss")
        utc_str = now_utc.toString("yyyy-MM-dd hh:mm:ss")

        # 글자 크기 16pt, 줄 간격 추가(line-height)
        self.time_label.setStyleSheet("""
            color: white;
            font-size: 16pt;
            font-weight: bold;
            line-height: 140%;
        """)

        # KST / UTC 사이 한 줄 간격 추가
        self.time_label.setText(
            f"(KST) {kst_str}\n\n"
            f"(UTC) {utc_str}"
        )