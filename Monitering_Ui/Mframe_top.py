import os
from PyQt6.QtWidgets import QLabel, QFrame, QHBoxLayout, QVBoxLayout
from PyQt6.QtCore import Qt, QTimer, QDateTime, QSize
from PyQt6.QtGui import QPixmap, QMovie
from db_manager import IMAGE_PATH


class FrameTop(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background-color: rgb(15,23,42); border-radius: 10px;")

        top_layout = QHBoxLayout(self)
        top_layout.setContentsMargins(10, 5, 10, 5)
        top_layout.setSpacing(10)

        # ------------------------------------
        # Left Image (VLBI antenna logo etc)
        # ------------------------------------
        # Use IMAGE_PATH for logo
        pixmap = QPixmap(os.path.join(IMAGE_PATH, "logo.png"))
        if not pixmap.isNull():
            pixmap = pixmap.scaled(
                250, 250,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )

        label_image = QLabel()
        label_image.setPixmap(pixmap)
        label_image.setFixedSize(250, 120)

        top_layout.addWidget(
            label_image,
            alignment=Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )

        # ------------------------------------
        # Titles
        # ------------------------------------
        title_layout = QVBoxLayout()
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(2)
        title_layout.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

        label_title = QLabel("측지 VLBI 시스템 모니터링")
        label_title.setStyleSheet("color:#38bdf8; font-family: 'Malgun Gothic'; font-weight:bold; font-size:35pt;")

        label_subtitle = QLabel("Geodetic VLBI System Monitoring")
        label_subtitle.setStyleSheet("color:#94a3b8; font-size:16pt;")

        title_layout.addWidget(label_title)
        title_layout.addWidget(label_subtitle)

        top_layout.addLayout(title_layout, stretch=1)

        # ------------------------------------
        # TIME LABEL (KST / UTC)
        # ------------------------------------
        self.time_label = QLabel()
        self.time_label.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.time_label.setStyleSheet(
            "color:white; font-size:25pt; font-weight:bold;"
        )

        time_layout = QVBoxLayout()
        time_layout.addWidget(
            self.time_label,
            alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )

        top_layout.addLayout(time_layout)

        # ------------------------------------
        # COMMUNICATION STATUS ICON (On/Off)
        # ------------------------------------
        self.icon_comm = QLabel()
        self.icon_comm.setFixedSize(100, 100)

        # Use IMAGE_PATH for status icons
        self.comm_on_movie = QMovie(os.path.join(IMAGE_PATH, "parsing_on.gif")) # Animated GIF for connected state
        self.comm_on_movie.setScaledSize(QSize(100, 100))
        
        self.comm_off_pixmap = QPixmap(os.path.join(IMAGE_PATH, "parsing_off.png")).scaled(
            100, 100, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
        )

        # 초기 상태를 'OFF'으로 설정
        self.icon_comm.setPixmap(self.comm_off_pixmap)

        top_layout.addWidget(
            self.icon_comm,
            alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )

        # Timer for clock tick
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        self.update_time()

    # ==================================================================
    # TIME UPDATE
    # ==================================================================
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
    # ==================================================================
    # COMMUNICATION ICON CHANGE
    # ==================================================================
    def set_comm_status(self, ok: bool):
        """통신 상태 아이콘을 제어.
        연결되면(ok=True) 애니메이션 GIF를 표시하고, 연결이 끊어지면(ok=False) 정적 PNG를 표시
        """
        if ok:
            # If connected, show and start the animated GIF
            self.icon_comm.setMovie(self.comm_on_movie)
            self.comm_on_movie.start()
        else:
            # If disconnected, stop the animation and show the static PNG
            self.comm_on_movie.stop()
            self.icon_comm.setPixmap(self.comm_off_pixmap)
