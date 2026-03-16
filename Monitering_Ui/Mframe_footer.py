from PyQt6.QtWidgets import QFrame, QLabel, QHBoxLayout
from PyQt6.QtCore import Qt


class FrameFooter(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setFixedHeight(32)
        self.setStyleSheet("""
            background-color: rgb(15,23,42);
            border-top: 1px solid #1E293B;
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 0, 12, 0)

        # Left text
        self.label_left = QLabel("VLBI Real-Time Monitoring")
        self.label_left.setStyleSheet(
            "color:#94a3b8; font-size:9pt;"
        )

        # Right text
        self.label_right = QLabel("© 2025 우주측지관측센터")
        self.label_right.setStyleSheet(
            "color:#64748b; font-size:9pt;"
        )
        self.label_right.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        layout.addWidget(self.label_left)
        layout.addStretch()
        layout.addWidget(self.label_right)
