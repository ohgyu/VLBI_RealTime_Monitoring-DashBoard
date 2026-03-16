# DashBoard_Ui/frame_left.py
from PyQt6.QtWidgets import (
    QFrame, QWidget, QVBoxLayout, QPushButton, QListWidget, QLabel,
    QScrollArea, QSizePolicy, QApplication
)
from PyQt6.QtCore import pyqtSignal, Qt
from functools import partial


class FrameLeft(QScrollArea):
    item_selected = pyqtSignal(str, str, bool)

    def __init__(self, parent=None):
        super().__init__(parent)

        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)
        FIXED_W = 340
        self.setMinimumWidth(FIXED_W)
        self.setMaximumWidth(FIXED_W)

        # 스크롤 설정
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        self.setStyleSheet("background-color: rgb(15,23,42); border-radius:15px;")

        # 콘텐츠 프레임
        content = QWidget()
        content.setStyleSheet("background-color: rgb(15,23,42);")
        self.setWidget(content)

        layout = QVBoxLayout(content)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)

        self.active_child = None

        # 제목
        title_label = QLabel("Select Parameter")
        title_label.setStyleSheet(
            "color: rgb(110,227,247); font-weight:bold; font-size:20pt;"
        )
        layout.addWidget(title_label)

        # 메뉴 항목
        self.menu_lists = {
            "2GHz Receiver Status Monitor": [
                "Normal Temperature RF", "Normal Temperature Load",
                "LNA Monitor LHCP Vd", "LNA Monitor LHCP Id",
                "LNA Monitor RHCP Vd", "LNA Monitor RHCP Id",
                "Cryogenic Temperature Cold", "Cryogenic Temperature Shield Box",
                "Pressure Sensor CH1",
                "RF out Power RHCP", "RF out Power LHCP"
            ],
            "8GHz Receiver Status Monitor": [
                "Normal Temperature RF", "Normal Temperature Load",
                "LNA Monitor LHCP Vg1", "LNA Monitor LHCP Vg2",
                "LNA Monitor LHCP Vd", "LNA Monitor LHCP Id",
                "LNA Monitor RHCP Vg1", "LNA Monitor RHCP Vg2",
                "LNA Monitor RHCP Vd", "LNA Monitor RHCP Id",
                "Cryogenic Temperature Cold", "Cryogenic Temperature Shield Box",
                "Pressure Sensor CH1",
                "RF out Power RHCP", "RF out Power LHCP"
            ],
            "22GHz Receiver Status Monitor": [
                "Normal Temperature RF", "Normal Temperature LO",
                "LNA Monitor LHCP Vg1", "LNA Monitor LHCP Vg2",
                "LNA Monitor LHCP Vd", "LNA Monitor LHCP Id",
                "LNA Monitor RHCP Vg1", "LNA Monitor RHCP Vg2",
                "LNA Monitor RHCP Vd", "LNA Monitor RHCP Id",
                "Cryogenic Temperature Cold", "Cryogenic Temperature Shield Box",
                "Pressure Sensor CH1",
                "RF out Power RHCP", "RF out Power LHCP", "RF out Power LO"
            ],
            "43GHz Receiver Status Monitor": [
                "Normal Temperature RF", "Normal Temperature LO",
                "LNA Monitor LHCP Vg1", "LNA Monitor LHCP Vg2",
                "LNA Monitor LHCP Vd", "LNA Monitor LHCP Id",
                "LNA Monitor RHCP Vg1", "LNA Monitor RHCP Vg2",
                "LNA Monitor RHCP Vd", "LNA Monitor RHCP Id",
                "Cryogenic Temperature Cold", "Cryogenic Temperature Shield Box",
                "Pressure Sensor CH1",
                "RF out Power RHCP", "RF out Power LHCP", "RF out Power LO"
            ],
            "S/X Down Converter": ["S", "X1", "X2"],
            "K Down Converter": ["K1", "K2", "K3", "K4"],
            "Q Down Converter": ["Q1", "Q2", "Q3", "Q4"],
            "Video Converter 1": ["CH1", "CH2", "LEVEL_CH3", "CH4", "CH5", "CH6", "CH7", "CH8"],
            "Video Converter 2": [
                "LEVELL CH9", "LEVELL CH10", "LEVELL CH11", "LEVELL CH12",
                "LEVELL CH13", "LEVELL CH14", "LEVELL CH15", "LEVELL CH16",
                "LEVELU CH9", "LEVELU CH10", "LEVELU CH11", "LEVELU CH12",
                "LEVELU CH13", "LEVELU CH14", "LEVELU CH15", "LEVELU CH16",
            ],
            "IF Selector": [
                "CH1", "CH2", "CH3", "CH4", "CH5", "CH6", "CH7", "CH8",
                "CH9", "CH10", "CH11", "CH12", "CH13", "CH14", "CH15", "CH16"
            ],
        }

        self.lists = {}
        self.buttons = {}

        for title, items in self.menu_lists.items():
            btn = QPushButton(title)
            btn.setStyleSheet(self.default_btn_style())
            self.buttons[title] = btn
            layout.addWidget(btn)

            lst = QListWidget()
            lst.addItems(items)
            lst.setVisible(False)
            lst.setStyleSheet(self.list_style())
            lst.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
            layout.addWidget(lst)
            self.lists[title] = lst

            btn.clicked.connect(partial(self.toggle_parent_item, title))
            lst.itemClicked.connect(partial(self.select_child_item, group_name=title))

        layout.addStretch()

    def default_btn_style(self):
        return """
        QPushButton {
            text-align:left;
            background-color:#1E293B;
            color:white;
            padding:8px;
            border-radius:5px;
            font-size:16pt;
        }
        QPushButton:hover { background-color:#334155; }
        """

    def active_btn_style(self):
        return """
        QPushButton {
            text-align:left;
            background-color:#2563EB;
            color:white;
            padding:8px;
            border-radius:5px;
            font-size:16pt;
            font-weight:bold;
        }
        """

    def list_style(self):
        return """
        QListWidget {
            background-color:#1E293B;
            color:white;
            border:1px solid rgb(80,160,200);
            border-radius:5px;
            font-size:12pt;
        }
        QListWidget::item:selected {
            background-color:#2563EB;
            color:white;
            font-weight:bold;
        }
        QListWidget::item:hover {
            background-color:#334155;
        }
        """

    def toggle_parent_item(self, group_name):
        lst = self.lists[group_name]
        lst.setVisible(not lst.isVisible())

        active_parents = [name for name, l in self.lists.items() if l.isVisible()]
        self.update_button_highlight(active_parents)

        # ✅ 잔상 제거용 강제 갱신
        self.widget().layout().activate()
        self.widget().adjustSize()
        QApplication.processEvents()  # 레이아웃/스크롤바 계산 반영
        self.viewport().update()  # 스크롤 viewport 다시 그리기

    def update_button_highlight(self, active_names):
        for name, btn in self.buttons.items():
            if name in active_names:
                btn.setStyleSheet(self.active_btn_style())
            else:
                btn.setStyleSheet(self.default_btn_style())

    def select_child_item(self, item, group_name):

        self.item_selected.emit(group_name, item.text(), False)

    def refresh_child_selection(self):
        """오른쪽 그래프에 실제 표시중인 (parent, child) 를 기준으로 왼쪽 선택 표시 갱신"""

        if not hasattr(self, "frame_center"):
            return

        # 현재 실제 표시중인 그래프 항목들
        # [(parent, child), ...]
        active_items = set(self.frame_center.get_current_selected_items())

        # parent 기준으로 정리
        active_by_parent = {}
        for parent, child in active_items:
            active_by_parent.setdefault(parent, set()).add(child)

        # 왼쪽 메뉴 전체 순회
        for parent, lst in self.lists.items():  # parent = "2GHz 수신기 상태 모니터" ...
            # 리스트 비어 있을 시 패스
            if lst.count() == 0:
                continue

            for i in range(lst.count()):
                item = lst.item(i)
                child_name = item.text()

                if parent in active_by_parent and child_name in active_by_parent[parent]:
                    # 표시되는 그래프에 포함 → 선택된 상태로
                    item.setSelected(True)
                else:
                    item.setSelected(False)

    def clear_all_selection(self):
        """왼쪽 모든 리스트/버튼 선택 해제 + 리스트 접기"""
        for parent, lst in self.lists.items():
            lst.clearSelection()
            lst.setVisible(False)
        # 버튼 스타일도 전부 기본으로
        self.update_button_highlight([])

    def set_frame_center(self, frame_center):
        self.frame_center = frame_center
