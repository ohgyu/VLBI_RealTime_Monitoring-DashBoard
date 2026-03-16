from typing import TYPE_CHECKING
from PyQt6.QtWidgets import (
    QFrame, QLabel, QPushButton, QHBoxLayout, QVBoxLayout, QListWidgetItem,
    QDialog, QListWidget, QGraphicsOpacityEffect, QMessageBox
)
from PyQt6.QtCore import Qt, QPropertyAnimation, QEasingCurve, QSize
from PyQt6.QtGui import QPixmap, QIcon
from db_manager import DB_PATH, IMAGE_PATH
import os
import sqlite3
import sys
import subprocess
import re
from collections import Counter
from pathlib import Path

# MonitoringWindow 타입 힌트
if TYPE_CHECKING:
    from MoniteringMain import MonitoringWindow

from Monitering_Ui.threshold_dialog import ThresholdDialog

ICON_PATHS = {
    "1": (os.path.join(IMAGE_PATH, "power_on.png")), #ON
    "0": (os.path.join(IMAGE_PATH, "power_off.png")),  #OFF
    "ERR": (os.path.join(IMAGE_PATH, "power_error.png")) #ERROR
}
ICON_MANUAL = (os.path.join(IMAGE_PATH, "manual.png"))
# ======================================================
# MiniCard (경고/주의 개별 카드)
# ======================================================
class MiniCard(QFrame):
    def __init__(self, name, color, parent=None):
        super().__init__(parent)

        # 전체 박스만 외곽선 있음
        self.setStyleSheet(f"""
            QFrame {{
                background-color: #0F172A;
                border-radius: 8px;
                border: 2px solid {color};
            }}
            QLabel {{
                background: transparent;
                color:white;
                font-size:16pt;
                font-weight:bold;
                border: none;
            }}
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 8, 14, 8)
        layout.setSpacing(10)

        self.label_name = QLabel(name)
        self.label_value = QLabel("0")

        layout.addWidget(self.label_name)
        layout.addWidget(self.label_value)

        self.setCursor(Qt.CursorShape.PointingHandCursor)

        # 깜빡임 효과 그대로 유지
        self.effect = QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self.effect)

        self.anim = QPropertyAnimation(self.effect, b"opacity")
        self.anim.setDuration(600)
        self.anim.setStartValue(0.3)
        self.anim.setEndValue(1.0)
        self.anim.setEasingCurve(QEasingCurve.Type.InOutQuad)
        self.anim.setLoopCount(-1)

    def set_count(self, c):
        self.label_value.setText(f"{c}")

        if c >= 1:
            self.anim.start()
        else:
            self.anim.stop()
            self.effect.setOpacity(1.0)


# ======================================================
# GroupCard (상한/하한 박스)
# ======================================================
class GroupCard(QFrame):
    def __init__(self, title: str, parent=None):
        super().__init__(parent)

        self.setStyleSheet("""
            QFrame {
                background-color:transparent;
                border-radius:14px;
                border:2px solid #2563EB;
            }
            QLabel {
                border:none;
                color:white;
                font-size:15pt;
                font-weight:bold;
            }
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 12, 15, 12)
        layout.setSpacing(10)

        # 제목
        self.title_label = QLabel(title)
        layout.addWidget(self.title_label)

        # 경고/주의 카드
        self.card_red = MiniCard("Warning", "#f87171")
        self.card_yellow = MiniCard("Caution", "#facc15")

        layout.addWidget(self.card_red)
        layout.addWidget(self.card_yellow)

        layout.addStretch()

    def update(self, red_count: int, yellow_count: int):
        self.card_red.set_count(red_count)
        self.card_yellow.set_count(yellow_count)

# ======================================================
# MiniCard (CalChopper-PCal-Flat Mirror))
# ======================================================

class StatusMiniCard(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(50, 50)
        self.setStyleSheet("background: transparent; border: none;")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.icon_label = QLabel()
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.icon_label)

    def set_status(self, status_value):
        # Convert status_value to string to handle both int and string inputs
        status_key = str(status_value) if status_value in [0, 1] else "ERR"
        path = ICON_PATHS.get(status_key, ICON_PATHS["ERR"])

        if not os.path.exists(path):
            print(f"Error: Can not find file at {path}")
            self.icon_label.setText("⚠️")
            return

        pix = QPixmap(path)
        if not pix.isNull():
            self.icon_label.setPixmap(
                pix.scaled(50, 50, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            )
        else:
            self.icon_label.setText("❓")
# ======================================================
# 장비 Groupcard (CalChopper-PCal-Flat Mirror))
# ======================================================
class DeviceGroupCard(QFrame):
    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QFrame {
                background-color: #0F172A;
                border-radius: 10px;
                border: 1px solid #f2f2f2;
            }
            QLabel {
                border: none;
                color: white;
                font-size: 15pt;
                font-weight: bold;
            }
        """)
        self.setFixedWidth(170)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(2)

        self.title_label = QLabel(title)
        self.status_icon_widget = StatusMiniCard()

        # ✅ 밴드 라벨 추가
        self.band_label = QLabel("")
        self.band_label.setStyleSheet("""
            QLabel {
                border: none;
                color: rgba(255,255,255,180);
                font-size: 10pt;
                font-weight: 700;
            }
        """)

        # ✅ 제목 + 밴드 묶기 (세로)
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)
        left_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        left_layout.addWidget(self.title_label)
        left_layout.addWidget(self.band_label)

        layout.addLayout(left_layout)
        layout.addStretch(1)
        layout.addWidget(self.status_icon_widget)

    def update_val(self, val):
        self.status_icon_widget.set_status(val)

    def set_band(self, text: str):
        self.band_label.setText(text or "")

# ======================================================
# FrameSummary (전체 Summary UI)
# ======================================================
class FrameSummary(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setStyleSheet("background-color:#0F172A; border-radius:10px;")

        # 리스트 데이터
        self.upper_cautions = []
        self.upper_warnings = []
        self.lower_cautions = []
        self.lower_warnings = []

        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 25, 20, 20)
        layout.setSpacing(10)

        #---------------------Cal Chopper, PCal, Flat Mirror 장비 추가 ---------------------
        self.card_pcal = DeviceGroupCard("PCal")
        self.card_calchopper = DeviceGroupCard("Cal-Chop")
        self.card_flatmirror = DeviceGroupCard("Flat Mirror")
        self.card_flatmirror.set_band("2/8GHz")


        layout.addWidget(self.card_pcal)
        layout.addWidget(self.card_calchopper)
        layout.addWidget(self.card_flatmirror)

        layout.addStretch()

        # --------------------- 상한/하한 박스 ---------------------
        self.card_upper = GroupCard("Upper", self)
        self.card_lower = GroupCard("Lower", self)

        layout.addWidget(self.card_upper)
        layout.addWidget(self.card_lower)

        # --------------------- 음소거 버튼 ---------------------
        self.btn_mute = QPushButton("🔊")
        self.btn_mute.setStyleSheet("""
            QPushButton {
                background-color:#2563EB;
                color:white;
                border-radius:10px;
                font-size:22pt;
            }
        """)
        self.btn_mute.clicked.connect(self.toggle_mute)
        layout.addWidget(self.btn_mute)

        # --------------------- 임계값 설정 버튼 ---------------------
        self.btn_setting = QPushButton("Threshold Setting")
        self.btn_setting.setStyleSheet("""
            QPushButton {
                background-color:#2563EB;
                color:white;
                padding:10px 20px;
                border-radius:8px;
                font-size:12pt;
                font-weight:bold;
            }
            QPushButton:hover { background-color:#1E40AF; }
        """)
        self.btn_setting.clicked.connect(self.open_threshold_dialog)
        layout.addWidget(self.btn_setting)

        # --------------------- 클릭 이벤트 연결 ---------------------
        self.card_upper.card_red.mousePressEvent = lambda e: self.show_list("Upper Threshold Warning", self.upper_warnings)
        self.card_upper.card_yellow.mousePressEvent = lambda e: self.show_list("Upper Threshold Caution", self.upper_cautions)
        self.card_lower.card_red.mousePressEvent = lambda e: self.show_list("Lower Threshold Warning", self.lower_warnings)
        self.card_lower.card_yellow.mousePressEvent = lambda e: self.show_list("Lower Threshold Caution", self.lower_cautions)

        # --------------------- 알람 이력 확인 ---------------------
        self.btn_history = QPushButton("Alarm History")
        self.btn_history.setStyleSheet("""
            QPushButton {
                background-color:#2563EB;
                color:white;
                padding:10px 20px;
                border-radius:8px;
                font-size:12pt;
                font-weight:bold;
            }
            QPushButton:hover { background-color:#1E40AF; }
        """)
        self.btn_history.clicked.connect(self.open_alarm_history)
        layout.addWidget(self.btn_history)

        # -------------- Manual Open 버튼 ---------------------
        self.btn_manual = QPushButton()
        self.btn_manual.setIcon(QIcon(ICON_MANUAL))
        self.btn_manual.setIconSize(QSize(30, 30))
        self.btn_manual.setToolTip("Manual")
        self.btn_manual.setStyleSheet("""
            QPushButton {
                background-color:#2563EB;
                border-radius:10px;
                padding:6px;
            }
        """)
        self.btn_manual.clicked.connect(self.open_manual_pdf)
        layout.addWidget(self.btn_manual)

    # --------------------- 알람 이력 저장 및 확인---------------------
    def open_alarm_history(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Alarm History")
        dlg.resize(800, 500)

        layout = QVBoxLayout(dlg)

        lst = QListWidget()
        lst.setStyleSheet("""
            QListWidget {
                background-color:#0F172A;
                color:white;
                font-size:11pt;
                padding:10px;
            }
            QListWidget::item {
                padding:8px;
                border-bottom:1px solid #334155;
            }
        """)
        layout.addWidget(lst)

        btn_export = QPushButton("Export CSV (Period Stats)")
        btn_export.setStyleSheet("""
                QPushButton {
                    background-color:#2563EB;
                    color:white;
                    padding:10px 20px;
                    border-radius:8px;
                    font-size:12pt;
                    font-weight:bold;
                }
                QPushButton:hover { background-color:#1E40AF; }
            """)
        btn_export.clicked.connect(lambda: self.export_alarm_stats_csv(dlg))
        layout.addWidget(btn_export)

        # -------- Load data from DB --------

        try:
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()

            cur.execute("""
                SELECT datetime, device, alarm_level, message
                FROM Alarm_log
                ORDER BY datetime DESC
                LIMIT 200
            """)

            rows = cur.fetchall()
            conn.close()

            if not rows:
                lst.addItem("No alarm history found.")
            else:
                for dt, device, level, msg in rows:
                    text = f"[{dt}]  [{level}]  {device}\n{msg}"
                    item = QListWidgetItem(text)

                    # ---- Color by alarm level ----
                    if level == "WARNING":
                        item.setForeground(Qt.GlobalColor.red)
                    elif level == "CAUTION":
                        item.setForeground(Qt.GlobalColor.yellow)

                    lst.addItem(item)

        except Exception as e:
            lst.addItem(f"DB Error: {e}")

        dlg.exec()

    def resource_path(self, rel_path: str) -> str:
        """
        PyInstaller(onefile/onedir)에서 리소스 경로를 안전하게 찾기
        - 개발환경: 이 파일(FrameSummary가 있는 파일) 기준 폴더 = Monitering_Ui
        - exe환경 : _MEIPASS/Monitering_Ui
        """
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            base = Path(sys._MEIPASS) / "Monitering_Ui"
        else:
            base = Path(__file__).resolve().parent  # Monitering_Ui 폴더

        return str(base / rel_path)

    # ------ Cal Chopper, PCal, Flat Mirror 장비 상태 표시--------------------------------------------
    def update_device_status(self, status: dict):
        """이 함수는 이제 스레드로부터 딕셔너리를 받아 UI에 표시하는 역할만 함."""
        if not status:
            # 딕셔너리가 비어있으면 ERR을 표시
            if not status:
                for card in [self.card_pcal, self.card_calchopper, self.card_flatmirror]:
                    card.update_val("ERR")
                return

        def quick_validate(val):
            # Handle None or empty values from DB
            if val is None:
                return "ERR"
            try:
                # Convert to float then int to handle strings like "1.0"
                v = int(float(val))
                return v if v in (0, 1) else "ERR"
            except (ValueError, TypeError):
                return "ERR"

        # 스레드에서 전송된 데이터를 기반으로 UI를 업데이트
        self.card_pcal.update_val(quick_validate(status.get("PCAL")))
        self.card_calchopper.update_val(quick_validate(status.get("CalChop")))

        fm = quick_validate(status.get("FlatMirror"))
        self.card_flatmirror.update_val(fm)

        if fm == 1:
            self.card_flatmirror.set_band("22/43GHz")
        elif fm == 0:
            self.card_flatmirror.set_band("2/8GHz")
        else:
            self.card_flatmirror.set_band("")

    def show_list(self, title: str, dataset: list):
        dlg = QDialog(self)
        dlg.setWindowTitle(title)
        dlg.resize(450, 550)

        lst = QListWidget()

        # ★ 선택 색 제거 (클릭해도 색 변화 없음)
        lst.setStyleSheet("""
            QListWidget {
                background-color: #0F172A;
                color: white;
                font-size: 13pt;
                padding: 10px;
            }
            QListWidget::item {
                padding: 12px 8px;
            }
            QListWidget::item:selected {
                background-color: transparent;
                color: white;
                border: none;
            }
        """)

        for x in dataset:
            lst.addItem(x)

        lst.itemClicked.connect(lambda item: self.jump_to_device(item.text()))

        layout = QVBoxLayout(dlg)
        layout.addWidget(lst)

        dlg.exec()

    def jump_to_device(self, text: str):
        """
        예: '22GHz Receiver - RF_LO: -97.0'
        → 장비명 = '22GHz Receiver'
        """

        # -----------------------
        # 1. 텍스트 파싱
        # -----------------------
        try:
            device_name = text.split(" - ")[0].strip()
        except:
            return

        # -----------------------
        # 2. Left 패널 객체 가져오기
        # -----------------------
        win: "MonitoringWindow" = self.window()
        if not hasattr(win, "frame_left"):
            return

        fl = win.frame_left

        # -----------------------
        # 3. 해당 장비 패널 펼치기
        # -----------------------
        if device_name in fl.device_widgets:
            info = fl.device_widgets[device_name]
            btn = info["button"]
            panel = info["panel"]

            # 패널이 닫혀 있다면 열기
            if not panel.isVisible():
                btn.setChecked(True)
                fl._reload_panel(device_name)
                panel.setVisible(True)

            # -----------------------
            # 4. 자동 스크롤
            # -----------------------
            fl.ensureWidgetVisible(panel)

        # -----------------------
        # 팝업 닫기
        # -----------------------

    # --------------------------------------------------
    def update_alerts(self,
                      upper_cautions: list, upper_warnings: list,
                      lower_cautions: list, lower_warnings: list):

        self.upper_cautions = upper_cautions
        self.upper_warnings = upper_warnings
        self.lower_cautions = lower_cautions
        self.lower_warnings = lower_warnings

        self.card_upper.update(len(upper_warnings), len(upper_cautions))
        self.card_lower.update(len(lower_warnings), len(lower_cautions))

    # --------------------------------------------------
    def toggle_mute(self):
        win: "MonitoringWindow" = self.window()

        if not hasattr(win, "frame_left"):
            return

        fl = win.frame_left
        fl.sound_enabled = not fl.sound_enabled

        if fl.sound_enabled:
            self.btn_mute.setText("🔊")

            # ★ 음소거 해제 시 알람 상태 초기화 (중요) ★
            fl.alarm_is_active = False

            fl.last_alarm = 0

            fl.ignore_existing_errors = True

            # Summary UI도 새로 반영
            self.update_alerts(
                self.upper_cautions, self.upper_warnings,
                self.lower_cautions, self.lower_warnings
            )
        else:
            self.btn_mute.setText("🔇")
            fl.alarm.stop()

            # ★ 음소거 ON일 때도 상태 통일해서 끔 ★
            fl.alarm_is_active = False

    # --------------------------------------------------
    def open_threshold_dialog(self):
        win: "MonitoringWindow" = self.window()
        if not hasattr(win, "frame_left"):
            return

        fl = win.frame_left

        # ★ 임계값 설정 시작
        fl.threshold_editing = True

        dlg = ThresholdDialog(parent=self.window())
        dlg.exec()

        # ★ 임계값 설정 종료
        fl.threshold_editing = False

        # ★ 현재 상태를 기준 상태로 재설정 (알람 X)
        fl.baseline_errors = set(fl.prev_error_set)
        fl.alarm_is_active = False

        fl.thresholds.load()
        fl.refresh_expanded()

    def open_manual_pdf(self):
        pdf_path = self.resource_path("mannual.pdf")

        if not os.path.exists(pdf_path):
            QMessageBox.warning(self, "Manual PDF not found", pdf_path)
            return

        try:
            if sys.platform.startswith("win"):
                os.startfile(pdf_path)
            elif sys.platform.startswith("darwin"):
                subprocess.call(["open", pdf_path])
            else:
                subprocess.call(["xdg-open", pdf_path])
        except Exception as e:
            QMessageBox.critical(self, "Failed to open manual PDF", str(e))

    def export_alarm_stats_csv(self, parent_dialog: QDialog):
        import csv
        from PyQt6.QtWidgets import QFileDialog, QDateTimeEdit, QDialogButtonBox
        from PyQt6.QtCore import QDateTime

        # -----------------------------
        # 1) 기간 선택 다이얼로그
        # -----------------------------
        period_dlg = QDialog(parent_dialog)
        period_dlg.setWindowTitle("Select Period")

        period_dlg.setStyleSheet("""
            QDialog {
                background-color: #0F172A;
            }

            QLabel {
                color: white;
                font-size: 11pt;
            }

            /* 날짜/시간 입력창 */
            QDateTimeEdit {
                background-color: #020617;
                color: white;
                border: 1px solid #334155;
                border-radius: 8px;
                padding: 6px 28px 6px 10px;   /* 오른쪽 드롭다운 공간 확보 */
                font-size: 11pt;
                selection-background-color: #2563EB;
                selection-color: white;
            }

            /* 기간 옆(드롭다운 버튼 영역) */
            QDateTimeEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 22px;
                border-left: 1px solid #334155;
                background-color: #1E293B;
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
            }

            /* 기간 옆(화살표 표시) - Qt 기본 화살표가 안 보이면 텍스트 색만 맞춰도 보임 */
            QDateTimeEdit::down-arrow {
                width: 8px;
                height: 8px;
            }

            /* 캘린더 위젯 전체 */
            QCalendarWidget QWidget {
                background-color: #020617;
                color: white;
            }
            QCalendarWidget QToolButton {
                color: white;
                background-color: #1E293B;
                border: 1px solid #334155;
                border-radius: 6px;
                padding: 4px 8px;
            }
            QCalendarWidget QMenu {
                background-color: #020617;
                color: white;
            }

            /* 캘린더 날짜 셀 선택 */
            QCalendarWidget QAbstractItemView {
                selection-background-color: #2563EB;
                selection-color: white;
                background-color: #020617;
                color: white;
            }

            /* ✅ OK/Cancel 버튼 (QDialogButtonBox 내부 버튼) */
            QDialogButtonBox QPushButton {
                background-color: #2563EB;
                color: white;
                padding: 6px 16px;
                border-radius: 8px;
                font-weight: bold;
                border: none;
                min-width: 80px;
            }
            QDialogButtonBox QPushButton:hover {
                background-color: #1E40AF;
            }
            QDialogButtonBox QPushButton:pressed {
                background-color: #1D4ED8;
            }
        """)

        v = QVBoxLayout(period_dlg)

        v.addWidget(QLabel("Select period (From ~ To)"))

        dt_from = QDateTimeEdit()
        dt_to = QDateTimeEdit()
        dt_from.setCalendarPopup(True)
        dt_to.setCalendarPopup(True)

        # 기본값: 최근 24시간
        now = QDateTime.currentDateTime()
        dt_to.setDateTime(now)
        dt_from.setDateTime(now.addDays(-1))

        v.addWidget(QLabel("From:"))
        v.addWidget(dt_from)
        v.addWidget(QLabel("To:"))
        v.addWidget(dt_to)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        v.addWidget(btns)
        btns.accepted.connect(period_dlg.accept)
        btns.rejected.connect(period_dlg.reject)

        if period_dlg.exec() != QDialog.DialogCode.Accepted:
            return

        # ★ 중요: DB에 저장된 datetime 문자열 포맷과 반드시 일치해야 함
        # 권장 포맷: "YYYY-MM-DD HH:MM:SS"
        start_str = dt_from.dateTime().toString("yyyy-MM-dd HH:mm:ss")
        end_str = dt_to.dateTime().toString("yyyy-MM-dd HH:mm:ss")

        # -----------------------------
        # 2) 저장 경로 선택
        # -----------------------------
        path, _ = QFileDialog.getSaveFileName(
            parent_dialog,
            "Save Alarm Stats CSV",
            "alarm_stats.csv",
            "CSV Files (*.csv)"
        )
        if not path:
            return

        # -----------------------------
        # 3) 기간 조건으로 DB 조회
        # -----------------------------
        try:
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()
            cur.execute("""
                        SELECT datetime, device, alarm_level, message
                        FROM Alarm_log
                        WHERE datetime(datetime) >= datetime(?)
                          AND datetime(datetime) <= datetime(?)
                        ORDER BY datetime(datetime) ASC
                        """, (start_str, end_str))
            rows = cur.fetchall()
            conn.close()
        except Exception as e:
            self._show_msg(parent_dialog, "DB Error", str(e), QMessageBox.Icon.Critical)
            return

        # -----------------------------
        # 4) 통계 계산
        # -----------------------------
        total = len(rows)

        # rows: (datetime, device, alarm_level, message)
        by_level = Counter(level for _, _, level, _ in rows)

        by_device = Counter(device for _, device, _, _ in rows)
        by_device_level = Counter((device, level) for _, device, level, _ in rows)

        # ✅ 자식(파라미터) 통계
        by_param = Counter(self.extract_child_param(msg) for _, _, _, msg in rows)
        by_param_level = Counter((self.extract_child_param(msg), level) for _, _, level, msg in rows)

        # ✅ 부모+자식 통계 (가장 중요)
        by_device_param = Counter((device, self.extract_child_param(msg)) for _, device, _, msg in rows)
        by_device_param_level = Counter((device, self.extract_child_param(msg), level) for _, device, level, msg in rows)

        # -----------------------------
        # 5) CSV 저장
        # -----------------------------
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)

                w.writerow(["Alarm Statistics"])
                w.writerow(["From", start_str])
                w.writerow(["To", end_str])
                w.writerow(["Total", total])
                w.writerow([])

                w.writerow(["By Level"])
                w.writerow(["alarm_level", "count"])
                for level, cnt in by_level.most_common():
                    w.writerow([level, cnt])
                w.writerow([])

                w.writerow(["By Device"])
                w.writerow(["device", "count"])
                for dev, cnt in by_device.most_common():
                    w.writerow([dev, cnt])
                w.writerow([])

                # ✅ By Device + Param
                w.writerow(["By Device + Param"])
                w.writerow(["device", "param", "count"])
                for (dev, p), cnt in by_device_param.most_common():
                    w.writerow([dev, p, cnt])
                w.writerow([])

                # ✅ By Device + Param + Level (추천)
                w.writerow(["By Device + Param + Level"])
                w.writerow(["device", "param", "alarm_level", "count"])
                for (dev, p, level), cnt in by_device_param_level.most_common():
                    w.writerow([dev, p, level, cnt])

            self._show_msg(parent_dialog, "Done", f"Saved:\n{path}", QMessageBox.Icon.Information)


        except PermissionError:
            self._show_msg(
                parent_dialog,
                "CSV Save Error",
                "Permission denied.\n\n"
                "1) 엑셀에서 CSV 파일을 닫고 다시 시도\n"
                "2) 다른 파일명으로 저장\n"
                "3) Desktop 대신 Downloads/Documents에 저장",
                QMessageBox.Icon.Critical
            )
            return

        except Exception as e:
            self._show_msg(parent_dialog, "CSV Save Error", str(e), QMessageBox.Icon.Critical)

    def _show_msg(self, parent, title: str, text: str, icon: QMessageBox.Icon):
        msg = QMessageBox(parent)
        msg.setWindowTitle(title)
        msg.setIcon(icon)
        msg.setText(text)

        msg.setStyleSheet("""
            QMessageBox { background-color: #0F172A; }
            QLabel { color: white; font-size: 11pt; }
            QPushButton {
                background-color:#2563EB;
                color:white;
                padding:6px 16px;
                border-radius:8px;
                font-weight:bold;
                min-width:80px;
                border:none;
            }
            QPushButton:hover { background-color:#1E40AF; }
            QPushButton:pressed { background-color:#1D4ED8; }
        """)
        msg.exec()

    @staticmethod
    def extract_child_param(msg: str) -> str:
        if not msg:
            return "UNKNOWN"
        head = msg.split(":", 1)[0].strip()
        return head if head else "UNKNOWN"

    @staticmethod
    def extract_rule(msg: str) -> str:
        if not msg:
            return "UNKNOWN"
        m = re.search(r"exceed\s+(\w+)\s+limit", msg)
        return m.group(1) if m else "UNKNOWN"