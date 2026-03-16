# Monitering_Ui/Mframe_left.py
from PyQt6.QtWidgets import (
    QScrollArea, QWidget, QVBoxLayout, QPushButton,
    QLabel, QFrame, QHBoxLayout, QGridLayout
)
from PyQt6.QtMultimedia import QSoundEffect
from PyQt6.QtCore import QUrl, Qt, pyqtSignal
from Monitering_Ui.threshold_manager import ThresholdManager
from Monitering_Ui.Alarm_log import AlarmLogger
from Monitering_Ui.Email_notifier import EmailNotifier
from db_manager import get_connection, DB_PATH  # Import DB_PATH
import os
from datetime import datetime
from pathlib import Path
import sys


def resource_path(rel_path: str) -> str:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS) / "Monitering_Ui"
    else:
        base = Path(__file__).resolve().parent  # Monitering_Ui 폴더
    return str(base / rel_path)

class MFrameLeft(QScrollArea):

    device_selected = pyqtSignal(str)

    DEVICE_TABLE_MAP = {
        "2GHz Receiver": "Frontend_2ghz",
        "8GHz Receiver": "Frontend_8ghz",
        "22GHz Receiver": "Frontend_22ghz",
        "43GHz Receiver": "Frontend_43ghz",
        "S/X Down Converter": "SXDownConverter",
        "K Down Converter": "KDownConverter",
        "Q Down Converter": "QDownConverter",
        "Video Converter 1": None,
        "Video Converter 2": "VideoConverter2",
        "IF Selector": "IFSelector",
    }

    # 여러 줄을 합쳐야 하는 테이블에 대한 row 개수 정의
    ROW_MERGE_COUNT = {
        "SXDownConverter": 3,
        "KDownConverter": 3,
        "QDownConverter": 4,
        "IFSelector": 3,
        "VideoConverter2": 5,
    }

    def __init__(self, parent=None):
        super().__init__(parent)

        # --- Alarm logger ---
        # Use DB_PATH from db_manager
        self.alarm_logger = AlarmLogger(DB_PATH)

        # --- Email Notifier Setup ---
        EMAIL_SENDER = "vlbi.alarm@gmail.com"
        EMAIL_PASSWORD = "djbv myfj qeuo yvxy"  # App Password
        EMAIL_RECEIVER = "lan.ttn124@gmail.com"

        self.notifier = EmailNotifier(EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_RECEIVER)

        # 이메일 전송 상태 추적용 사전: {(device_name, column_name): True/False}
        self.alert_status = {}

        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setStyleSheet("background-color:#0F172A; border-radius:10px;")

        self.content = QWidget()
        self.setWidget(self.content)

        self.main_layout = QVBoxLayout(self.content)
        self.main_layout.setContentsMargins(10, 10, 10, 10)
        self.main_layout.setSpacing(12)

        self.device_widgets = {}

        title = QLabel("REAL-TIME DEVICE STATUS")
        title.setStyleSheet("color:#38BDF8; font-size:20pt; font-weight:bold;")
        self.main_layout.addWidget(title)

        for device in self.DEVICE_TABLE_MAP.keys():
            self._create_device_row(device)

        self.main_layout.addStretch()
        # ===================================================
        # Alarm Logger & # Threshold Manager
        # ===================================================
        self.thresholds = ThresholdManager()

        # ===================================================
        # 안정 버전: QSoundEffect 절대 크래시 방지
        # ===================================================
        self.alarm = QSoundEffect()
        wav_path = resource_path("alarm.wav")

        if os.path.exists(wav_path):
            self.alarm.setSource(QUrl.fromLocalFile(wav_path))
            self.sound_enabled = True
        else:
            print("[경고] alarm.wav 파일 없음 → 알람 비활성화 (크래시 방지)")
            self.sound_enabled = False

        self.alarm.setVolume(0.9)
        self.last_alarm = 0

        self.alarm_is_active = False

        # ★ 음소거 해제 후 기존 경보 무시용
        self.ignore_existing_errors = False
        # ★ 이전 사이클의 에러 집합
        self.prev_error_set = set()
        # ★ 임계값 변경 중 여부
        self.threshold_editing = False
        # ★ 기준 상태 (임계값 변경 시 스냅샷)
        self.baseline_errors = set()

    # ---------------------------------------------------------
    def _create_device_row(self, device_name: str):
        btn = QPushButton(device_name)
        btn.setCheckable(True)
        btn.setStyleSheet(self._btn_style_default())
        self.main_layout.addWidget(btn)

        panel = QWidget()
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(20, 5, 10, 5)
        panel_layout.setSpacing(6)
        panel.setVisible(False)
        self.main_layout.addWidget(panel)

        btn.clicked.connect(lambda checked, name=device_name: self._toggle_panel(name))

        self.device_widgets[device_name] = {
            "button": btn,
            "panel": panel,
        }

    # ---------------------------------------------------------
    def _btn_style_default(self):
        return """
        QPushButton {
            text-align:left;
            background-color:#1E293B;
            color:white;
            padding:10px;
            border-radius:8px;
            font-size:16pt;
        }
        QPushButton:hover { background-color:#334155; }
        QPushButton:checked {
            background-color:#2563EB;
            font-weight:bold;
        }
        """

    # ---------------------------------------------------------
    def _toggle_panel(self, device_name: str):
        info = self.device_widgets.get(device_name)
        if not info:
            return

        btn = info["button"]
        panel = info["panel"]

        if panel.isVisible():
            panel.setVisible(False)
            btn.setChecked(False)
        else:
            panel.setVisible(True)
            btn.setChecked(True)
            self._reload_panel(device_name)

        self.device_selected.emit(device_name)

    # ---------------------------------------------------------
    def refresh_expanded(self):
        for device_name, info in self.device_widgets.items():
            if info["panel"].isVisible():
                self._reload_panel(device_name)

    # ---------------------------------------------------------
    def _fetch_latest_row_with_merge(self, table: str, conn):
        """
        테이블에서 가장 최신 데이터를 읽되,
        ROW_MERGE_COUNT에 정의된 경우 여러 줄을 합쳐서 하나의 row처럼 반환한다.
        반환값: (col_names, row_tuple 또는 None)
        """
        merge_count = self.ROW_MERGE_COUNT.get(table, 1)
        cur = conn.cursor()

        # 병합 불필요: 기존 방식 그대로
        if merge_count <= 1:
            cur.execute(f"SELECT * FROM {table} ORDER BY datetime DESC LIMIT 1")
            row = cur.fetchone()
            col_names = [d[0] for d in cur.description] if row else []
            return col_names, row

        # 병합 필요한 테이블: 최근 merge_count 줄을 가져와 병합
        cur.execute(
            f"SELECT * FROM {table} ORDER BY datetime DESC LIMIT {merge_count}"
        )
        rows = cur.fetchall()
        if not rows:
            return [], None

        col_names = [d[0] for d in cur.description]

        # 각 컬럼별로 첫 번째 non-NULL 값 선택 (최신 row부터 순서대로)
        merged = [None] * len(col_names)
        for row in rows:
            for idx, val in enumerate(row):
                if merged[idx] is None and val is not None:
                    merged[idx] = val

        # datetime 컬럼이 있다면 가장 최신 row의 값 최소 보장
        try:
            dt_idx = col_names.index("datetime")
            if merged[dt_idx] is None:
                merged[dt_idx] = rows[0][dt_idx]
        except ValueError:
            # datetime 컬럼이 없는 경우는 무시
            pass

        return col_names, tuple(merged)

    # ---------------------------------------------------------
    def _reload_panel(self, device_name: str):

        info = self.device_widgets.get(device_name)
        if not info:
            return

        try:
            self.thresholds.load()
        except Exception as e:
            print("Threshold reload error:", e)

        panel = info["panel"]

        # ★ 핵심: 기존 레이아웃 자체를 완전히 제거
        old_layout = panel.layout()
        if old_layout is not None:
            QWidget().setLayout(old_layout)  # 레이아웃 강제 분리

        # ★ 새 레이아웃 생성 (매번 새로)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(20, 5, 10, 5)
        layout.setSpacing(6)

        table = self.DEVICE_TABLE_MAP.get(device_name)
        if not table:
            layout.addWidget(self._make_label("실시간 데이터 미구성 (테이블 없음)"))
            return

        try:
            conn = get_connection(readonly=True)
            col_names, row = self._fetch_latest_row_with_merge(table, conn)
            if not row:
                layout.addWidget(self._make_label("데이터가 없습니다."))
                return
        except Exception as e:
            layout.addWidget(self._make_label(f"DB 오류: {e}"))
            return
        finally:
            try:
                conn.close()
            except:
                pass

        # -------------------------
        # 임계값 검사 (항상 초기화 + 배타 처리)
        # -------------------------
        level_map = {}

        for col, raw_val in zip(col_names, row):

            level = None  # ★ 매번 초기화

            try:
                val = float(raw_val)
            except:
                level_map[col] = None
                continue

            th = self.thresholds.get_threshold(table, col)
            if not th:
                level_map[col] = None
                continue

            ly = th.get("하한 주의")
            lr = th.get("하한 경고")
            uy = th.get("상한 주의")
            ur = th.get("상한 경고")

            if lr is not None and val <= lr:
                level = "red"
            elif ly is not None and val <= ly:
                level = "yellow"
            elif ur is not None and val >= ur:
                level = "red"
            elif uy is not None and val >= uy:
                level = "yellow"

            level_map[col] = level

        # ===================================================
        # UI 표시 (3열 Grid — 완전 고정)
        # ===================================================
        grid = QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(16)  # ★ 핵심
        grid.setContentsMargins(10, 5, 10, 5)

        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(2, 1)

        layout.addLayout(grid)

        row_idx = 0
        col_idx = 0

        for col, val in zip(col_names, row):
            card = self._make_value_card(col, val, level_map.get(col))

            wrapper = QWidget()
            wrapper.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
            wrapper.setStyleSheet("background: transparent;")

            wrapper_layout = QVBoxLayout(wrapper)
            wrapper_layout.setContentsMargins(0, 0, 0, 0)
            wrapper_layout.setSpacing(0)
            wrapper_layout.addWidget(card)

            grid.addWidget(wrapper, row_idx, col_idx)

            col_idx += 1
            if col_idx >= 3:
                col_idx = 0
                row_idx += 1

    # ---------------------------------------------------------
    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            w = item.widget()
            if w:
                w.setParent(None)
        layout.invalidate()

    # ---------------------------------------------------------
    def _make_label(self, text: str):
        lbl = QLabel(text)
        lbl.setStyleSheet("color:#9CA3AF; font-size:12pt;")
        lbl.setWordWrap(True)
        return lbl

    # ---------------------------------------------------------
    def _make_value_card(self, name, value, level=None):

        card = QFrame()

        color_border = "#1E293B"
        color_bg = "#020617"

        if level == "yellow":
            color_border = "#facc15"
            color_bg = "#4a3f05"
        elif level == "red":
            color_border = "#f87171"
            color_bg = "#5c1f1f"

        card.setStyleSheet(f"""
            QFrame {{
                background-color: {color_bg};
                border-radius: 8px;
                border: 2px solid {color_border};
            }}
        """)

        from PyQt6.QtWidgets import QSizePolicy

        card.setMinimumWidth(260)
        card.setMinimumHeight(52)
        card.setMaximumHeight(52)  # ★ 흔들림 방지 핵심

        card.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Fixed
        )

        hl = QHBoxLayout(card)
        hl.setContentsMargins(10, 5, 10, 5)
        hl.setSpacing(8)

        lbl_name = QLabel(str(name))
        lbl_name.setStyleSheet("color:#9CA3AF; font-size:12pt;")

        lbl_val = QLabel("" if value is None else str(value))
        lbl_val.setStyleSheet("color:white; font-size:14pt; font-weight:bold;")
        lbl_val.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        lbl_val.setMinimumWidth(90)

        hl.addWidget(lbl_name, 3)
        hl.addWidget(lbl_val, 2)

        return card

# ---------------------------------------------------------
# Update Thresholds & Alarm Logic
# ---------------------------------------------------------

    #---------------경고 발생 시 경고 로그 생성---------------
    def update_all_thresholds(self):

        if self.threshold_editing:
            return

        current_error_set = set()

        upper_cautions = []
        upper_warnings = []
        lower_cautions = []
        lower_warnings = []

        # ------- 모든 장치를 스캔 -------
        for device, table in self.DEVICE_TABLE_MAP.items():
            if not table:
                continue

            try:
                conn = get_connection(readonly=True)
                col_names, row = self._fetch_latest_row_with_merge(table, conn)
                if not row:
                    continue
            finally:
                try:
                    conn.close()
                except:
                    pass

            # 기기의 날짜 및 시간을 가져옴
            try:
                dt_idx = col_names.index("datetime")
                device_time = row[dt_idx]
            except:
                device_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # 2. 각 열을 임계값과 비교하여 확인
            for col, raw_val in zip(col_names, row):
                try:
                    val = float(raw_val)
                except:
                    continue

                th = self.thresholds.get_threshold(table, col)
                if not th:
                    continue

                ly = th.get("하한 주의")
                lr = th.get("하한 경고")
                uy = th.get("상한 주의")
                ur = th.get("상한 경고")

                alarm_level = None
                threshold_type = None
                threshold_value = None

                # 메일/로그에서 쓸 메시지 기본값(안전장치)
                msg = None

                # -------- LOWER --------
                if lr is not None and val <= lr:
                    alarm_level = "WARNING"
                    threshold_type = "lower_warning"
                    threshold_value = lr
                    lower_warnings.append(f"{device} - {col}: {val} (Limit: {lr})")

                elif ly is not None and val <= ly:
                    alarm_level = "CAUTION"
                    threshold_type = "lower_caution"
                    threshold_value = ly
                    lower_cautions.append(f"{device} - {col}: {val} (Limit: {ly})")

                # -------- UPPER --------
                elif ur is not None and val >= ur:
                    alarm_level = "WARNING"
                    threshold_type = "upper_warning"
                    threshold_value = ur
                    upper_warnings.append(f"{device} - {col}: {val} (Limit: {ur})")

                elif uy is not None and val >= uy:
                    alarm_level = "CAUTION"
                    threshold_type = "upper_caution"
                    threshold_value = uy
                    upper_cautions.append(f"{device} - {col}: {val} (Limit: {uy})")

                # =========================================================
                # LOGIC DATABASE LOGGING & ERROR SET
                # =========================================================
                if alarm_level:
                    error_key = (device, col, alarm_level, threshold_type)
                    current_error_set.add(error_key)

                    # 메시지 생성(메일/로그 공용)
                    msg = f"{col}: {val} exceed {threshold_type} limit {threshold_value}"

                    # 이전 스캔과 비교하여 "새로 생긴 오류"만 DB에 기록
                    if error_key not in self.prev_error_set:
                        self.alarm_logger.log(
                            log_time=device_time,
                            device=device,
                            level=alarm_level,
                            message=msg
                        )

                # =========================================================
                # 이메일 전송 로직 (오류가 감지될 때 한 번만 실행)
                # =========================================================
                alert_key = (device, col)

                if alarm_level:
                    if not self.alert_status.get(alert_key, False):
                        print(f"{device} - {col}: {alarm_level} {threshold_type}")

                        if alarm_level == "WARNING":
                            self.notifier.send_warning(device_time, device, alarm_level, msg)
                        else:
                            self.notifier.send_caution(device_time, device, alarm_level, msg)

                        self.alert_status[alert_key] = True
                else:
                    if self.alert_status.get(alert_key, False) is True:
                        print(f"{device} - {col} is normal again")
                        self.alert_status[alert_key] = False

        # =========================================================
        # ✅ 여기부터는 "한 틱에 딱 1번만" 실행되게 루프 밖으로 뺌
        # =========================================================

        # 3. Update states for next cycle (새로 생긴 오류만 계산)
        new_errors = current_error_set - self.prev_error_set
        self.prev_error_set = current_error_set.copy()

        # 4. Update Summary UI (한 번만 갱신)
        if hasattr(self, "summary") and self.summary:
            self.summary.update_alerts(
                upper_cautions, upper_warnings,
                lower_cautions, lower_warnings
            )

        # 5. Handle Sound Alarm (한 번만 판단/재생)
        if self.sound_enabled and new_errors:
            if not self.alarm_is_active:
                self.alarm.play()
                self.alarm_is_active = True

        # 오류가 완전히 없어지면 stop (한 번만)
        if not current_error_set and self.alarm_is_active:
            self.alarm.stop()
            self.alarm_is_active = False