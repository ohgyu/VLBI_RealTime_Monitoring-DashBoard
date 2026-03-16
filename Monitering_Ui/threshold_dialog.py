from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QPushButton,
    QComboBox, QLineEdit, QMessageBox, QFileDialog
)
from PyQt6.QtCore import QTimer
from Monitering_Ui.threshold_manager import ThresholdManager
import csv

# ✅ 통계 대시보드(좌측 메뉴)에서 쓰는 매핑을 그대로 사용
# DashBoard_Ui/frame_center.py 에 TABLE_MAP가 있다고 했으니 이 경로가 맞음
from DashBoard_Ui.frame_center import TABLE_MAP


class ThresholdDialog(QDialog):
    """
    - UI에 보여주는 목록: 통계대시보드 좌측 메뉴와 동일 (TABLE_MAP 기준)
      * 장비 = TABLE_MAP의 key (예: "2GHz Receiver Status Monitor")
      * 항목 = TABLE_MAP[장비]["columns"]의 key (예: "Normal Temperature RF")

    - thresholds 저장 키: 기존 방식 유지 (table, db_column)
      * table = TABLE_MAP[장비]["table"]
      * db_column = TABLE_MAP[장비]["columns"][항목]
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("임계값 설정")
        self.resize(500, 520)

        self.tm = ThresholdManager()

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        self.setStyleSheet("""
            QDialog { background-color:#0F172A; color:white; }
            QLabel { color:white; }
            QComboBox {
                color:white; background-color:#1E293B;
                border:1px solid #334155;
                padding:6px; border-radius:5px;
            }
            QComboBox QAbstractItemView {
                background-color: #1E293B;
                color: white;
                selection-background-color: #2563EB;
                selection-color: white;
                border: 1px solid #334155;
            }
            QLineEdit {
                color: white;
                background-color: #1E293B;
                border: 1px solid #334155;
                padding: 6px;
                border-radius: 5px;
                selection-background-color: #2563EB;
                selection-color: white;
            }
            QPushButton {
                background-color:#2563EB; color:white;
                padding:10px; border-radius:8px; font-weight:bold;
            }
            QPushButton:hover { background-color:#1E40AF; }
        """)

        # -------------------------------
        # 장비 선택 (통계 대시보드 메뉴 그룹)
        # -------------------------------
        layout.addWidget(QLabel("장비 선택"))
        self.combo_device = QComboBox()
        layout.addWidget(self.combo_device)

        # ✅ 통계 대시보드 좌측 메뉴(=TABLE_MAP key) 그대로 넣기
        # (Python 3.7+ dict insertion order 유지됨)
        for device_name in TABLE_MAP.keys():
            self.combo_device.addItem(device_name)

        # -------------------------------
        # 컬럼 선택 (통계 대시보드 메뉴 항목)
        # -------------------------------
        layout.addWidget(QLabel("컬럼 선택"))
        self.combo_column = QComboBox()
        layout.addWidget(self.combo_column)

        layout.addWidget(QLabel("하한 주의"))
        self.input_ly = QLineEdit()
        layout.addWidget(self.input_ly)

        layout.addWidget(QLabel("하한 경고"))
        self.input_lr = QLineEdit()
        layout.addWidget(self.input_lr)

        layout.addWidget(QLabel("상한 주의"))
        self.input_uy = QLineEdit()
        layout.addWidget(self.input_uy)

        layout.addWidget(QLabel("상한 경고"))
        self.input_ur = QLineEdit()
        layout.addWidget(self.input_ur)

        btn_save = QPushButton("저장")
        btn_save.clicked.connect(self.save_threshold)
        layout.addWidget(btn_save)

        # -------------------------------
        # CSV 내보내기 / 불러오기 버튼
        # -------------------------------
        btn_export = QPushButton("CSV 내보내기")
        btn_export.clicked.connect(self.export_csv)
        layout.addWidget(btn_export)

        btn_import = QPushButton("CSV 불러오기")
        btn_import.clicked.connect(self.import_csv)
        layout.addWidget(btn_import)

        QTimer.singleShot(0, self._late_init)

    # ----------------------------------------------------------
    # 초기 UI 로드
    # ----------------------------------------------------------
    def _late_init(self):
        self.combo_device.currentIndexChanged.connect(self._reload_columns)
        self.combo_column.currentIndexChanged.connect(self._load_existing_threshold)
        self._reload_columns()

    # ----------------------------------------------------------
    # (device,label) -> (table, db_column)
    # ----------------------------------------------------------
    def _current_table_and_dbcol(self):
        device = self.combo_device.currentText()
        label = self.combo_column.currentText()

        info = TABLE_MAP.get(device)
        if not info:
            return None, None

        table = info.get("table")
        db_col = info.get("columns", {}).get(label)
        return table, db_col

    # ----------------------------------------------------------
    # 컬럼 목록 로드: TABLE_MAP 기준(통계 메뉴와 동일)
    # ----------------------------------------------------------
    def _reload_columns(self):
        device = self.combo_device.currentText()
        self.combo_column.clear()

        info = TABLE_MAP.get(device, {})
        cols_map = info.get("columns", {})

        # ✅ 통계 대시보드 메뉴 항목(label) 그대로 넣기
        for label in cols_map.keys():
            self.combo_column.addItem(label)

        self._load_existing_threshold()

    # ----------------------------------------------------------
    # 기존 임계값 로드 (table, db_column 기준)
    # ----------------------------------------------------------
    def _load_existing_threshold(self):
        table, db_col = self._current_table_and_dbcol()

        self.input_ly.clear()
        self.input_lr.clear()
        self.input_uy.clear()
        self.input_ur.clear()

        if not table or not db_col:
            return

        th = self.tm.get_threshold(table, db_col)
        if not th:
            return

        if th.get("하한 주의") is not None:
            self.input_ly.setText(str(th["하한 주의"]))
        if th.get("하한 경고") is not None:
            self.input_lr.setText(str(th["하한 경고"]))
        if th.get("상한 주의") is not None:
            self.input_uy.setText(str(th["상한 주의"]))
        if th.get("상한 경고") is not None:
            self.input_ur.setText(str(th["상한 경고"]))

    # ----------------------------------------------------------
    # 저장 버튼
    # ----------------------------------------------------------
    def save_threshold(self):
        table, db_col = self._current_table_and_dbcol()

        if not table or not db_col:
            QMessageBox.warning(self, "오류", "선택한 항목의 DB 매핑(TABLE_MAP)이 없습니다.")
            return

        try:
            ly = float(self.input_ly.text())
            lr = float(self.input_lr.text())
            uy = float(self.input_uy.text())
            ur = float(self.input_ur.text())
        except:
            QMessageBox.warning(self, "입력 오류", "모든 임계값에 숫자를 입력하세요.")
            return

        self.tm.set_threshold(table, db_col, ly, lr, uy, ur)
        QMessageBox.information(self, "저장 완료", "임계값이 저장되었습니다.")
        self.close()

    # ----------------------------------------------------------
    # CSV 내보내기
    # ----------------------------------------------------------
    def export_csv(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "임계값 CSV 내보내기", "thresholds.csv", "CSV Files (*.csv)"
        )
        if not path:
            return

        # ✅ 엑셀 호환: utf-8-sig
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["table", "column", "하한 주의", "하한 경고", "상한 주의", "상한 경고"])

            # ✅ TABLE_MAP 전체 기준으로 "모든 장비/항목"을 다 CSV에 출력 (값 없으면 빈칸)
            for device_name, info in TABLE_MAP.items():
                table = info.get("table")
                cols_map = info.get("columns", {})

                for label, db_col in cols_map.items():
                    th = self.tm.get_threshold(table, db_col) or {}
                    writer.writerow([
                        table,
                        db_col,
                        th.get("하한 주의", ""),
                        th.get("하한 경고", ""),
                        th.get("상한 주의", ""),
                        th.get("상한 경고", ""),
                    ])

        QMessageBox.information(self, "완료", "CSV로 내보내기 완료되었습니다. (전체 항목 포함)")

    # ----------------------------------------------------------
    # CSV 불러오기
    # ----------------------------------------------------------
    def import_csv(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "CSV 선택", "", "CSV Files (*.csv)"
        )
        if not path:
            return

        new_data = {}

        def to_float(x):
            try:
                x = (x or "").strip()
                if x == "":
                    return None
                return float(x)
            except:
                return None

        # ✅ 소문자 테이블명 -> 대문자 테이블명 보정
        table_fix = {
            "frontend_2ghz": "Frontend_2ghz",
            "frontend_8ghz": "Frontend_8ghz",
            "frontend_22ghz": "Frontend_22ghz",
            "frontend_43ghz": "Frontend_43ghz",
        }

        # ✅ (선택) 표시명(label) → db_col 역매핑
        label_to_db = {}
        for _dev, info in TABLE_MAP.items():
            t = info.get("table")
            for label, db_col in (info.get("columns") or {}).items():
                label_to_db[(t, label)] = db_col

        last_err = None

        for enc in ("utf-8-sig", "utf-8", "cp949"):
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    reader = csv.DictReader(f)

                    for row in reader:
                        table = (row.get("table") or "").strip()
                        col = (row.get("column") or "").strip()
                        if not table or not col:
                            continue

                        table = table_fix.get(table, table)

                        # ✅ 사용자가 column에 label(표시명)을 넣은 경우도 자동 보정
                        col = label_to_db.get((table, col), col)

                        new_data.setdefault(table, {})
                        new_data[table][col] = {
                            "하한 주의": to_float(row.get("하한 주의")),
                            "하한 경고": to_float(row.get("하한 경고")),
                            "상한 주의": to_float(row.get("상한 주의")),
                            "상한 경고": to_float(row.get("상한 경고")),
                        }

                # ✅ 기존 thresholds에 "병합" (CSV에 없는 값은 유지)
                merged = self.tm.thresholds or {}
                for table, cols in new_data.items():
                    merged.setdefault(table, {})
                    merged[table].update(cols)

                self.tm.thresholds = merged
                self.tm.save()

                QMessageBox.information(self, "완료", f"CSV에서 임계값이 불러와졌습니다. (encoding={enc})")
                self._reload_columns()
                return

            except Exception as e:
                last_err = e

        QMessageBox.warning(self, "오류", f"CSV 불러오기 실패: {last_err}")