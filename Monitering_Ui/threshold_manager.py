from thresholds_store import load_thresholds, save_thresholds
from Monitering_Ui.Alarm_log import AlarmLogger
from db_manager import DB_PATH


def _migrate_lowercase_tables(thresholds: dict) -> tuple[dict, bool]:
    mapping = {
        "frontend_2ghz": "Frontend_2ghz",
        "frontend_8ghz": "Frontend_8ghz",
        "frontend_22ghz": "Frontend_22ghz",
        "frontend_43ghz": "Frontend_43ghz",
    }

    changed = False

    for low, up in mapping.items():
        if low not in thresholds:
            continue

        thresholds.setdefault(up, {})
        low_cols = thresholds.get(low, {})
        up_cols = thresholds.get(up, {})

        # 대문자 우선, 대문자에 없으면 소문자 값 복사
        for col, val in low_cols.items():
            if col not in up_cols:
                up_cols[col] = val
                changed = True

        thresholds[up] = up_cols

        # 소문자 키 삭제
        thresholds.pop(low, None)
        changed = True

    return thresholds, changed

class ThresholdManager:
    def __init__(self):
        self.thresholds = load_thresholds() or {}
        self.thresholds, changed = _migrate_lowercase_tables(self.thresholds)
        if changed:
            save_thresholds(self.thresholds)
        self.logger = AlarmLogger(DB_PATH)

    def load(self):
        self.thresholds = load_thresholds() or {}

    def save(self):
        save_thresholds(self.thresholds)

    def get_threshold(self, table, column):
        th = self.thresholds.get(table, {}).get(column)
        if not th:
            return None

        return {
            "하한 주의": th.get("하한 주의"),
            "하한 경고": th.get("하한 경고"),
            "상한 주의": th.get("상한 주의") or th.get("yellow"),
            "상한 경고": th.get("상한 경고") or th.get("red"),
        }

    def set_threshold(self, table, column, ly, lr, uy, ur):
        self.thresholds.setdefault(table, {})
        self.thresholds[table][column] = {
            "하한 주의": ly,
            "하한 경고": lr,
            "상한 주의": uy,
            "상한 경고": ur
        }
        self.save()

