import os, sys, json
from pathlib import Path

APP_NAME = "VLBI_Monitoring"
DEFAULT_NAME = "thresholds.default.json"
USER_NAME = "thresholds.json"

def _appdata_dir() -> Path:
    base = Path(os.getenv("APPDATA") or Path.home())
    d = base / APP_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d

def user_threshold_path() -> Path:
    return _appdata_dir() / USER_NAME

def _default_threshold_path() -> Path:
    # PyInstaller 실행: _MEIPASS 내부
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / DEFAULT_NAME
    # 개발 실행: 프로젝트 루트(이 파일 기준)
    return Path(__file__).resolve().parent / DEFAULT_NAME

def ensure_user_thresholds():
    up = user_threshold_path()
    if not up.exists():
        dp = _default_threshold_path()
        up.write_text(dp.read_text(encoding="utf-8"), encoding="utf-8")
    return up

def load_thresholds() -> dict:
    p = ensure_user_thresholds()
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}

def save_thresholds(data: dict):
    p = ensure_user_thresholds()
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")