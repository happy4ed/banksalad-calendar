import os


def _env(key: str, default: str = "") -> str:
    return os.environ.get(key, default).strip()


# 서비스 계정 키는 base64 문자열(권장) 또는 파일 경로로 받습니다.
SERVICE_ACCOUNT_JSON_B64 = _env("SERVICE_ACCOUNT_JSON_B64")
SERVICE_ACCOUNT_FILE = _env("SERVICE_ACCOUNT_FILE")

EXCEL_PASSWORD = _env("EXCEL_PASSWORD")

CALENDAR_EXPENSE = _env("CALENDAR_EXPENSE")
CALENDAR_INCOME = _env("CALENDAR_INCOME")

ROOT_FOLDER_NAME = _env("ROOT_FOLDER_NAME", "가계부")
INBOX_FOLDER_NAME = _env("INBOX_FOLDER_NAME", "inbox")
DONE_FOLDER_NAME = _env("DONE_FOLDER_NAME", "done")

INBOX_FOLDER_ID = _env("INBOX_FOLDER_ID")
DONE_FOLDER_ID = _env("DONE_FOLDER_ID")

POLL_INTERVAL_SEC = int(_env("POLL_INTERVAL_SEC", "300"))

TIMEZONE = _env("TIMEZONE", "Asia/Seoul")

# Skip transactions older than this many days. 0 disables the filter.
MAX_AGE_DAYS = int(_env("MAX_AGE_DAYS", "0"))

# Treat 이체 (transfers between own accounts) as noise by default.
INCLUDE_TRANSFERS = _env("INCLUDE_TRANSFERS", "false").lower() == "true"

# Minimum amount to register. Filters out 0-won rows.
MIN_AMOUNT = int(_env("MIN_AMOUNT", "1"))

# Duration in minutes for timed events.
EVENT_DURATION_MIN = int(_env("EVENT_DURATION_MIN", "15"))

DRY_RUN = _env("DRY_RUN", "false").lower() == "true"

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/calendar",
]


def validate() -> None:
    missing = []
    if not CALENDAR_EXPENSE:
        missing.append("CALENDAR_EXPENSE")
    if not CALENDAR_INCOME:
        missing.append("CALENDAR_INCOME")
    if not SERVICE_ACCOUNT_JSON_B64:
        if not SERVICE_ACCOUNT_FILE:
            missing.append("SERVICE_ACCOUNT_JSON_B64")
        elif not os.path.exists(SERVICE_ACCOUNT_FILE):
            missing.append(f"SERVICE_ACCOUNT_FILE (경로에 파일 없음: {SERVICE_ACCOUNT_FILE})")
    if missing:
        raise SystemExit("[config] 설정 누락: " + ", ".join(missing))
