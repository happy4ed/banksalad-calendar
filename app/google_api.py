import base64
import binascii
import io
import json
import logging
from datetime import timedelta
from typing import Optional

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

import config

log = logging.getLogger("google")

SPREADSHEET_MIMES = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
    "application/vnd.google-apps.spreadsheet",
    # 뱅크샐러드는 비밀번호가 걸린 zip으로 보냅니다.
    "application/zip",
    "application/x-zip-compressed",
    "application/octet-stream",
)


def _credentials():
    if config.SERVICE_ACCOUNT_JSON_B64:
        blob = config.SERVICE_ACCOUNT_JSON_B64
        try:
            decoded = base64.b64decode(blob, validate=True)
        except (binascii.Error, ValueError) as err:
            raise SystemExit(
                "[auth] SERVICE_ACCOUNT_JSON_B64 가 올바른 base64가 아닙니다. "
                "줄바꿈이나 따옴표가 섞이지 않았는지 확인해 주세요."
            ) from err
        try:
            info = json.loads(decoded)
        except json.JSONDecodeError as err:
            raise SystemExit("[auth] 디코딩 결과가 JSON이 아닙니다.") from err
        log.info("서비스 계정: %s", info.get("client_email", "?"))
        return service_account.Credentials.from_service_account_info(
            info, scopes=config.SCOPES
        )
    return service_account.Credentials.from_service_account_file(
        config.SERVICE_ACCOUNT_FILE, scopes=config.SCOPES
    )


class Drive:
    def __init__(self):
        self.svc = build("drive", "v3", credentials=_credentials(), cache_discovery=False)
        self.inbox_id = config.INBOX_FOLDER_ID or None
        self.done_id = config.DONE_FOLDER_ID or None

    def _find_folder(self, name: str, parent: Optional[str] = None) -> Optional[str]:
        query = (
            "mimeType='application/vnd.google-apps.folder' and trashed=false "
            f"and name='{name}'"
        )
        if parent:
            query += f" and '{parent}' in parents"
        res = self.svc.files().list(
            q=query, fields="files(id,name)", pageSize=10,
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
        files = res.get("files", [])
        return files[0]["id"] if files else None

    def resolve_folders(self) -> None:
        if self.inbox_id and self.done_id:
            return
        root = self._find_folder(config.ROOT_FOLDER_NAME)
        if not root:
            raise SystemExit(
                f"[drive] '{config.ROOT_FOLDER_NAME}' 폴더를 찾을 수 없습니다. "
                "서비스 계정에 폴더를 공유했는지 확인해 주세요."
            )
        self.inbox_id = self.inbox_id or self._find_folder(config.INBOX_FOLDER_NAME, root)
        self.done_id = self.done_id or self._find_folder(config.DONE_FOLDER_NAME, root)
        if not self.inbox_id or not self.done_id:
            raise SystemExit(
                f"[drive] '{config.ROOT_FOLDER_NAME}' 안에 "
                f"'{config.INBOX_FOLDER_NAME}' / '{config.DONE_FOLDER_NAME}' 폴더가 필요합니다."
            )
        log.info("폴더 확인: inbox=%s done=%s", self.inbox_id, self.done_id)

    def list_inbox(self) -> list[dict]:
        mimes = " or ".join(f"mimeType='{m}'" for m in SPREADSHEET_MIMES)
        res = self.svc.files().list(
            q=f"'{self.inbox_id}' in parents and trashed=false and ({mimes})",
            fields="files(id,name,mimeType,createdTime)",
            orderBy="createdTime",
            pageSize=50,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        return res.get("files", [])

    def download(self, file: dict) -> bytes:
        if file["mimeType"] == "application/vnd.google-apps.spreadsheet":
            request = self.svc.files().export_media(
                fileId=file["id"],
                mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            request = self.svc.files().get_media(fileId=file["id"], supportsAllDrives=True)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return buf.getvalue()

    def move_to_done(self, file_id: str) -> None:
        meta = self.svc.files().get(
            fileId=file_id, fields="parents", supportsAllDrives=True
        ).execute()
        parents = ",".join(meta.get("parents", []))
        self.svc.files().update(
            fileId=file_id,
            addParents=self.done_id,
            removeParents=parents,
            fields="id,parents",
            supportsAllDrives=True,
        ).execute()


class Calendar:
    def __init__(self):
        self.svc = build("calendar", "v3", credentials=_credentials(), cache_discovery=False)

    def insert(self, txn) -> str:
        """Returns 'created', 'duplicate', or 'failed'."""
        calendar_id = (
            config.CALENDAR_INCOME if txn.kind == "income" else config.CALENDAR_EXPENSE
        )

        if txn.at:
            start = f"{txn.day.isoformat()}T{txn.at.strftime('%H:%M:%S')}"
            end_dt = (
                txn.at.hour * 60 + txn.at.minute + config.EVENT_DURATION_MIN
            )
            end_h, end_m = divmod(end_dt, 60)
            if end_h > 23:
                end_h, end_m = 23, 59
            when = {
                "start": {"dateTime": start, "timeZone": config.TIMEZONE},
                "end": {
                    "dateTime": f"{txn.day.isoformat()}T{end_h:02d}:{end_m:02d}:00",
                    "timeZone": config.TIMEZONE,
                },
            }
        else:
            when = {
                "start": {"date": txn.day.isoformat()},
                "end": {"date": (txn.day + timedelta(days=1)).isoformat()},
            }

        body = {
            "id": txn.event_id,
            "summary": txn.title,
            "description": txn.description,
            "transparency": "transparent",
            "reminders": {"useDefault": False},
            "extendedProperties": {"private": {"source": "banksalad"}},
            **when,
        }

        if config.DRY_RUN:
            log.info("[DRY_RUN] %s | %s", txn.day, txn.title)
            return "created"

        try:
            self.svc.events().insert(calendarId=calendar_id, body=body).execute()
            return "created"
        except HttpError as err:
            if err.resp.status == 409:
                return "duplicate"
            log.warning("등록 실패 (%s): %s | %s", err.resp.status, txn.day, txn.title)
            return "failed"
