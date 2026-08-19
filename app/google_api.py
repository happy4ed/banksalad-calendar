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

    def restore_done_to_inbox(self) -> int:
        """정리 후 재생성을 위해 done의 파일을 inbox로 되돌립니다."""
        mimes = " or ".join(f"mimeType='{m}'" for m in SPREADSHEET_MIMES)
        res = self.svc.files().list(
            q=f"'{self.done_id}' in parents and trashed=false and ({mimes})",
            fields="files(id,name)",
            pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = res.get("files", [])
        for file in files:
            self.svc.files().update(
                fileId=file["id"],
                addParents=self.inbox_id,
                removeParents=self.done_id,
                fields="id",
                supportsAllDrives=True,
            ).execute()
            log.info("재처리를 위해 inbox로 복귀: %s", file["name"])
        return len(files)

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

    def purge(self, calendar_id: str, label: str) -> int:
        """이 앱이 만든 일정을 지웁니다. 지운 개수를 반환합니다."""
        deleted = 0
        scanned = 0
        page_token = None

        while True:
            params = {
                "calendarId": calendar_id,
                "maxResults": 250,
                "singleEvents": True,
                "showDeleted": False,
                "fields": "nextPageToken,items(id,summary,extendedProperties)",
            }
            if page_token:
                params["pageToken"] = page_token
            if not config.PURGE_ALL:
                params["privateExtendedProperty"] = "source=banksalad"

            page = self.svc.events().list(**params).execute()
            items = page.get("items", [])
            scanned += len(items)

            for event in items:
                if config.DRY_RUN:
                    deleted += 1
                    continue
                try:
                    self.svc.events().delete(
                        calendarId=calendar_id, eventId=event["id"]
                    ).execute()
                    deleted += 1
                except HttpError as err:
                    if err.resp.status in (404, 410):
                        continue  # 이미 지워짐
                    log.warning(
                        "삭제 실패 (%s): %s", err.resp.status, event.get("summary", "?")
                    )

            page_token = page.get("nextPageToken")
            if not page_token:
                break
            if config.DRY_RUN:
                break  # DRY_RUN에서는 삭제하지 않으므로 무한 반복 방지

        prefix = "[DRY_RUN] " if config.DRY_RUN else ""
        log.info("%s%s 정리: %d건 조회 → %d건 삭제", prefix, label, scanned, deleted)
        return deleted

    def purge_all(self) -> None:
        scope = "모든 일정" if config.PURGE_ALL else "이 앱이 만든 일정"
        log.warning("PURGE_ONCE 활성화 — %s을 삭제합니다.", scope)
        total = self.purge(config.CALENDAR_EXPENSE, "가계부-지출")
        total += self.purge(config.CALENDAR_INCOME, "가계부-수입")
        log.warning("정리 완료: 총 %d건", total)
        if not config.DRY_RUN:
            log.warning(
                "compose에서 PURGE_ONCE를 false로 되돌린 뒤 재시작해 주세요. "
                "그대로 두면 재시작할 때마다 일정이 지워집니다."
            )

    def upsert(self, summary) -> str:
        """Returns 'created', 'updated', 'unchanged', or 'failed'."""
        calendar_id = (
            config.CALENDAR_INCOME if summary.kind == "income" else config.CALENDAR_EXPENSE
        )

        body = {
            "id": summary.event_id,
            "summary": summary.title,
            "description": summary.description,
            "transparency": "transparent",
            "reminders": {"useDefault": False},
            "extendedProperties": {"private": {"source": "banksalad"}},
            "start": {"date": summary.day.isoformat()},
            "end": {"date": (summary.day + timedelta(days=1)).isoformat()},
        }

        if config.DRY_RUN:
            log.info("[DRY_RUN] %s | %s", summary.day, summary.title)
            for line in summary.description.splitlines()[:3]:
                log.info("[DRY_RUN]     %s", line)
            return "created"

        try:
            self.svc.events().insert(calendarId=calendar_id, body=body).execute()
            return "created"
        except HttpError as err:
            if err.resp.status != 409:
                log.warning("등록 실패 (%s): %s | %s", err.resp.status, summary.day, summary.title)
                return "failed"

        # 이미 있는 날짜 — 내용이 달라졌을 때만 갱신합니다.
        try:
            existing = self.svc.events().get(
                calendarId=calendar_id, eventId=summary.event_id
            ).execute()
            if (
                existing.get("summary") == body["summary"]
                and existing.get("description") == body["description"]
                and existing.get("status") != "cancelled"
            ):
                return "unchanged"
            self.svc.events().update(
                calendarId=calendar_id, eventId=summary.event_id, body=body
            ).execute()
            return "updated"
        except HttpError as err:
            log.warning("갱신 실패 (%s): %s | %s", err.resp.status, summary.day, summary.title)
            return "failed"
