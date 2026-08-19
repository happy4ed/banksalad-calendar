import logging
import time as time_mod
from datetime import date, timedelta

import config
import parser
from google_api import Calendar, Drive

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("main")


def should_keep(txn) -> bool:
    if txn.amount < config.MIN_AMOUNT:
        return False
    if txn.kind == "transfer" and not config.INCLUDE_TRANSFERS:
        return False
    if config.MAX_AGE_DAYS > 0:
        if txn.day < date.today() - timedelta(days=config.MAX_AGE_DAYS):
            return False
    return True


# 같은 파일에 대해 같은 오류를 5분마다 반복 출력하지 않기 위한 기록
_reported: set[str] = set()


def process_file(drive: Drive, cal: Calendar, file: dict) -> None:
    log.info("처리 시작: %s", file["name"])
    raw = drive.download(file)

    try:
        txns = list(parser.parse(raw, config.EXCEL_PASSWORD))
    except Exception as err:
        key = f"{file['id']}:{err}"
        if key not in _reported:
            _reported.add(key)
            log.error("파싱 실패 (%s): %s — inbox에 그대로 둡니다.", file["name"], err)
        return

    kept = [t for t in txns if should_keep(t)]
    skipped = len(txns) - len(kept)
    summaries = parser.summarize(kept)

    stats = {"created": 0, "updated": 0, "unchanged": 0, "failed": 0}
    for summary in summaries:
        stats[cal.upsert(summary)] += 1

    log.info(
        "완료: %s — 거래 %d건(제외 %d) → 일정 %d개 · 신규 %d · 갱신 %d · 변동없음 %d · 실패 %d",
        file["name"], len(txns), skipped, len(summaries),
        stats["created"], stats["updated"], stats["unchanged"], stats["failed"],
    )

    _reported.discard(file["id"])

    if stats["failed"] and not config.DRY_RUN:
        log.warning("실패 건이 있어 파일을 inbox에 남겨둡니다: %s", file["name"])
        return

    if config.DRY_RUN:
        log.info("[DRY_RUN] 파일 이동 생략")
        return

    drive.move_to_done(file["id"])
    log.info("done 폴더로 이동: %s", file["name"])


def main() -> None:
    config.validate()
    if config.DRY_RUN:
        log.warning("DRY_RUN 모드 — 캘린더에 실제로 쓰지 않습니다.")

    drive = Drive()
    drive.resolve_folders()
    cal = Calendar()

    if config.PURGE_ONCE:
        cal.purge_all()
        if not config.DRY_RUN:
            restored = drive.restore_done_to_inbox()
            if restored:
                log.info("done → inbox %d개 복귀, 곧 다시 처리됩니다.", restored)

    log.info("감시 시작 (%d초 간격)", config.POLL_INTERVAL_SEC)
    while True:
        try:
            files = drive.list_inbox()
            if files:
                log.info("inbox에 %d개 파일 발견", len(files))
                for file in files:
                    process_file(drive, cal, file)
        except Exception as err:
            log.exception("주기 실행 중 오류: %s", err)
        time_mod.sleep(config.POLL_INTERVAL_SEC)


if __name__ == "__main__":
    main()
