from __future__ import annotations
import argparse
import csv
import json
import logging
import os
import re
import shutil
import time
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from io import BytesIO, StringIO
from pathlib import Path
from typing import Any
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from xml.etree import ElementTree
from zoneinfo import ZoneInfo
import openpyxl
import pandas as pd
from openpyxl.styles import Font, PatternFill

from .constants import CACHE_DIR, CONFIG_DIR, DATA_DIR, LOG_DIR, PLACEHOLDER_API_KEYS, QUARTER_PERIODS, REPORT_DIR, WORK_DIR
from .util import write_json

def resolve_run_date(value: str, timezone_name: str = "Asia/Seoul") -> str:
    if value == "today":
        return datetime.now(ZoneInfo(timezone_name)).strftime("%Y-%m-%d")
    return value

def should_run_scheduled_date(run_date: str, start_date: str, end_date: str) -> tuple[bool, str]:
    current = date.fromisoformat(run_date)
    if current < date.fromisoformat(start_date):
        return False, "before_start_date"
    if current > date.fromisoformat(end_date):
        return False, "after_end_date"
    if current.weekday() >= 5:
        return False, "weekend"
    if run_date in load_market_holidays():
        return False, "market_holiday"
    return True, "market_day"

def load_market_holidays() -> set[str]:
    config_path = CONFIG_DIR / "config.yml"
    if not config_path.exists():
        return set()
    holidays = set()
    in_holidays = False
    for line in config_path.read_text(encoding="utf-8", errors="replace").splitlines():
        stripped = line.strip()
        if stripped == "holidays:":
            in_holidays = True
            continue
        if in_holidays:
            match = re.match(r'-\s+"?(\d{4}-\d{2}-\d{2})"?', stripped)
            if match:
                holidays.add(match.group(1))
                continue
            if stripped and not stripped.startswith("-"):
                break
    return holidays

def quarter_period(quarter: str) -> dict[str, str] | None:
    if quarter in QUARTER_PERIODS:
        return QUARTER_PERIODS[quarter]
    match3 = re.fullmatch(r"(\d{4})Q3", quarter)
    if match3:
        year = int(match3.group(1))
        return {
            "quarter_start": f"{year}-07-01",
            "quarter_end": f"{year}-09-30",
            "prev_quarter": f"{year}Q2",
            "yoy_quarter": f"{year - 1}Q3",
            "current_cumulative": f"{year} 9M",
            "yoy_cumulative": f"{year - 1} 9M",
        }
    match = re.fullmatch(r"(\d{4})Q2", quarter)
    if not match:
        return None
    year = int(match.group(1))
    return {
        "quarter_start": f"{year}-04-01",
        "quarter_end": f"{year}-06-30",
        "prev_quarter": f"{year}Q1",
        "yoy_quarter": f"{year - 1}Q2",
        "current_cumulative": f"{year}H1",
        "yoy_cumulative": f"{year - 1}H1",
    }

def ensure_dirs() -> None:
    for path in [WORK_DIR, REPORT_DIR, DATA_DIR, LOG_DIR, CACHE_DIR]:
        path.mkdir(parents=True, exist_ok=True)

def configure_logging(quarter: str, end_date: str) -> logging.Logger:
    logger = logging.getLogger("quarterly_earnings_filter")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    file_handler = logging.FileHandler(run_log_path(quarter, end_date), encoding="utf-8-sig")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    return logger

def run_log_path(quarter: str, end_date: str) -> Path:
    return LOG_DIR / f"{quarter}_{end_date}.log"

def write_last_run_status(
    args: argparse.Namespace,
    run_started: datetime,
    status: str,
    reason: str,
    collected_disclosures: int,
    passed_count: int,
    report_path: Path | None,
    log_path: Path | None,
) -> None:
    requested_end = getattr(args, "requested_end", None) or getattr(args, "end", "")
    start = getattr(args, "start", "")
    end = getattr(args, "end", "")
    quarter = getattr(args, "quarter", "")
    market = getattr(args, "market", "")
    note = status_note(status, reason, end, start)
    payload = {
        "last_run_at": run_started.strftime("%Y-%m-%d %H:%M:%S"),
        "status": status,
        "reason": reason,
        "start": start,
        "end": end,
        "quarter": quarter,
        "market": market,
        "collected_disclosures": collected_disclosures,
        "passed_count": passed_count,
        "report_path": str(report_path) if report_path else "",
        "log_path": str(log_path) if log_path else "",
    }
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    write_json(LOG_DIR / "last_run_status.json", payload)
    lines = [
        f"마지막 실행 시각: {payload['last_run_at']}",
        f"상태: {status}",
        f"검색 기간: {start} ~ {requested_end}",
        f"분석 분기: {quarter}",
        f"수집 공시 수: {collected_disclosures}",
        f"조건 통과 종목 수: {passed_count}",
        f"생성된 엑셀 파일: {payload['report_path']}",
        f"비고: {note}",
    ]
    (LOG_DIR / "last_run.txt").write_text("\n".join(lines) + "\n", encoding="utf-8-sig")

def status_note(status: str, reason: str, end: str, start: str) -> str:
    if status == "skipped":
        notes = {
            "before_start_date": "현재 실행일이 시작일 이전이면 기간 전 실행으로 정상",
            "after_end_date": "실행 기간 종료 후라 건너뜀",
            "weekend": "주말이라 건너뜀",
            "market_holiday": "국내 휴장일이라 건너뜀",
        }
        return notes.get(reason, reason)
    if status == "success" and start and end and end < start:
        return "현재 실행일이 시작일 이전이면 기간 전 실행으로 정상"
    if status == "success":
        return "오류 없이 종료됨"
    return f"실패: {reason}"

def load_env_api_key(env_path: Path) -> str | None:
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8", errors="replace").splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#") or "=" not in stripped:
                continue
            key, value = stripped.split("=", 1)
            if key.strip() and key.strip() not in os.environ:
                os.environ[key.strip()] = value.strip().strip('"').strip("'")
    api_key = (os.getenv("DART_API_KEY") or "").strip()
    return None if api_key in PLACEHOLDER_API_KEYS else api_key
