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


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "")

def format_yyyymmdd(value: str) -> str:
    value = value.strip()
    if re.fullmatch(r"\d{8}", value):
        return f"{value[:4]}-{value[4:6]}-{value[6:]}"
    return value

def normalize_text(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("\xa0", " ")
    return re.sub(r"\s+", " ", text)

def find_near_metric(text: str, section_terms: list[str], metric_terms: list[str]) -> float | None:
    for section in section_terms:
        for metric in metric_terms:
            pattern = re.compile(rf"{re.escape(section)}.{{0,220}}?{re.escape(metric)}.{{0,80}}?([-+]?\(?[\d,]+(?:\.\d+)?\)?)")
            match = pattern.search(text)
            if match:
                return parse_amount_to_eok(match.group(1))
    return None

def extract_labeled_numbers(text: str) -> dict[str, float | None]:
    mapping = {}
    patterns = {
        "h1_revenue": r"매출액[^0-9\-()]{0,20}당해사업연도[^0-9\-()]{0,20}([-+]?\(?[\d,]+(?:\.\d+)?\)?)",
        "h1_operating_income": r"영업이익[^0-9\-()]{0,20}당해사업연도[^0-9\-()]{0,20}([-+]?\(?[\d,]+(?:\.\d+)?\)?)",
        "y2025_h1_revenue": r"매출액[^0-9\-()]{0,20}직전사업연도[^0-9\-()]{0,20}([-+]?\(?[\d,]+(?:\.\d+)?\)?)",
        "y2025_h1_operating_income": r"영업이익[^0-9\-()]{0,20}직전사업연도[^0-9\-()]{0,20}([-+]?\(?[\d,]+(?:\.\d+)?\)?)",
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            mapping[key] = parse_amount_to_eok(match.group(1))
    return mapping

def parse_amount_to_eok(value: Any) -> float | None:
    if value is None:
        return None
    text = str(value).strip().replace(",", "")
    if not text or text in {"-", "N/A", "nan"}:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    try:
        number = float(text)
    except ValueError:
        return None
    return -number if negative else number

def growth(current: float | None, base: float | None) -> float | None:
    if current is None or base is None or base == 0:
        return None
    return (current - base) / abs(base) * 100

def is_turnaround(current: float | None, base: float | None) -> bool:
    return current is not None and base is not None and base <= 0 and current > 0

def is_loss_turnaround(current: float | None, base: float | None) -> bool:
    return current is not None and base is not None and base > 0 and current <= 0

def clean_html(html: str) -> str:
    return clean_text(re.sub(r"<[^>]+>", " ", html or ""))

def clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()

def decode_bytes(data: bytes) -> str:
    for encoding in ["utf-8", "euc-kr", "cp949"]:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")

def get_json(url: str, params: dict[str, Any], timeout: int) -> dict[str, Any]:
    return json.loads(get_text(f"{url}?{urlencode(params)}", timeout=timeout))

def get_text(url: str, timeout: int, encoding: str | None = None) -> str:
    request = Request(url, headers={"User-Agent": "Mozilla/5.0", "Accept-Language": "ko-KR,ko;q=0.9"})
    with urlopen(request, timeout=timeout) as response:
        raw = response.read()
    if encoding:
        return raw.decode(encoding, errors="replace")
    return decode_bytes(raw)

def get_bytes(url: str, params: dict[str, Any], timeout: int) -> bytes:
    request = Request(f"{url}?{urlencode(params)}", headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(request, timeout=timeout) as response:
        return response.read()

def dq(name: str, item: str, status: str, issue: str, needs_review: bool, source: str) -> dict[str, Any]:
    return {
        "종목명": name,
        "데이터 항목": item,
        "상태": status,
        "문제 내용": issue,
        "확인 필요 여부": "Y" if needs_review else "N",
        "출처 링크": source,
    }

def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        path.write_text("", encoding="utf-8-sig")
        return
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
