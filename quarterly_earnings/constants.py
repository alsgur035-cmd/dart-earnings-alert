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


SCRIPT_DIR = Path(__file__).resolve().parent.parent

PROJECT_DIR = SCRIPT_DIR.parents[1]

CONFIG_DIR = PROJECT_DIR / "00_코덱스_작업용" / "02_설정파일"

WORK_DIR = PROJECT_DIR / "01_결과물" / "03_분기실적_필터링"

TEMPLATE_PATH = WORK_DIR / "양식.xlsx"

REPORT_DIR = WORK_DIR / "reports"

DATA_DIR = WORK_DIR / "data"

LOG_DIR = WORK_DIR / "logs"

CACHE_DIR = WORK_DIR / "cache"

DART_API_BASE = "https://opendart.fss.or.kr/api"

DART_VIEWER_URL = "https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcp_no}"

NAVER_ITEM_URL = "https://finance.naver.com/item/main.naver?code={ticker}"

PLACEHOLDER_API_KEYS = {"", "your_api_key_here", "YOUR_API_KEY"}

TARGET_REPORT_NAMES = [
    "연결재무제표기준영업(잠정)실적(공정공시)",
    "영업(잠정)실적(공정공시)",
    "매출액또는손익구조30%(대규모법인은15%)이상변동",
    "매출액또는손익구조30%(대규모법인은15%)이상변경",
    "매출액 또는 손익구조 30%(대규모법인은 15%) 이상 변동",
    "매출액 또는 손익구조 30%(대규모법인은 15%) 이상 변경",
    "매출액 또는 손익구조 30% 이상 변동",
    "매출액 또는 손익구조 30% 이상 변경",
    "매출액 또는 손익구조 15% 이상 변동",
    "매출액 또는 손익구조 15% 이상 변경",
]

# 잠정 없이 정식 반기/분기보고서로 실적을 내는 종목 대응.
# 결산월이 회사마다 달라 같은 시즌 보고서라도 기간표기가 다르다
# (12월결산 반기=2026.06, 11월결산 현대약품 반기=2026.05, 8월결산 풍강 3분기=2026.05).
# 달력 2분기(4~6월)와 겹치는 기간(2026.04~08)만 대상 — 과거연도 기재정정과
# 1분기 시즌 보고서(2026.03 지연제출·정정)는 제외.
TARGET_REPORT_PATTERNS = [
    re.compile(r"(반기|분기)보고서\(2026\.0[4-8]\)"),
]

MAIN_HEADERS = [
    "종목명",
    "업종",
    "시가총액",
    "매출액(누적, 억원)",
    "영익(단독, 억원)",
    "영익(누적, 억원)",
    "PER",
    "QoQ 매출액(%)",
    "QoQ 영익(%)",
    "YoY 매출액(%)",
    "YoY 영익(%)",
    "YoY누적 매출액(%)",
    "YoY누적 영익(%)",
]

EXCLUDE_NAME_PATTERNS = [
    ("우선주", re.compile(r"(우$|우B$|우C$|우선주|[0-9]+우$|우\(전환\))")),
    ("스팩", re.compile(r"(스팩|SPAC|기업인수목적)", re.IGNORECASE)),
    ("ETF", re.compile(r"(ETF|KODEX|TIGER|ACE|SOL|KOSEF|KBSTAR|HANARO|ARIRANG|RISE)", re.IGNORECASE)),
    ("ETN", re.compile(r"(ETN)", re.IGNORECASE)),
    ("ELW", re.compile(r"(ELW)", re.IGNORECASE)),
    ("리츠", re.compile(r"(리츠|REIT)", re.IGNORECASE)),
    ("인프라펀드", re.compile(r"(인프라|맥쿼리인프라|펀드)", re.IGNORECASE)),
    ("선박펀드", re.compile(r"(선박)", re.IGNORECASE)),
    ("파생상품", re.compile(r"(선물|레버리지|인버스|파생)", re.IGNORECASE)),
]

QUARTER_PERIODS = {
    "2026Q2": {
        "quarter_start": "2026-04-01",
        "quarter_end": "2026-06-30",
        "prev_quarter": "2026Q1",
        "yoy_quarter": "2025Q2",
        "current_cumulative": "2026H1",
        "yoy_cumulative": "2025H1",
    },
    # 2026-10-19 추가: 3분기. 당기=3Q 3개월, 전기=2Q 3개월(QoQ), 누적=9개월(1~3Q)
    "2026Q3": {
        "quarter_start": "2026-07-01",
        "quarter_end": "2026-09-30",
        "prev_quarter": "2026Q2",
        "yoy_quarter": "2025Q3",
        "current_cumulative": "2026 9M",
        "yoy_cumulative": "2025 9M",
    },
}

# 2026-10-19 추가: 분기별 정기보고서 패턴·출력 라벨.
# 내부 필드명(q2_*/h1_*/q1_*/y2025_*)은 역사적 이유로 유지하고 "당기/누적/전기/전년동기/전년누적" 슬롯 의미로만 쓴다.
# 라벨만 분기별로 바꾼다. 결산월·보고서종류 가정 금지: 달력 해당분기와 겹치는 보고서 기간 전부 포함.
_PERIODIC_PATTERNS_BY_Q = {
    "Q2": r"(반기|분기)보고서\(2026\.0[4-8]\)",
    "Q3": r"(반기|분기)보고서\(2026\.(0[7-9]|1[0-1])\)",
}
_LABELS_BY_Q = {
    "Q2": {"cur": "2026Q2", "cum": "2026H1 누적", "prev": "2026Q1", "yoy": "2025Q2", "yoycum": "2025H1 누적", "prev1": "2025Q1"},
    "Q3": {"cur": "2026Q3", "cum": "2026 9M 누적", "prev": "2026Q2", "yoy": "2025Q3", "yoycum": "2025 9M 누적", "prev1": "2025Q2"},
}
QUARTER_LABELS: dict[str, str] = dict(_LABELS_BY_Q["Q2"])


def set_active_quarter(quarter: str) -> None:
    """cli가 실행 시작 때 호출. 정기보고서 패턴(리스트 in-place)과 출력 라벨을 해당 분기로 맞춘다."""
    key = "Q3" if str(quarter).endswith("Q3") else "Q2"
    TARGET_REPORT_PATTERNS[:] = [re.compile(_PERIODIC_PATTERNS_BY_Q[key])]
    QUARTER_LABELS.clear()
    QUARTER_LABELS.update(_LABELS_BY_Q[key])
