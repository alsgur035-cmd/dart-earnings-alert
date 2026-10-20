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

from .constants import CACHE_DIR, DART_API_BASE, DART_VIEWER_URL
from .util import compact, dq, extract_labeled_numbers, find_near_metric, get_bytes, normalize_text, parse_amount_to_eok
from .disclosures import extract_text_from_dart_document

def fetch_disclosure_text(api_key: str | None, rcp_no: str, offline: bool, quality_rows: list[dict[str, Any]]) -> str:
    if offline or not api_key:
        return ""
    cache_path = CACHE_DIR / "dart_document" / f"{rcp_no}.txt"
    if cache_path.exists():
        cached = cache_path.read_text(encoding="utf-8", errors="replace")
        # 2026-08-04: DART document.xml이 HTTP 200과 함께 <status> 오류 XML(예: "014 파일이
        # 존재하지 않습니다")을 반환할 때가 있는데, 예외를 던지지 않아 그대로 캐시돼 영구히
        # "확인 필요"로 잘못 탈락하는 사고가 있었다(디아이·씨엠티엑스). 오류 응답은 캐시로 인정하지 않고 재시도한다.
        if "<status>" in cached and "<status>000</status>" not in cached and len(cached) < 2000:
            cache_path.unlink(missing_ok=True)
        else:
            return cached

    try:
        raw = get_bytes(f"{DART_API_BASE}/document.xml", {"crtfc_key": api_key, "rcept_no": rcp_no}, timeout=20)
        if not raw.startswith(b"PK") and b"<status>000</status>" not in raw:
            quality_rows.append(dq("", "DART 공시원문", "WARN", f"{rcp_no} 원문 수집 실패: DART 오류 응답({raw[:120]!r})", True, DART_VIEWER_URL.format(rcp_no=rcp_no)))
            return ""
        text = extract_text_from_dart_document(raw)
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(text, encoding="utf-8")
        time.sleep(0.25)
        return text
    except Exception as exc:
        quality_rows.append(dq("", "DART 공시원문", "WARN", f"{rcp_no} 원문 수집 실패: {exc}", True, DART_VIEWER_URL.format(rcp_no=rcp_no)))
        return ""

def parse_financial_values(text: str) -> dict[str, float | None]:
    if not text:
        return {}
    unit_scale = detect_amount_unit_scale(text)
    table_values = extract_values_from_tables(text)
    normalized = normalize_text(text)
    values = {
        "q2_revenue": find_near_metric(normalized, ["당해실적", "당기실적", "당분기"], ["매출액", "영업수익"]),
        "q2_operating_income": find_near_metric(normalized, ["당해실적", "당기실적", "당분기"], ["영업이익"]),
        "q1_revenue": find_near_metric(normalized, ["전기실적", "전분기"], ["매출액", "영업수익"]),
        "q1_operating_income": find_near_metric(normalized, ["전기실적", "전분기"], ["영업이익"]),
        "y2025_q2_revenue": find_near_metric(normalized, ["전년동기실적", "전년동기"], ["매출액", "영업수익"]),
        "y2025_q2_operating_income": find_near_metric(normalized, ["전년동기실적", "전년동기"], ["영업이익"]),
        "y2025_q1_revenue": find_near_metric(normalized, ["전년전기실적", "전년전분기", "전년1분기", "2025Q1"], ["매출액", "영업수익"]),
        "y2025_q1_operating_income": find_near_metric(normalized, ["전년전기실적", "전년전분기", "전년1분기", "2025Q1"], ["영업이익"]),
        "h1_revenue": find_near_metric(normalized, ["누계실적", "당해사업연도누계", "당해사업연도"], ["매출액", "영업수익"]),
        "h1_operating_income": find_near_metric(normalized, ["누계실적", "당해사업연도누계", "당해사업연도"], ["영업이익"]),
        "y2025_h1_revenue": find_near_metric(normalized, ["전년동기누계", "직전사업연도", "전기사업연도"], ["매출액", "영업수익"]),
        "y2025_h1_operating_income": find_near_metric(normalized, ["전년동기누계", "직전사업연도", "전기사업연도"], ["영업이익"]),
    }
    for key, value in table_values.items():
        if value is not None:
            values[key] = value
    fallback_numbers = extract_labeled_numbers(normalized)
    for key, value in fallback_numbers.items():
        values.setdefault(key, value)
        if values[key] is None:
            values[key] = value
    return {key: (value * unit_scale if value is not None else None) for key, value in values.items()}

def detect_amount_unit_scale(text: str) -> float:
    """공시 표 상단의 '단위 : 억원, %' 같은 선언을 읽어 배율을 정한다.
    2026-08-07 버그 수정: 예전엔 '단위'와 단위 키워드(조원/원 등)가 문서 앞부분 20000자
    "어딘가에" 둘 다 있기만 하면 매칭했다 — 본문 비고란의 "수주잔액은 24.57조원입니다" 같은
    무관한 문장 때문에 실제 단위(억원)를 무시하고 10000배로 부풀리는 사고가 났다
    (LIG디펜스앤에어로스페이스, 매출·영익이 10000배 부풀려짐). '단위' 글자 바로 뒤
    좁은 창(30자)만 봐서 실제 선언된 단위만 읽도록 고쳤었다.
    2026-08-11 추가 수정: 그 20000자 컷오프 자체가 문제였다 — 회사개요 등 전치 텍스트가 길면
    실제 '단위 : 백만원, %' 선언이 20000자를 넘겨서 나오는 경우 아예 못 찾고 기본값(억원, 배율 1)으로
    빠져 매출·영익이 100배 부풀려졌다(서울바이오시스, 39470자 문서에서 단위 선언이 23126자 지점).
    컷오프를 없애고 문서 전체에서 '단위'를 찾는다(문서가 아무리 커도 최대 수십 KB라 성능 문제 없음)."""
    idx = text.find("단위")
    if idx == -1:
        return 1.0
    window = text[idx: idx + 30]
    if "조원" in window:
        return 10000
    if "억원" in window:
        return 1.0
    if "백만원" in window:
        return 0.01
    if "천원" in window:
        return 0.00001
    if re.search(r"[^백천]만원", window):
        return 0.0001
    if "원" in window:
        return 0.00000001
    return 1.0

def extract_values_from_tables(text: str) -> dict[str, float | None]:
    values: dict[str, float | None] = {}
    try:
        tables = pd.read_html(StringIO(text))
    except Exception:
        return values

    for table in tables:
        values.update({key: value for key, value in extract_provisional_performance_table(table).items() if value is not None})
        normalized = table.copy()
        normalized.columns = [flatten_column(column) for column in normalized.columns]
        for _, row in normalized.iterrows():
            row_values = {str(col): row[col] for col in normalized.columns}
            metric_name = compact(" ".join(str(value) for value in row_values.values()))
            metric = None
            if "영업이익" in metric_name:
                metric = "operating_income"
            elif "매출액" in metric_name or "영업수익" in metric_name:
                metric = "revenue"
            if metric is None:
                continue

            for column, raw_value in row_values.items():
                column_key = compact(column)
                value = parse_amount_to_eok(raw_value)
                if value is None:
                    continue
                target = period_key_from_column(column_key, metric)
                if target and values.get(target) is None:
                    values[target] = value
    return values

_EXACT_REVENUE_LABELS = {"매출액", "영업수익", "매출액영업수익"}
_EXACT_OPERATING_INCOME_LABELS = {"영업이익", "영업이익손실"}


def extract_provisional_performance_table(table: pd.DataFrame) -> dict[str, float | None]:
    table_text = compact(table.to_string())
    if not (("당기실적" in table_text or "당해실적" in table_text) and "전기실적" in table_text and "전년동기실적" in table_text):
        return {}

    values: dict[str, float | None] = {}
    for _, row in table.iterrows():
        row_items = list(row.tolist())
        if len(row_items) < 7:
            continue
        metric = compact(row_items[0])
        line_type = compact(row_items[1])
        if len(metric) > 20 or len(line_type) > 20:
            continue
        if not any(token in line_type for token in ["당해실적", "당기실적", "누계실적"]):
            continue
        # 2026-08-12 버그 수정: "매출액" in metric처럼 부분일치로 판정하면 일부 유통업체(예: 광주신세계)가
        # 함께 공시하는 "총매출액"(특정매입원가 포함, 회계상 매출액과 다른 별도 지표) 행도 걸려서
        # 뒤에 나온 행이 값을 덮어써 진짜 매출액이 사라지는 사고가 났다(광주신세계 매출이 2.9배 부풀려짐).
        # 라벨을 정확히 일치하는 것만 인정한다 — "총매출액"/"순매출액" 등 접두어 붙은 변형은 제외.
        if metric in _EXACT_OPERATING_INCOME_LABELS:
            suffix = "operating_income"
        elif metric in _EXACT_REVENUE_LABELS:
            suffix = "revenue"
        else:
            continue

        current = parse_amount_to_eok(row_items[2])
        prior_quarter = parse_amount_to_eok(row_items[3])
        prior_year = parse_amount_to_eok(row_items[6])

        if "누계" in line_type:
            values[f"h1_{suffix}"] = current
            values[f"y2025_h1_{suffix}"] = prior_year
        elif "당해실적" in line_type or "당기실적" in line_type:
            values[f"q2_{suffix}"] = current
            values[f"q1_{suffix}"] = prior_quarter
            values[f"y2025_q2_{suffix}"] = prior_year
    return values

def period_key_from_column(column_key: str, metric: str) -> str | None:
    prefix = "operating_income" if metric == "operating_income" else "revenue"
    if any(token in column_key for token in ["당해실적", "당기실적", "당분기"]):
        return f"q2_{prefix}"
    if any(token in column_key for token in ["전기실적", "전분기"]):
        return f"q1_{prefix}"
    if any(token in column_key for token in ["전년전기실적", "전년전분기", "전년1분기", "2025Q1"]):
        return f"y2025_q1_{prefix}"
    if any(token in column_key for token in ["누계실적", "당해사업연도누계", "당해사업연도", "당기누계"]):
        return f"h1_{prefix}"
    if any(token in column_key for token in ["전년동기누계", "직전사업연도", "전기사업연도", "전기누계"]):
        return f"y2025_h1_{prefix}"
    if any(token in column_key for token in ["전년동기실적", "전년동기"]):
        return f"y2025_q2_{prefix}"
    return None

def flatten_column(column: Any) -> str:
    if isinstance(column, tuple):
        parts = [str(part) for part in column if str(part) and not str(part).startswith("Unnamed")]
        return " ".join(parts).strip()
    return str(column).strip()

# ---------------------------------------------------------------------------
# 정식 반기보고서 처리: 본문 텍스트 대신 DART 정식 재무제표 API(fnlttSinglAcntAll)로 수집.
# 필드 의미(삼성전자 2025 반기로 실측 검증):
#   thstrm_amount     = 당기 3개월(2Q 단독)   thstrm_add_amount = 당기 누적(반기 1H)
#   frmtrm_q_amount   = 전년동기 3개월(전년 2Q) frmtrm_add_amount = 전년 누적(전년 1H)
# 1분기 값(QoQ용)은 reprt_code=11013(1분기보고서)의 thstrm_amount.
# ---------------------------------------------------------------------------

_REVENUE_IDS = {"ifrs-full_Revenue", "ifrs_Revenue"}
_REVENUE_NAMES = {"매출액", "수익(매출액)", "영업수익", "매출"}
# ifrs-full_ProfitLossFromOperatingActivities: 금융사(증권 등) XBRL이 쓰는 표준 영업이익 태그
# (2026-08-11 신영증권 사례 — dart_OperatingIncomeLoss 없이 이 태그만 있어 전분기 영익이 '확인 필요'로 빠졌음)
_OP_IDS = {"dart_OperatingIncomeLoss", "ifrs-full_ProfitLossFromOperatingActivities"}
_OP_NAMES = {"영업이익", "영업이익(손실)"}

_ITEM_PREFIX_RE = re.compile(r"^[0-9IVXivxⅠ-Ⅻⅰ-ⅻ]+\.")


def _find_fs_row(rows: list[dict[str, Any]], ids: set[str], names: set[str]) -> dict[str, Any] | None:
    for sj in ("IS", "CIS"):  # 손익계산서 우선, 없으면 포괄손익계산서
        subset = [r for r in rows if r.get("sj_div") == sj]
        for r in subset:
            if (r.get("account_id") or "").strip() in ids:
                return r
        for r in subset:
            # 회사·연도에 따라 "III. 영업이익"처럼 항목 번호가 붙기도 함(신영증권 FY2025) — 접두어 떼고도 비교
            nm = compact(r.get("account_nm") or "")
            if nm in names or _ITEM_PREFIX_RE.sub("", nm) in names:
                return r
    return None


def _fs_amount_eok(row: dict[str, Any] | None, key: str) -> float | None:
    if not row:
        return None
    # 2026-08-18 버그 수정: fnlttSinglAcntAll 응답엔 currency 필드가 있는데, 해외 자회사 비중이 큰
    # 회사(두산밥캣 등)는 연결재무제표를 KRW가 아니라 USD로 신고한다. 예전엔 currency를 안 보고
    # 무조건 raw/1e8(원→억원)로 계산해서 두산밥캣 매출이 1484배(대략 USD/KRW 환율만큼) 축소됐다
    # (정정 반기보고서에서 재현 — 매출 46,953억이 31.64로, PER이 12.3에서 18,458로 튐).
    # 환율을 하드코딩해 억지로 환산하면 환율 변동에 따라 또 틀리므로, KRW가 아니면 계산하지 않고
    # None(확인 필요)으로 안전하게 빠진다 — 틀린 숫자보다 빈 칸이 낫다.
    currency = (row.get("currency") or "").strip().upper()
    if currency and currency not in ("KRW", "WON"):
        return None
    raw = str(row.get(key) or "").replace(",", "").strip()
    if not raw or raw in ("-", "None"):
        return None
    try:
        return round(float(raw) / 1e8, 2)  # 원 → 억원
    except ValueError:
        return None



def _doc_row_values(text: str, acode: str) -> list[float] | None:
    """원문 손익계산서에서 XBRL 태그(ACODE)가 붙은 행의 숫자를 추출.
    표준 열 순서 [당기3개월, 당기누적, 전년동기3개월, 전년동기누적], 단위 원 → 억원.
    연결·별도가 둘 다 있으면 먼저 나오는(연결) 행을 쓴다."""
    for tr in re.split(r"<TR[\s>]", text):
        if f'ACODE="{acode}"' not in tr:
            continue
        vals: list[float] = []
        for token in re.findall(r">\s*(\(?-?[0-9][0-9,]*\)?)\s*</T[DE]>", tr):
            digits = re.sub(r"[^0-9]", "", token)
            if not digits:
                continue
            value = round(float(digits) / 1e8, 2)
            if token.startswith("(") or token.startswith("-"):
                value = -value
            vals.append(value)
        if len(vals) >= 4:
            return vals[:4]
    return None


def fetch_periodic_financials(
    api_key: str | None,
    corp_code: str,
    quarter: str,
    rcp_no: str,
    offline: bool,
    quality_rows: list[dict[str, Any]],
    corp_name: str,
    link: str,
    report_type: str = "반기보고서",
) -> dict[str, float]:
    """정식 반기/분기보고서의 재무를 fnlttSinglAcntAll에서 수집.

    결산월이 회사마다 달라 같은 "(2026.05)" 보고서도 회사에 따라 11012(반기)/11013(1분기)/
    11014(3분기)로 인덱싱된다(풍강=8월결산 3분기, 현대약품=11월결산 반기 실측). 따라서
    세 reprt_code를 모두 조회해 **접수번호(rcp_no)가 정확히 일치하는 응답만** 사용한다.
    bsns_year는 보고서 기간의 연도(분기연도)와 일치함을 실측 확인.
    전분기(QoQ 기준값)는 일치한 유형에 따라: 반기→1분기보고서, 3분기→반기보고서,
    1분기→전년 사업보고서(연간)−전년 3분기 누적 차감으로 구한다(없으면 확인 필요).
    """
    from .util import get_json

    if offline or not api_key or not corp_code:
        return {}
    bsns_year = quarter[:4]

    def load(reprt_code: str, year: str) -> list[dict[str, Any]]:
        # 연결(CFS) 우선, 손익 계정이 없으면 개별/별도(OFS) 폴백 (연결대상 없는 회사)
        for fs_div in ("CFS", "OFS"):
            try:
                payload = get_json(
                    f"{DART_API_BASE}/fnlttSinglAcntAll.json",
                    {"crtfc_key": api_key, "corp_code": corp_code, "bsns_year": year,
                     "reprt_code": reprt_code, "fs_div": fs_div},
                    timeout=20,
                )
            except Exception as exc:
                quality_rows.append(dq(corp_name, "정기보고서 재무", "WARN", f"fnlttSinglAcntAll 조회 실패: {exc}", True, link))
                return []
            if str(payload.get("status")) == "000" and payload.get("list"):
                rows = payload["list"]
                if _find_fs_row(rows, _REVENUE_IDS, _REVENUE_NAMES) or _find_fs_row(rows, _OP_IDS, _OP_NAMES):
                    return rows
        return []

    matched_rows: list[dict[str, Any]] = []
    matched_code = ""
    loaded: dict[str, list[dict[str, Any]]] = {}
    for reprt_code in ("11012", "11013", "11014"):
        rows = load(reprt_code, bsns_year)
        loaded[reprt_code] = rows
        if rows and str(rows[0].get("rcept_no") or "") == rcp_no:
            matched_rows, matched_code = rows, reprt_code
            break
    if not matched_rows:
        # XBRL 미반영(신규 제출 직후) → 공시 원문에서 즉시 추출 (실시간 수집, 사용자 방침)
        text = fetch_disclosure_text(api_key, rcp_no, offline, quality_rows)
        rev_vals = _doc_row_values(text, "ifrs-full_Revenue") if text else None
        op_vals = _doc_row_values(text, "dart_OperatingIncomeLoss") if text else None
        if not rev_vals and not op_vals:
            quality_rows.append(dq(corp_name, "정기보고서 재무", "WARN",
                                   "XBRL 미반영 + 원문 추출 실패 — 다음 갱신 때 재시도", True, link))
            return {}
        parsed: dict[str, float | None] = {}
        if rev_vals:
            parsed["q2_revenue"], parsed["h1_revenue"], parsed["y2025_q2_revenue"], parsed["y2025_h1_revenue"] = rev_vals
        if op_vals:
            parsed["q2_operating_income"], parsed["h1_operating_income"], parsed["y2025_q2_operating_income"], parsed["y2025_h1_operating_income"] = op_vals
        # 전분기(QoQ 기준값) 도출
        if report_type == "반기보고서":
            if parsed.get("h1_revenue") is not None and parsed.get("q2_revenue") is not None:
                parsed["q1_revenue"] = round(parsed["h1_revenue"] - parsed["q2_revenue"], 2)
            if parsed.get("h1_operating_income") is not None and parsed.get("q2_operating_income") is not None:
                parsed["q1_operating_income"] = round(parsed["h1_operating_income"] - parsed["q2_operating_income"], 2)
        else:  # 분기보고서: 누적≈3개월이면 회계 1분기, 아니면 3분기(직전 반기의 3개월값이 전분기)
            q2r, h1r = parsed.get("q2_revenue"), parsed.get("h1_revenue")
            if q2r is not None and h1r is not None and abs(h1r - q2r) <= abs(h1r) * 0.01:
                prev_annual = load("11011", str(int(bsns_year) - 1))
                prev_q3 = load("11014", str(int(bsns_year) - 1))
                a_rev, a_op = _find_fs_row(prev_annual, _REVENUE_IDS, _REVENUE_NAMES), _find_fs_row(prev_annual, _OP_IDS, _OP_NAMES)
                c_rev, c_op = _find_fs_row(prev_q3, _REVENUE_IDS, _REVENUE_NAMES), _find_fs_row(prev_q3, _OP_IDS, _OP_NAMES)
                ar, cr = _fs_amount_eok(a_rev, "thstrm_amount"), _fs_amount_eok(c_rev, "thstrm_add_amount")
                ao, co = _fs_amount_eok(a_op, "thstrm_amount"), _fs_amount_eok(c_op, "thstrm_add_amount")
                parsed["q1_revenue"] = round(ar - cr, 2) if ar is not None and cr is not None else None
                parsed["q1_operating_income"] = round(ao - co, 2) if ao is not None and co is not None else None
            else:
                prev = load("11012", bsns_year)
                prev_rev, prev_op = _find_fs_row(prev, _REVENUE_IDS, _REVENUE_NAMES), _find_fs_row(prev, _OP_IDS, _OP_NAMES)
                parsed["q1_revenue"] = _fs_amount_eok(prev_rev, "thstrm_amount")
                parsed["q1_operating_income"] = _fs_amount_eok(prev_op, "thstrm_amount")
        quality_rows.append(dq(corp_name, "정기보고서 재무", "INFO",
                               "XBRL 미반영이라 공시 원문에서 직접 추출함(단위: 원, 연결 우선)", False, link))
        return {k: v for k, v in parsed.items() if v is not None}

    rev = _find_fs_row(matched_rows, _REVENUE_IDS, _REVENUE_NAMES)
    op = _find_fs_row(matched_rows, _OP_IDS, _OP_NAMES)
    parsed: dict[str, float | None] = {
        "q2_revenue": _fs_amount_eok(rev, "thstrm_amount"),
        "h1_revenue": _fs_amount_eok(rev, "thstrm_add_amount") or _fs_amount_eok(rev, "thstrm_amount"),
        "y2025_q2_revenue": _fs_amount_eok(rev, "frmtrm_q_amount"),
        "y2025_h1_revenue": _fs_amount_eok(rev, "frmtrm_add_amount") or _fs_amount_eok(rev, "frmtrm_q_amount"),
        "q2_operating_income": _fs_amount_eok(op, "thstrm_amount"),
        "h1_operating_income": _fs_amount_eok(op, "thstrm_add_amount") or _fs_amount_eok(op, "thstrm_amount"),
        "y2025_q2_operating_income": _fs_amount_eok(op, "frmtrm_q_amount"),
        "y2025_h1_operating_income": _fs_amount_eok(op, "frmtrm_add_amount") or _fs_amount_eok(op, "frmtrm_q_amount"),
    }

    # 전분기(QoQ 기준값)
    if matched_code == "11012":  # 반기 → 같은 해 1분기보고서
        prev = loaded.get("11013") or load("11013", bsns_year)
        prev_rev, prev_op = _find_fs_row(prev, _REVENUE_IDS, _REVENUE_NAMES), _find_fs_row(prev, _OP_IDS, _OP_NAMES)
        parsed["q1_revenue"] = _fs_amount_eok(prev_rev, "thstrm_amount")
        parsed["q1_operating_income"] = _fs_amount_eok(prev_op, "thstrm_amount")
    elif matched_code == "11014":  # 3분기 → 같은 해 반기보고서의 당기 3개월
        prev = loaded.get("11012") or load("11012", bsns_year)
        prev_rev, prev_op = _find_fs_row(prev, _REVENUE_IDS, _REVENUE_NAMES), _find_fs_row(prev, _OP_IDS, _OP_NAMES)
        parsed["q1_revenue"] = _fs_amount_eok(prev_rev, "thstrm_amount")
        parsed["q1_operating_income"] = _fs_amount_eok(prev_op, "thstrm_amount")
    elif matched_code == "11013":  # 1분기 → 전년 사업보고서(연간) − 전년 3분기 누적
        prev_year = str(int(bsns_year) - 1)
        annual = load("11011", prev_year)
        q3 = load("11014", prev_year)
        a_rev, a_op = _find_fs_row(annual, _REVENUE_IDS, _REVENUE_NAMES), _find_fs_row(annual, _OP_IDS, _OP_NAMES)
        c_rev, c_op = _find_fs_row(q3, _REVENUE_IDS, _REVENUE_NAMES), _find_fs_row(q3, _OP_IDS, _OP_NAMES)
        ar, cr = _fs_amount_eok(a_rev, "thstrm_amount"), _fs_amount_eok(c_rev, "thstrm_add_amount")
        ao, co = _fs_amount_eok(a_op, "thstrm_amount"), _fs_amount_eok(c_op, "thstrm_add_amount")
        parsed["q1_revenue"] = round(ar - cr, 2) if ar is not None and cr is not None else None
        parsed["q1_operating_income"] = round(ao - co, 2) if ao is not None and co is not None else None
    if parsed.get("q1_revenue") is None and parsed.get("q1_operating_income") is None:
        quality_rows.append(dq(corp_name, "전분기 재무", "WARN", "전분기(QoQ 기준) 값 확인 필요 — YoY 조건으로만 판정", True, link))

    return {k: v for k, v in parsed.items() if v is not None}


# 하위호환 별칭(기존 이름으로 접근하는 코드/테스트 대비)
fetch_semiannual_financials = fetch_periodic_financials
