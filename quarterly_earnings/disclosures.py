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

from .constants import CACHE_DIR, DART_API_BASE, DART_VIEWER_URL, EXCLUDE_NAME_PATTERNS, TARGET_REPORT_NAMES, TARGET_REPORT_PATTERNS
from .models import Disclosure
from .util import compact, decode_bytes, dq, format_yyyymmdd, get_bytes, get_json

def fetch_dart_disclosures(
    api_key: str,
    start: str,
    end: str,
    market: str,
    logger: logging.Logger,
    quality_rows: list[dict[str, Any]],
    offline: bool,
) -> list[Disclosure]:
    if offline:
        quality_rows.append(dq("", "DART", "WARN", "offline 옵션으로 DART 공시 검색을 건너뜀", True, ""))
        return []

    corp_classes = {"kospi": ["Y"], "kosdaq": ["K"], "all": ["Y", "K"]}[market]
    results: list[Disclosure] = []
    for corp_cls in corp_classes:
        page_no = 1
        while True:
            params = {
                "crtfc_key": api_key,
                "bgn_de": start.replace("-", ""),
                "end_de": end.replace("-", ""),
                "corp_cls": corp_cls,
                "page_no": page_no,
                "page_count": 100,
            }
            try:
                payload = get_json(f"{DART_API_BASE}/list.json", params, timeout=20)
            except Exception as exc:
                quality_rows.append(dq("", "DART 공시검색", "ERROR", f"DART 공시검색 실패: {exc}", True, ""))
                logger.exception("DART disclosure search failed")
                break

            status = str(payload.get("status", ""))
            if status == "013":
                break
            if status != "000":
                quality_rows.append(dq("", "DART 공시검색", "ERROR", f"{status} {payload.get('message', '')}", True, ""))
                break

            for item in payload.get("list") or []:
                stock_code = str(item.get("stock_code") or "").strip().zfill(6)
                results.append(
                    Disclosure(
                        rcp_no=str(item.get("rcept_no") or "").strip(),
                        corp_code=str(item.get("corp_code") or "").strip(),
                        stock_code=stock_code,
                        corp_name=str(item.get("corp_name") or "").strip(),
                        report_nm=str(item.get("report_nm") or "").strip(),
                        rcept_dt=format_yyyymmdd(str(item.get("rcept_dt") or "")),
                        corp_cls=str(item.get("corp_cls") or corp_cls).strip(),
                        flr_nm=str(item.get("flr_nm") or "").strip(),
                        rm=str(item.get("rm") or "").strip(),
                    )
                )

            total_page = int(payload.get("total_page") or 1)
            if page_no >= total_page:
                break
            page_no += 1
            time.sleep(0.15)
    return results

def filter_report_names(disclosures: list[Disclosure]) -> list[Disclosure]:
    target_keys = [compact(name) for name in TARGET_REPORT_NAMES]
    filtered = []
    for item in disclosures:
        report_key = compact(item.report_nm).replace("[정정]", "").replace("정정", "")
        # 2026-10-26: 금융지주사 등이 자회사 실적을 지주사 이름으로 내는 '(자회사의 주요경영사항)' 공시는 지주사 실적이 아니므로 제외(본 공시를 기다림).
        if "자회사의주요경영사항" in report_key:
            continue
        if any(target in report_key for target in target_keys) or any(
            pattern.search(report_key) for pattern in TARGET_REPORT_PATTERNS
        ):
            filtered.append(item)
    return filtered

def dedupe_disclosures(
    disclosures: list[Disclosure],
    api_key: str | None = None,
    offline: bool = False,
    quality_rows: list[dict[str, Any]] | None = None,
) -> list[Disclosure]:
    """회사당 1건만 남긴다. 기본 우선순위는 disclosure_rank(유형>정정여부>접수일>rcp_no).
    2026-08-12 버그 수정: 같은 회사가 같은 날 같은 유형(잠정실적)의 disclosure_rank가 완전히
    동점인 공시를 2건 내는 경우가 있다 — 예: 광주신세계가 08-11에 '영업(잠정)실적(공정공시)'을
    두 번 냈는데, 하나는 진짜 2분기(2026.04.01~06.30, 91일) 실적이고 다른 하나는 7월 한 달치
    (2026.07.01~07.31, 31일) 월간 매출 추정치였다(영업이익 칸이 전부 "-"). 예전엔 마지막 타이브레이커가
    rcp_no 숫자 크기였는데, 이 경우 우연히 월간 쪽 rcp_no가 더 커서 그게 선택됐고 진짜 분기실적이
    버려졌다(영업이익 항목이 "확인 필요"로 빠짐). 이제 동점 후보가 있으면 원문에서 '당기실적' 기간을
    읽어 분기 길이(약 90일)에 가까운 쪽을 우선한다 — 월간형(약 30일 이하)은 뒤로 민다."""
    by_rcp = {item.rcp_no: item for item in disclosures if item.rcp_no}
    groups: dict[str, list[Disclosure]] = {}
    for item in by_rcp.values():
        key = item.stock_code or item.corp_code or item.corp_name
        groups.setdefault(key, []).append(item)

    by_company: dict[str, Disclosure] = {}
    for key, items in groups.items():
        # 2026-10-19: 월간 잠정(실적기간 ≤45일: 신세계·이마트·오리온 등의 월 매출 공시)은 분기실적이 아니다.
        # 같은 회사에 분기 길이 공시가 함께 있으면 월간 건은 후보에서 뺀다(접수일이 더 늦어도 월간이 이기지 않게).
        if len(items) > 1 and not offline and api_key:
            non_monthly = []
            for it in items:
                if it.report_type != "잠정실적":
                    non_monthly.append(it)
                    continue
                days = _extract_report_period_days(_fetch_document_text_cached(api_key, it.rcp_no, offline, quality_rows) or "")
                if days is None or days > 45:
                    non_monthly.append(it)
            if non_monthly:
                items = non_monthly
        items.sort(key=disclosure_rank, reverse=True)
        # 동점 판정은 (유형+정정점수, 접수일)까지만 본다 — rcp_no는 공시마다 유일해서
        # 여기까지 포함하면 절대 동점이 안 나므로 일부러 뺀다.
        top_score, top_date = disclosure_rank(items[0])[:2]
        tied = [item for item in items if disclosure_rank(item)[:2] == (top_score, top_date)]
        if len(tied) <= 1:
            by_company[key] = items[0]
            continue
        by_company[key] = _pick_by_report_period(tied, api_key, offline, quality_rows)
    return list(by_company.values())


def _pick_by_report_period(
    tied: list[Disclosure],
    api_key: str | None,
    offline: bool,
    quality_rows: list[dict[str, Any]] | None,
) -> Disclosure:
    """disclosure_rank가 완전히 동점인 후보들 중 원문 '당기실적' 기간이 분기 길이(75~100일)에
    가장 가까운 것을 고른다. 전부 실패/판별불가면 예전 동작(rcp_no 최대)으로 안전하게 폴백."""
    scored: list[tuple[int, Disclosure]] = []
    for item in tied:
        text = _fetch_document_text_cached(api_key, item.rcp_no, offline, quality_rows)
        days = _extract_report_period_days(text) if text else None
        if days is not None:
            scored.append((days, item))
    if scored:
        # 분기(약 91일)에 가장 가까운 후보를 고른다. 월간형(약 30일 이하)은 거리값이 커서 자동으로 밀려난다.
        best_days, best_item = min(scored, key=lambda pair: abs(pair[0] - 91))
        if quality_rows is not None and len(scored) < len(tied):
            quality_rows.append(dq(
                tied[0].corp_name, "공시 중복판정", "INFO",
                f"동일 유형 동점 공시 {len(tied)}건 중 {len(scored)}건만 기간 판별 성공, {best_days}일짜리 채택",
                False, tied[0].link,
            ))
        return best_item
    # 기간 판별이 전부 실패하면 기존 방식(rcp_no 최대)으로 폴백
    if quality_rows is not None:
        quality_rows.append(dq(
            tied[0].corp_name, "공시 중복판정", "WARN",
            f"동일 유형 동점 공시 {len(tied)}건의 실적기간을 판별하지 못해 rcp_no 최대값으로 폴백",
            True, tied[0].link,
        ))
    return max(tied, key=lambda item: item.rcp_no)


def _fetch_document_text_cached(
    api_key: str | None,
    rcp_no: str,
    offline: bool,
    quality_rows: list[dict[str, Any]] | None,
) -> str:
    """financials.fetch_disclosure_text와 같은 캐시(rcp_no.txt)를 공유한다 — 여기서 먼저 받아두면
    나중 financials 단계가 캐시를 그대로 재사용해 이중 호출이 되지 않는다."""
    if offline or not api_key:
        return ""
    cache_path = CACHE_DIR / "dart_document" / f"{rcp_no}.txt"
    if cache_path.exists():
        cached = cache_path.read_text(encoding="utf-8", errors="replace")
        if "<status>" in cached and "<status>000</status>" not in cached and len(cached) < 2000:
            cache_path.unlink(missing_ok=True)
        else:
            return cached
    try:
        raw = get_bytes(f"{DART_API_BASE}/document.xml", {"crtfc_key": api_key, "rcept_no": rcp_no}, timeout=20)
        if not raw.startswith(b"PK") and b"<status>000</status>" not in raw:
            return ""
        text = extract_text_from_dart_document(raw)
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(text, encoding="utf-8")
        time.sleep(0.25)
        return text
    except Exception:
        return ""


def _extract_report_period_days(text: str) -> int | None:
    """공시 원문 상단 '실적기간' 표에서 '당기실적 YYYY-MM-DD ~ YYYY-MM-DD'를 찾아 일수를 반환.
    못 찾으면 None(분기/반기/손익구조 등 표 형식이 다른 유형에서도 안전하게 폴백하도록)."""
    idx = text.find("당기실적")
    if idx == -1:
        return None
    window = text[idx: idx + 400]
    dates = re.findall(r"(20\d{2}-\d{2}-\d{2})", window)
    if len(dates) < 2:
        return None
    try:
        start = datetime.strptime(dates[0], "%Y-%m-%d")
        end = datetime.strptime(dates[1], "%Y-%m-%d")
    except ValueError:
        return None
    return (end - start).days + 1


def disclosure_rank(item: Disclosure) -> tuple[int, str, str]:
    # 잠정실적 > 반기보고서 > 손익구조변동. 정정 여부는 같은 유형 안에서만 우선순위를 올린다
    # (예전처럼 detail+correction 합산하면 정정 손익구조가 잠정과 동점이 되는 문제가 있었음).
    type_score = {"잠정실적": 3, "반기보고서": 2, "분기보고서": 2, "손익구조변동": 1}.get(item.report_type, 0)
    correction_score = 1 if item.is_correction else 0
    return type_score * 2 + correction_score, item.rcept_dt, item.rcp_no

def filter_common_stocks(disclosures: list[Disclosure], excluded_rows: list[dict[str, Any]]) -> list[Disclosure]:
    results = []
    for item in disclosures:
        reason = common_stock_exclude_reason(item.corp_name, item.stock_code)
        if reason:
            excluded_rows.append(
                {
                    "제외된 종목": item.corp_name,
                    "종목코드": item.stock_code,
                    "제외 사유": reason,
                    "우선주/스팩/ETF/ETN/리츠/파생상품/비보통주 여부": "Y",
                    "공시 링크": item.link,
                }
            )
            continue
        results.append(item)
    return results

def common_stock_exclude_reason(name: str, ticker: str) -> str:
    reasons = []
    if not re.fullmatch(r"\d{6}", ticker or ""):
        reasons.append("종목코드 없음")
    for label, pattern in EXCLUDE_NAME_PATTERNS:
        if pattern.search(name or ""):
            reasons.append(label)
    return ", ".join(dict.fromkeys(reasons))

def base_report_name(name: str) -> str:
    return compact(name).replace("[정정]", "").replace("정정", "")

def extract_text_from_dart_document(raw: bytes) -> str:
    if raw.startswith(b"PK"):
        texts = []
        with zipfile.ZipFile(BytesIO(raw)) as zf:
            for name in zf.namelist():
                if name.lower().endswith((".xml", ".html", ".htm", ".txt")):
                    data = zf.read(name)
                    texts.append(decode_bytes(data))
        return "\n".join(texts)
    return decode_bytes(raw)
