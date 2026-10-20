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

from .constants import CACHE_DIR, DART_API_BASE, NAVER_ITEM_URL
from .models import Disclosure
from .util import clean_html, clean_text, compact, dq, get_json, get_text, parse_amount_to_eok

# ---------------------------------------------------------------------------
# 2026-10-19: 네이버가 finance.naver.com/item/main.naver 를 stock.naver.com SPA로 302 리다이렉트하면서
# (UA와 무관하게 전 종목) _market_sum / 동일업종비교 HTML이 사라져 시총·업종이 전부 '확인 필요'가 됐다.
# 대체 소스: 네이버 모바일 JSON API. 시총=totalInfos.marketValue("1,256조 9,499억"),
# 업종=integration.industryCode → /api/stocks/industry 그룹목록의 업종명(구 '동일업종비교' 업종명과 동일 체계, 무가공).
# 구 HTML 경로는 폴백으로 남겨 둔다(과거 캐시 호환).
# ---------------------------------------------------------------------------
NAVER_MOBILE_INTEGRATION_URL = "https://m.stock.naver.com/api/stock/{ticker}/integration"
NAVER_MOBILE_INDUSTRY_URL = "https://m.stock.naver.com/api/stocks/industry"
_MOBILE_CACHE: dict[str, tuple[float, dict[str, Any]]] = {}
_INDUSTRY_MAP: dict[str, str] = {}


def parse_market_value_text(text: str) -> float | None:
    """'1,256조 9,499억' / '8,000억' / '3조' -> 억원."""
    raw = str(text or "").replace(",", "").strip()
    if not raw:
        return None
    jo = re.search(r"(\d+)\s*조", raw)
    eok = re.search(r"(\d+)\s*억", raw)
    if not jo and not eok:
        return None
    return float((int(jo.group(1)) * 10000 if jo else 0) + (int(eok.group(1)) if eok else 0))


def fetch_naver_mobile(ticker: str, quality_rows: list[dict[str, Any]], ttl: float = 300.0) -> dict[str, Any]:
    cached = _MOBILE_CACHE.get(ticker)
    if cached and time.time() - cached[0] < ttl:
        return cached[1]
    data: dict[str, Any] = {}
    try:
        data = get_json(NAVER_MOBILE_INTEGRATION_URL.format(ticker=ticker), {}, timeout=12)
        time.sleep(0.15)
    except Exception as exc:
        quality_rows.append(dq(ticker, "네이버증권(모바일)", "WARN", f"네이버 모바일 API 요청 실패: {exc}", True, NAVER_MOBILE_INTEGRATION_URL.format(ticker=ticker)))
    _MOBILE_CACHE[ticker] = (time.time(), data if isinstance(data, dict) else {})
    return _MOBILE_CACHE[ticker][1]


def naver_industry_name(code: str) -> str:
    if not _INDUSTRY_MAP:
        try:
            payload = get_json(NAVER_MOBILE_INDUSTRY_URL, {"page": 1, "pageSize": 100}, timeout=12)
            for group in payload.get("groups") or []:
                _INDUSTRY_MAP[str(group.get("no"))] = str(group.get("name") or "").strip()
        except Exception:
            pass
    return _INDUSTRY_MAP.get(str(code), "")


def mobile_market_cap(data: dict[str, Any]) -> float | None:
    for info in data.get("totalInfos") or []:
        if info.get("code") == "marketValue":
            return parse_market_value_text(info.get("value"))
    return None


def mobile_sector(data: dict[str, Any]) -> str:
    return naver_industry_name(data.get("industryCode")) if data.get("industryCode") else ""


def fetch_business_model(api_key: str | None, disclosure: Disclosure, offline: bool, quality_rows: list[dict[str, Any]]) -> str:
    if offline:
        return "확인 필요"

    sector = mobile_sector(fetch_naver_mobile(disclosure.stock_code, quality_rows))
    if not sector:
        naver_html = fetch_naver_html(disclosure.stock_code, quality_rows)
        sector = parse_naver_sector(naver_html)
    # 2026-08-03 사용자 확정: 업종은 **네이버증권 업종명을 무가공 그대로** 쓴다.
    # 키워드 매핑(map_business_model)은 오분류 사고 5회(부품·소재·NC·항공 등·엔터→게임/콘텐츠)로 전면 폐지.
    # map_business_model 함수는 호환성 위해 남겨두되 여기서 호출하지 않는다.
    if sector:
        return sector

    if api_key and disclosure.corp_code:
        try:
            payload = get_json(f"{DART_API_BASE}/company.json", {"crtfc_key": api_key, "corp_code": disclosure.corp_code}, timeout=15)
            text = " ".join(str(payload.get(key) or "") for key in ["induty_code", "hm_url", "corp_name"])
            mapped = map_business_model(text, disclosure.corp_name)
            if mapped != "확인 필요":
                return mapped
        except Exception as exc:
            quality_rows.append(dq(disclosure.corp_name, "업종", "WARN", f"DART 회사개요 조회 실패: {exc}", True, disclosure.link))

    # 매핑 규칙에 없더라도 네이버 동일업종비교의 업종명이 있으면 그대로 사용
    if sector:
        return sector
    quality_rows.append(dq(disclosure.corp_name, "업종", "WARN", "사업모델 분류 확인 필요", True, disclosure.link))
    return "확인 필요"

def fetch_market_cap(ticker: str, offline: bool, quality_rows: list[dict[str, Any]]) -> float | None:
    if offline:
        return None
    mobile_cap = mobile_market_cap(fetch_naver_mobile(ticker, quality_rows, ttl=60.0))
    if mobile_cap is not None:
        return mobile_cap
    html = fetch_naver_html(ticker, quality_rows, force_refresh=True)
    if not html:
        return None
    market_cap = parse_market_cap_from_html(html)
    if market_cap is None:
        quality_rows.append(dq(ticker, "시가총액", "WARN", "네이버증권 시가총액 파싱 실패", True, NAVER_ITEM_URL.format(ticker=ticker)))
    return market_cap

def fetch_naver_html(ticker: str, quality_rows: list[dict[str, Any]], force_refresh: bool = False) -> str:
    cache_path = CACHE_DIR / "naver" / f"{ticker}.html"
    if not force_refresh and cache_path.exists() and (time.time() - cache_path.stat().st_mtime) < 24 * 3600:
        return cache_path.read_text(encoding="utf-8", errors="replace")
    url = NAVER_ITEM_URL.format(ticker=ticker)
    try:
        html = get_text(url, timeout=12)
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(html, encoding="utf-8")
        time.sleep(0.3)
        return html
    except Exception as exc:
        quality_rows.append(dq(ticker, "네이버증권", "WARN", f"네이버증권 요청 실패: {exc}", True, url))
        return ""

def parse_market_cap_from_html(html: str) -> float | None:
    # 1순위: 네이버 종목 헤더의 _market_sum 필드(표시 종목의 시가총액, 조/억). 가장 신뢰 가능.
    ms = re.search(r"id=[\"']_market_sum[\"'][^>]*>(.*?)</", html, re.S)
    if ms:
        raw = re.sub(r"<[^>]+>", "", ms.group(1)).replace(",", "")
        jo = re.search(r"(\d+)\s*조", raw)
        rest = re.sub(r"\d+\s*조", "", raw)
        eok = re.search(r"(\d+)", rest)
        total = (int(jo.group(1)) * 10000 if jo else 0) + (int(eok.group(1)) if eok else 0)
        if total:
            return float(total)
    try:
        tables = pd.read_html(StringIO(html))
        for table in tables:
            text = table.to_string()
            if "시가총액" not in text:
                continue
            indexed_value = parse_market_cap_from_peer_table(table)
            if indexed_value is not None:
                return indexed_value
            for _, row in table.iterrows():
                row_text = " ".join(str(value) for value in row.tolist())
                if "시가총액" in row_text:
                    value = parse_market_cap_text(row_text)
                    if value is not None:
                        return value
    except Exception:
        pass
    match = re.search(r"시가총액.*?([\d,]+)\s*억", html, re.S)
    if match:
        return parse_amount_to_eok(match.group(1))
    dynamic_match = re.search(r"(?:market_sum|marketSum|sise_market_sum).*?([\d,]+)\s*억", html, re.S | re.I)
    if dynamic_match:
        return parse_amount_to_eok(dynamic_match.group(1))
    meta_match = re.search(r"시가총액[^0-9]{0,80}([\d,]+)", html, re.S)
    return parse_amount_to_eok(meta_match.group(1)) if meta_match else None

def parse_market_cap_from_peer_table(table: pd.DataFrame) -> float | None:
    if table.empty or table.shape[1] < 2:
        return None
    for _, row in table.iterrows():
        label = compact(row.iloc[0])
        if label == "시가총액(억)":
            return parse_amount_to_eok(row.iloc[1])
    return None

def parse_market_cap_text(text: str) -> float | None:
    trillion_match = re.search(r"([\d,]+)\s*조\s*([\d,]+)?\s*억", text)
    if trillion_match:
        trillion = parse_amount_to_eok(trillion_match.group(1)) or 0
        eok = parse_amount_to_eok(trillion_match.group(2)) if trillion_match.group(2) else 0
        return trillion * 10000 + (eok or 0)
    match = re.search(r"([\d,]+)\s*억", text)
    return parse_amount_to_eok(match.group(1)) if match else None

_SECTOR_LABEL_PLACEHOLDERS = {"PER", "등락률", "동일업종", "동일업종 PER", "동일업종 등락률", "업종명"}

def parse_naver_sector(html: str) -> str:
    if not html:
        return ""
    text = clean_html(html)
    for pattern in [
        r"sise_group_detail\.naver\?type=upjong[^>]*>\s*([^<]{2,40})\s*</a>",
        r"업종명\s*[:：]\s*([가-힣A-Za-z0-9·/&\-\s]{2,40}?)(?:\s*[|)]|\s*재무정보|\s*$)",
        r"동일업종비교[^가-힣A-Za-z0-9]+([가-힣A-Za-z0-9·/&\-\s]{2,40}?)(?:\s|$)",
        r"업종명\s+([가-힣A-Za-z0-9·/&\-\s]{2,40}?)(?:\s|$)",
        r"업종명\s*</th>\s*<td[^>]*>(.*?)</td>",
        r"업종\s*</th>\s*<td[^>]*>(.*?)</td>",
        r"업종</span>\s*<em[^>]*>(.*?)</em>",
    ]:
        match = re.search(pattern, html, re.I | re.S) or re.search(pattern, text, re.I)
        if match:
            value = clean_text(re.sub(r"<[^>]+>", "", match.group(1)))
            value = re.sub(r"^(업종명|동일업종비교)\s*", "", value)
            value = re.split(r"\s{2,}|전일대비|등락률|현재가|시가총액", value)[0].strip(" -:：|")
            # 2026-08-05 발견: 일부 종목 페이지는 "동일업종 PER"/"동일업종 등락률" 표 헤더의
            # <a> 태그가 sise_group_detail 링크를 감싸고 있어(진짜 업종명이 아님) 첫 패턴이
            # 그 라벨 문구를 그대로 캡처하는 버그가 있었다(예: 롯데하이마트→"동일업종 PER" 오기재).
            # 그런 플레이스홀더 값은 버리고 다음 패턴/폴백으로 넘어간다.
            if value and value not in _SECTOR_LABEL_PLACEHOLDERS:
                return value

    # 위 패턴이 전부 플레이스홀더만 잡았거나 실패한 경우의 최종 폴백:
    # sise_group_detail 링크의 no= 파라미터(업종 그룹 코드)로 업종 비교 페이지를 직접 열어
    # <title>업종명 : Npay 증권</title>에서 진짜 업종명을 얻는다.
    group_match = re.search(r"sise_group_detail\.naver\?type=upjong&(?:amp;)?no=(\d+)", html, re.I)
    if group_match:
        try:
            group_html = get_text(
                f"https://finance.naver.com/sise/sise_group_detail.naver?type=upjong&no={group_match.group(1)}",
                timeout=10,
            )
            title_match = re.search(r"<title>\s*(.*?)\s*:\s*Npay", group_html, re.I | re.S)
            if title_match:
                value = clean_text(title_match.group(1))
                if value and value not in _SECTOR_LABEL_PLACEHOLDERS:
                    return value
        except Exception:
            pass
    return ""

def map_business_model(text: str, corp_name: str) -> str:
    source = f"{text} {corp_name}"
    rules = [
        ("반도체/스마트폰", ["삼성전자"]),
        ("반도체", ["SK하이닉스"]),
        ("2차전지", ["삼성SDI"]),
        ("2차전지", ["LG에너지솔루션"]),
        ("가전/전장", ["LG전자"]),
        ("반도체 장비", ["반도체장비", "반도체 장비", "반도체와반도체장비"]),
        # 범용 키워드 '소재'는 회사명(예: 코스모신소재)에 오매칭되므로 제거. 2차전지 소재를 먼저 둔다.
        ("2차전지 소재", ["양극재", "음극재", "양극활물질", "코스모신소재", "에코프로비엠", "엘앤에프"]),
        ("반도체 소재", ["반도체소재"]),
        # 범용 키워드 'IP'는 회사명(예: IPARK현대산업개발)에 오매칭되므로 제거.
        ("팹리스/IP", ["팹리스"]),
        ("OSAT/검사장비", ["후공정", "검사", "테스트", "OSAT"]),
        ("PCB/기판", ["PCB", "기판"]),
        ("디스플레이 장비", ["디스플레이", "OLED", "LCD"]),
        ("클린룸/유틸리티", ["클린룸"]),
        # '자동차부품'이 범용 키워드 '부품'(IT 부품)에 먼저 걸리지 않도록 자동차를 앞에 둔다
        ("자동차 부품", ["자동차부품"]),
        ("IT 부품", ["전자부품", "IT부품", "부품"]),
        ("바이오/의료기기", ["바이오", "의료", "제약"]),
        ("소프트웨어", ["소프트웨어", "SW", "플랫폼"]),
        ("자동차 부품", ["자동차부품"]),
        ("화장품", ["화장품", "코스메틱"]),
        ("조선기자재", ["조선기자재"]),
        ("2차전지 장비", ["2차전지", "이차전지", "배터리"]),
        ("건설/인프라", ["건설", "인프라", "토목", "플랜트"]),
        ("항공/물류", ["항공사", "항공화물", "공항", "물류", "운송", "해상운송"]),
        ("소프트웨어", ["KTIS", "케이티아이에스", "컨택센터", "고객센터", "BPO", "ICT"]),
        ("철강/소재", ["제강", "철강", "금속", "스틸"]),
        ("금융", ["증권", "금융", "은행", "보험", "캐피탈"]),
        ("유통/소비재", ["이마트", "신세계", "유통", "백화점", "식품", "푸드", "씨푸드"]),
        ("에너지/유틸리티", ["전력", "가스", "에너지", "발전"]),
        ("화학/소재", ["케미칼", "화학", "인더스트리"]),
        ("게임/콘텐츠", ["게임", "콘텐츠", "엔터", "SHIFTUP", "시프트업", "NC", "NHN"]),
        ("통신장비", ["통신장비", "네트워크장비", "5G장비"]),
        ("전기전자", ["전자장비와기기", "전자장비", "전자기기", "전기제품", "전기전자"]),
        ("기계/장비", ["기계", "기계장비", "산업기계"]),
        ("의류/패션", ["섬유", "의류", "신발", "호화품", "패션"]),
        ("생활용품", ["가정용품", "생활용품"]),
        ("호텔/레저", ["호텔", "레저", "카지노", "관광"]),
        ("지주회사", ["지주"]),
    ]
    compacted = compact(source).upper()
    for label, keywords in rules:
        if any(compact(keyword).upper() in compacted for keyword in keywords):
            return label
    return "확인 필요"
