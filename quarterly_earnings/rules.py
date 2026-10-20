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

from .models import CalcResult, Disclosure
from .util import dq, growth, is_loss_turnaround, is_turnaround
from .disclosures import _extract_report_period_days
from .financials import fetch_disclosure_text, fetch_periodic_financials, parse_financial_values
from .enrichment import fetch_business_model, fetch_market_cap
from .workbook import display_value
from .constants import QUARTER_LABELS as L

def calculate_results(
    api_key: str | None,
    disclosures: list[Disclosure],
    args: argparse.Namespace,
    logger: logging.Logger,
    quality_rows: list[dict[str, Any]],
    quarter_rows: list[dict[str, Any]],
) -> list[CalcResult]:
    results = []
    for index, disclosure in enumerate(disclosures):
        if index > 0:
            time.sleep(0.35)
        result = CalcResult(disclosure=disclosure)
        monthly_skip = False
        result.business_model = fetch_business_model(api_key, disclosure, args.offline, quality_rows)
        result.market_cap_eok = fetch_market_cap(disclosure.stock_code, args.offline, quality_rows)
        result.market_cap_status = "OK" if result.market_cap_eok is not None else "확인 필요"

        if disclosure.report_type in ("반기보고서", "분기보고서"):
            # 정식 반기/분기보고서는 본문이 방대해 텍스트 파싱 대신 정식 재무제표 API로 수집
            parsed = fetch_periodic_financials(
                api_key, disclosure.corp_code, args.quarter, disclosure.rcp_no,
                args.offline, quality_rows, disclosure.corp_name, disclosure.link,
                report_type=disclosure.report_type,
            )
            result.source_text_status = "OK" if parsed else "확인 필요"
        else:
            source_text = fetch_disclosure_text(api_key, disclosure.rcp_no, args.offline, quality_rows)
            result.source_text_status = "OK" if source_text else "확인 필요"
            parsed = parse_financial_values(source_text)
            # 2026-10-19: 월간 잠정(실적기간 ≤45일)은 분기 판정 대상이 아님 — 월 단위 값이 분기 슬롯에 들어가
            # QoQ/YoY가 엉뚱하게 계산(신세계·광주신세계가 8월→9월 월매출 증가로 오통과, 오리온 992 반복)되는 걸 막는다.
            period_days = _extract_report_period_days(source_text) if source_text else None
            if disclosure.report_type == "잠정실적" and period_days is not None and period_days <= 45:
                parsed = {}
                monthly_skip = True
                quality_rows.append(dq(disclosure.corp_name, "월간 잠정실적", "INFO",
                                       f"실적기간 {period_days}일짜리 월간 공시 — 분기실적이 아니라 판정 제외", False, disclosure.link))
        apply_financial_rules(result, parsed, disclosure.report_type, quarter_rows, quality_rows)
        if monthly_skip:
            result.calculation_method = "월간 잠정실적 공시(실적기간 45일 이하) — 분기 판정 제외"
        compute_metrics(result, quality_rows)
        results.append(result)
        logger.info("Processed %s %s passed=%s", disclosure.stock_code, disclosure.corp_name, result.passed)
    return results

def apply_financial_rules(
    result: CalcResult,
    parsed: dict[str, float | None],
    report_type: str,
    quarter_rows: list[dict[str, Any]],
    quality_rows: list[dict[str, Any]],
) -> None:
    for key, value in parsed.items():
        if hasattr(result, key):
            setattr(result, key, value)

    if report_type == "잠정실적":
        result.calculation_method = "공시 안의 당해/전기/전년동기/누계 실적을 우선 사용"
    elif report_type in ("반기보고서", "분기보고서"):
        result.calculation_method = "DART 정식 반기/분기보고서 재무제표(fnlttSinglAcntAll) 수집, 당분기단독=3개월값(없으면 누적-전분기 차감)"
        if result.q2_revenue is None and result.h1_revenue is not None and result.q1_revenue is not None:
            result.q2_revenue = result.h1_revenue - result.q1_revenue
        if result.q2_operating_income is None and result.h1_operating_income is not None and result.q1_operating_income is not None:
            result.q2_operating_income = result.h1_operating_income - result.q1_operating_income
    elif report_type == "손익구조변동":
        result.calculation_method = f"공시 누적값에서 전분기 값을 차감해 {L['cur']} 단독값 계산"
        if result.h1_revenue is not None and result.q1_revenue is not None:
            result.q2_revenue = result.h1_revenue - result.q1_revenue
        if result.h1_operating_income is not None and result.q1_operating_income is not None:
            result.q2_operating_income = result.h1_operating_income - result.q1_operating_income
        if result.y2025_h1_revenue is not None:
            if result.y2025_q1_revenue is not None:
                result.y2025_q2_revenue = result.y2025_h1_revenue - result.y2025_q1_revenue
            elif result.y2025_q2_revenue is None:
                quality_rows.append(dq(result.disclosure.corp_name, f"{L['yoy']} 매출", "WARN", f"{L['yoycum']}-{L['prev1']} 차감에 필요한 {L['prev1']} 값 확인 필요", True, result.disclosure.link))
        if result.y2025_h1_operating_income is not None:
            if result.y2025_q1_operating_income is not None:
                result.y2025_q2_operating_income = result.y2025_h1_operating_income - result.y2025_q1_operating_income
            elif result.y2025_q2_operating_income is None:
                quality_rows.append(dq(result.disclosure.corp_name, f"{L['yoy']} 영업이익", "WARN", f"{L['yoycum']}-{L['prev1']} 차감에 필요한 {L['prev1']} 값 확인 필요", True, result.disclosure.link))
    else:
        result.calculation_method = "공시 유형 확인 필요"

    for label, value, source in [
        (f"{L['cur']} 매출", result.q2_revenue, result.disclosure.report_nm),
        (f"{L['cur']} 영업이익", result.q2_operating_income, result.disclosure.report_nm),
        (f"{L['cum']} 매출", result.h1_revenue, result.disclosure.report_nm),
        (f"{L['cum']} 영업이익", result.h1_operating_income, result.disclosure.report_nm),
        (f"{L['prev']} 매출", result.q1_revenue, result.disclosure.report_nm),
        (f"{L['prev']} 영업이익", result.q1_operating_income, result.disclosure.report_nm),
        (f"{L['yoy']} 매출", result.y2025_q2_revenue, result.disclosure.report_nm),
        (f"{L['yoy']} 영업이익", result.y2025_q2_operating_income, result.disclosure.report_nm),
        (f"{L['prev1']} 매출", result.y2025_q1_revenue, result.disclosure.report_nm),
        (f"{L['prev1']} 영업이익", result.y2025_q1_operating_income, result.disclosure.report_nm),
        (f"{L['yoycum']} 매출", result.y2025_h1_revenue, result.disclosure.report_nm),
        (f"{L['yoycum']} 영업이익", result.y2025_h1_operating_income, result.disclosure.report_nm),
    ]:
        quarter_rows.append(
            {
                "종목명": result.disclosure.corp_name,
                "종목코드": result.disclosure.stock_code,
                "데이터 항목": label,
                "값": display_value(value),
                "수집 원천": source,
                "출처 링크": result.disclosure.link,
            }
        )

def compute_metrics(result: CalcResult, quality_rows: list[dict[str, Any]]) -> None:
    result.qoq_revenue_growth = growth(result.q2_revenue, result.q1_revenue)
    result.qoq_op_growth = growth(result.q2_operating_income, result.q1_operating_income)
    result.yoy_revenue_growth = growth(result.q2_revenue, result.y2025_q2_revenue)
    result.yoy_op_growth = growth(result.q2_operating_income, result.y2025_q2_operating_income)
    result.yoy_cum_revenue_growth = growth(result.h1_revenue, result.y2025_h1_revenue)
    result.yoy_cum_op_growth = growth(result.h1_operating_income, result.y2025_h1_operating_income)

    result.qoq_turnaround = is_turnaround(result.q2_operating_income, result.q1_operating_income)
    result.yoy_turnaround = is_turnaround(result.q2_operating_income, result.y2025_q2_operating_income)
    result.qoq_loss_turnaround = is_loss_turnaround(result.q2_operating_income, result.q1_operating_income)
    result.yoy_loss_turnaround = is_loss_turnaround(result.q2_operating_income, result.y2025_q2_operating_income)

    checks = [
        ("QoQ 매출 15% 이상", result.qoq_revenue_growth),
        ("QoQ 영익 15% 이상", result.qoq_op_growth),
        ("YoY 매출 15% 이상", result.yoy_revenue_growth),
        ("YoY 영익 15% 이상", result.yoy_op_growth),
    ]
    for label, value in checks:
        if value is not None and value >= 15:
            result.pass_reasons.append(label)
    if result.qoq_turnaround:
        result.pass_reasons.append("전분기 대비 영업이익 흑자전환")
    if result.yoy_turnaround:
        result.pass_reasons.append("전년동기 대비 영업이익 흑자전환")

    result.per_basis_op_income = result.h1_operating_income
    if result.market_cap_eok is not None and result.per_basis_op_income is not None and result.per_basis_op_income > 0:
        result.per = result.market_cap_eok / result.per_basis_op_income
    else:
        result.per = None
        if result.market_cap_eok is None:
            quality_rows.append(dq(result.disclosure.corp_name, "PER", "WARN", "시가총액 확인 필요로 PER 계산 불가", True, result.disclosure.link))
        elif result.per_basis_op_income is None:
            quality_rows.append(dq(result.disclosure.corp_name, "PER", "WARN", "누적 영업이익 확인 필요로 PER 계산 불가", True, result.disclosure.link))
        elif result.per_basis_op_income <= 0:
            quality_rows.append(dq(result.disclosure.corp_name, "PER", "INFO", "누적 영업이익이 0 이하라 PER N/A 처리", False, result.disclosure.link))

    required = [
        (f"{L['cur']} 매출", result.q2_revenue),
        (f"{L['cur']} 영업이익", result.q2_operating_income),
        (f"{L['cum']} 영업이익", result.h1_operating_income),
        (f"{L['prev']} 매출", result.q1_revenue),
        (f"{L['prev']} 영업이익", result.q1_operating_income),
        (f"{L['yoy']} 매출", result.y2025_q2_revenue),
        (f"{L['yoy']} 영업이익", result.y2025_q2_operating_income),
    ]
    for label, value in required:
        if value is None:
            quality_rows.append(dq(result.disclosure.corp_name, label, "WARN", f"{label} 확인 필요", True, result.disclosure.link))
