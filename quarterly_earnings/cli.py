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

from .constants import CONFIG_DIR, DATA_DIR, REPORT_DIR, set_active_quarter
from .models import Disclosure
from .util import dq, write_csv, write_json
from .runsetup import configure_logging, ensure_dirs, load_env_api_key, quarter_period, resolve_run_date, run_log_path, should_run_scheduled_date, write_last_run_status
from .disclosures import dedupe_disclosures, fetch_dart_disclosures, filter_common_stocks, filter_report_names
from .rules import calculate_results
from .stats import sort_key
from .workbook import build_workbook, calculation_row, disclosure_to_row, make_summary, print_run_summary

def main(argv: list[str] | None = None) -> int:
    run_started = datetime.now()
    args: argparse.Namespace | None = None
    logger: logging.Logger | None = None
    log_path: Path | None = None
    try:
        args = parse_args(argv)
        args.requested_end = args.end
        args.end = resolve_run_date(args.end)
        ensure_dirs()
        log_path = run_log_path(args.quarter, args.end)
        logger = configure_logging(args.quarter, args.end)

        api_key = load_env_api_key(CONFIG_DIR / ".env")
        set_active_quarter(args.quarter)
        period = quarter_period(args.quarter)
        if period is None:
            raise ValueError(f"Unsupported quarter: {args.quarter}. Use a Q2/Q3 quarter such as 2026Q2 or 2026Q3.")

        logger.info("Quarterly earnings filter started start=%s end=%s quarter=%s market=%s", args.start, args.end, args.quarter, args.market)
        # 2026-08-13: 이 하드코딩된 상한("2026-08-12")이 실제 운영일을 못 따라가서 오늘(08-13) 실행이
        # "after_end_date"로 통째로 스킵되는 사고가 났다. 이 같은 값이 Windows 작업 스케줄러가 매일 돌리는
        # run_quarterly_earnings_filter_2026Q2.ps1(--end today)도 똑같이 막고 있었다 — 즉 오늘부터 자동갱신
        # 자체가 조용히 안 돌 뻔했음. 요일/공휴일 체크(should_run_scheduled_date 자체)는 계속 유효하므로
        # 살려두되, 매일 손으로 늘려야 하는 상한 날짜는 2026Q2 실적시즌이 끝나는 시점까지 넉넉히 잡아
        # 더 이상 발이 걸리지 않게 한다.
        should_run, skip_reason = should_run_scheduled_date(args.end, args.start, "2026-12-31")
        if not should_run:
            logger.info("Skipping quarterly earnings filter run_date=%s reason=%s", args.end, skip_reason)
            write_last_run_status(args, run_started, "skipped", skip_reason, 0, 0, None, log_path)
            print(json.dumps({"skipped": True, "run_date": args.end, "reason": skip_reason}, ensure_ascii=False, indent=2))
            return 0

        quality_rows: list[dict[str, Any]] = []
        excluded_rows: list[dict[str, Any]] = []
        quarter_rows: list[dict[str, Any]] = []

        if not api_key:
            quality_rows.append(dq("", "DART_API_KEY", "ERROR", "DART_API_KEY가 .env에 없어 DART 수집을 건너뜀", True, ""))
            disclosures: list[Disclosure] = []
        else:
            disclosures = fetch_dart_disclosures(api_key, args.start, args.end, args.market, logger, quality_rows, args.offline)

        filtered_disclosures = filter_report_names(disclosures)
        if args.limit is not None:
            filtered_disclosures = filtered_disclosures[: args.limit]
        deduped_disclosures = dedupe_disclosures(filtered_disclosures, api_key, args.offline, quality_rows)
        common_disclosures = filter_common_stocks(deduped_disclosures, excluded_rows)

        raw_rows = [disclosure_to_row(item) for item in filtered_disclosures]
        calc_results = calculate_results(api_key, common_disclosures, args, logger, quality_rows, quarter_rows)
        passed_results = [result for result in calc_results if result.passed]
        passed_results.sort(key=sort_key)

        output_path = REPORT_DIR / f"{args.quarter}_분기실적_조건필터링_{args.end}.xlsx"
        latest_output_path = REPORT_DIR / f"{args.quarter}_분기실적_조건필터링_누적.xlsx"
        build_workbook(
            output_path=output_path,
            args=args,
            run_started=run_started,
            raw_rows=raw_rows,
            calc_results=calc_results,
            passed_results=passed_results,
            excluded_rows=excluded_rows,
            quality_rows=quality_rows,
            quarter_rows=quarter_rows,
        )
        if latest_output_path != output_path:
            shutil.copy2(output_path, latest_output_path)

        write_csv(DATA_DIR / f"{args.quarter}_raw_disclosures_{args.end}.csv", raw_rows)
        write_csv(DATA_DIR / f"{args.quarter}_calculation_detail_{args.end}.csv", [calculation_row(r) for r in calc_results])
        write_json(
            DATA_DIR / f"{args.quarter}_run_summary_{args.end}.json",
            make_summary(args, run_started, raw_rows, passed_results, excluded_rows, quality_rows),
        )

        write_last_run_status(args, run_started, "success", "completed", len(raw_rows), len(passed_results), latest_output_path, log_path)
        logger.info("Quarterly earnings filter finished output=%s latest=%s passed=%s", output_path, latest_output_path, len(passed_results))
        print_run_summary(args, latest_output_path, raw_rows, passed_results, excluded_rows, quality_rows)
        return 0
    except Exception as exc:
        ensure_dirs()
        if logger:
            logger.exception("Quarterly earnings filter failed")
        fallback_args = args or argparse.Namespace(start="", end="", requested_end="", quarter="", market="")
        write_last_run_status(fallback_args, run_started, "failed", str(exc), 0, 0, None, log_path)
        print(json.dumps({"status": "failed", "reason": str(exc)}, ensure_ascii=False, indent=2))
        return 1

def parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="분기 실적 시즌(2Q/3Q) 조건 필터링 자동화")
    parser.add_argument("--start", required=True, help="검색 시작일 YYYY-MM-DD")
    parser.add_argument("--end", required=True, help="검색 종료일/실행일 YYYY-MM-DD")
    parser.add_argument("--quarter", default="2026Q2", help="분석 분기, 기본값 2026Q2")
    parser.add_argument("--market", default="all", choices=["all", "kospi", "kosdaq"], help="대상 시장")
    parser.add_argument("--offline", action="store_true", help="네트워크 호출 없이 빈 리포트 구조만 생성")
    parser.add_argument("--limit", type=int, default=None, help="리허설용 최대 공시 처리 건수")
    return parser.parse_args(argv)
