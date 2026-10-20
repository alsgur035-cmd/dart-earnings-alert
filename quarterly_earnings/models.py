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

from .constants import DART_VIEWER_URL
from .util import compact

@dataclass
class Disclosure:
    rcp_no: str
    corp_code: str
    stock_code: str
    corp_name: str
    report_nm: str
    rcept_dt: str
    corp_cls: str
    flr_nm: str = ""
    rm: str = ""

    @property
    def link(self) -> str:
        return DART_VIEWER_URL.format(rcp_no=self.rcp_no)

    @property
    def report_type(self) -> str:
        normalized = compact(self.report_nm)
        if "잠정" in normalized and "실적" in normalized:
            return "잠정실적"
        if "손익구조" in normalized or "매출액또는손익구조" in normalized:
            return "손익구조변동"
        if "반기보고서" in normalized:
            return "반기보고서"
        if "분기보고서" in normalized:
            return "분기보고서"
        return "기타"

    @property
    def is_correction(self) -> bool:
        return "정정" in self.report_nm

@dataclass
class CalcResult:
    disclosure: Disclosure
    business_model: str = "확인 필요"
    market_cap_eok: float | None = None
    market_cap_status: str = "확인 필요"
    q2_revenue: float | None = None
    q2_operating_income: float | None = None
    h1_revenue: float | None = None
    h1_operating_income: float | None = None
    q1_revenue: float | None = None
    q1_operating_income: float | None = None
    y2025_q2_revenue: float | None = None
    y2025_q2_operating_income: float | None = None
    y2025_q1_revenue: float | None = None
    y2025_q1_operating_income: float | None = None
    y2025_h1_revenue: float | None = None
    y2025_h1_operating_income: float | None = None
    qoq_revenue_growth: float | None = None
    qoq_op_growth: float | None = None
    yoy_revenue_growth: float | None = None
    yoy_op_growth: float | None = None
    yoy_cum_revenue_growth: float | None = None
    yoy_cum_op_growth: float | None = None
    qoq_turnaround: bool = False
    yoy_turnaround: bool = False
    qoq_loss_turnaround: bool = False
    yoy_loss_turnaround: bool = False
    pass_reasons: list[str] = field(default_factory=list)
    per: float | None = None
    per_basis_op_income: float | None = None
    calculation_method: str = ""
    source_text_status: str = "확인 필요"
    quality: list[dict[str, Any]] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return bool(self.pass_reasons)
