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

from .models import CalcResult

def sort_key(result: CalcResult) -> tuple[int, float, float, float]:
    return (
        0 if (result.qoq_turnaround or result.yoy_turnaround) else 1,
        -(result.yoy_op_growth if result.yoy_op_growth is not None else -10**12),
        -(result.qoq_op_growth if result.qoq_op_growth is not None else -10**12),
        -(result.market_cap_eok if result.market_cap_eok is not None else -1),
    )

def count_reason(results: list[CalcResult], token: str) -> int:
    return sum(1 for result in results if any(token in reason for reason in result.pass_reasons))
