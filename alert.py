"""DART 신규 실적 공시 감시 → 텔레그램 폰 알림.

GitHub Actions에서 몇 분마다 실행된다(.github/workflows/alert.yml).
공시 검색·원문 파싱·15% 조건 판정은 분기 잠정실적 파이프라인(quarterly_earnings 패키지 사본)을 그대로 쓴다.
이미 알린 공시는 state/seen.json의 접수번호로 기억해 두 번 보내지 않는다.

환경변수: DART_API_KEY, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID, QUARTER(기본 2026Q3)
로컬 시험: python alert.py --start 2026-10-19 --end 2026-10-20 --dry-run --state 임시경로.json
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from datetime import datetime, timedelta
from html import escape
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from quarterly_earnings import constants, disclosures as qd, financials as qf  # noqa: E402
from quarterly_earnings.models import CalcResult  # noqa: E402
from quarterly_earnings.rules import apply_financial_rules, compute_metrics  # noqa: E402

KST = ZoneInfo("Asia/Seoul")

# 원문 캐시를 이 폴더 안으로 돌린다(원래 경로는 PC의 01_결과물 폴더 기준이라 클라우드에선 엉뚱한 곳이 됨).
_CACHE = ROOT / ".cache"
constants.CACHE_DIR = qd.CACHE_DIR = qf.CACHE_DIR = _CACHE

# 정기보고서는 시즌에 하루 수백 건이라 조건 통과 종목만 알린다. 잠정실적·손익구조변동은 전부 알린다.
PERIODIC_TYPES = ("분기보고서", "반기보고서")
RUN_BUDGET_SEC = int(os.getenv("RUN_BUDGET_SEC", "360"))  # 워크플로 제한(10분) 안에서 끝나도록


def load_state(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {"seen": [], "date": ""}


def save_state(path: Path, state: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    # 최근 3000건만 보관(파일이 무한히 커지지 않게). 조회 범위가 이틀이라 충분하다.
    state["seen"] = state["seen"][-3000:]
    state["alerted"] = state.get("alerted", [])[-3000:]
    path.write_text(json.dumps(state, ensure_ascii=False, indent=0), encoding="utf-8")


def evaluate(api_key: str, disclosure, quarter: str) -> tuple[CalcResult | None, str]:
    """(판정결과, 건너뛴 사유). 판정 대상이 아니면 결과는 None."""
    quality: list = []
    result = CalcResult(disclosure=disclosure)
    if disclosure.report_type in PERIODIC_TYPES:
        parsed = qf.fetch_periodic_financials(
            api_key, disclosure.corp_code, quarter, disclosure.rcp_no, False, quality,
            disclosure.corp_name, disclosure.link, report_type=disclosure.report_type,
        )
    else:
        text = qf.fetch_disclosure_text(api_key, disclosure.rcp_no, False, quality)
        if not text:
            return None, "원문 수집 실패"
        days = qd._extract_report_period_days(text)
        if disclosure.report_type == "잠정실적" and days is not None and days <= 45:
            return None, f"월간 공시(실적기간 {days}일)"
        parsed = qf.parse_financial_values(text)
    apply_financial_rules(result, parsed, disclosure.report_type, [], quality)
    if result.q2_revenue is None and result.q2_operating_income is None:
        return None, "매출·영업이익 수치 없음"
    compute_metrics(result, quality)
    return result, ""


def fmt_amount(value: float | None) -> str:
    return "-" if value is None else f"{value:,.0f}억"


def fmt_growth(value: float | None, turnaround: bool) -> str:
    if turnaround:
        return "흑전"
    if value is None:
        return "-"
    return f"{value:+.1f}%"


def build_message(r: CalcResult, quarter_label: str) -> str:
    d = r.disclosure
    head = "✅" if r.passed else "⚪"
    kind = {"잠정실적": "잠정", "손익구조변동": "손익구조변동"}.get(d.report_type, d.report_type)
    lines = [
        f"{head} <b>{escape(d.corp_name)}</b> ({d.stock_code}) · {quarter_label} {kind}{' [정정]' if d.is_correction else ''}",
        f"매출 {fmt_amount(r.q2_revenue)}  QoQ {fmt_growth(r.qoq_revenue_growth, False)} / YoY {fmt_growth(r.yoy_revenue_growth, False)}",
        f"영익 {fmt_amount(r.q2_operating_income)}  QoQ {fmt_growth(r.qoq_op_growth, r.qoq_turnaround)} / YoY {fmt_growth(r.yoy_op_growth, r.yoy_turnaround)}",
        ("통과: " + ", ".join(r.pass_reasons)) if r.passed else "조건 미달",
        f'<a href="{d.link}">공시 원문</a> · {d.rcept_dt}',
    ]
    return "\n".join(lines)


def build_line(r: CalcResult) -> str:
    """정기보고서 묶음용 한 줄 요약."""
    d = r.disclosure
    return (
        f'• <a href="{d.link}">{escape(d.corp_name)}</a> '
        f"매출 {fmt_amount(r.q2_revenue)} (Q {fmt_growth(r.qoq_revenue_growth, False)}, Y {fmt_growth(r.yoy_revenue_growth, False)}) / "
        f"영익 {fmt_amount(r.q2_operating_income)} (Q {fmt_growth(r.qoq_op_growth, r.qoq_turnaround)}, Y {fmt_growth(r.yoy_op_growth, r.yoy_turnaround)})"
    )


def send_telegram(token: str, chat_id: str, text: str) -> None:
    data = urlencode({"chat_id": chat_id, "text": text, "parse_mode": "HTML", "disable_web_page_preview": "true"}).encode()
    req = Request(f"https://api.telegram.org/bot{token}/sendMessage", data=data)
    with urlopen(req, timeout=20) as resp:
        payload = json.loads(resp.read().decode())
    if not payload.get("ok"):
        raise RuntimeError(f"텔레그램 전송 실패: {payload}")


def main() -> int:
    now = datetime.now(KST)
    parser = argparse.ArgumentParser()
    # 최근 4일: 주말·연휴 사이(금요일 21시 이후 → 월요일 07시) 공시도 놓치지 않게. 이미 보낸 건 seen으로 걸러짐.
    parser.add_argument("--start", default=(now - timedelta(days=4)).strftime("%Y-%m-%d"))
    parser.add_argument("--end", default=now.strftime("%Y-%m-%d"))
    parser.add_argument("--state", default=str(ROOT / "state" / "seen.json"))
    parser.add_argument("--dry-run", action="store_true", help="텔레그램 대신 화면에 출력, 상태파일은 그대로 저장")
    parser.add_argument("--passed-only", action="store_true", default=os.getenv("PASSED_ONLY") == "1")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(message)s")
    log = logging.getLogger("alert")
    api_key = os.getenv("DART_API_KEY", "").strip()
    token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
    # 받는 사람 여러 명: 쉼표로 구분 (예: 8450555165,1234567890)
    chat_ids = [c.strip() for c in os.getenv("TELEGRAM_CHAT_ID", "").split(",") if c.strip()]
    quarter = os.getenv("QUARTER", "2026Q3").strip()
    quarter_label = f"{quarter[-1]}Q{quarter[2:4]}"  # 2026Q3 → 3Q26
    if not api_key:
        log.error("DART_API_KEY 없음")
        return 1
    if not args.dry_run and not (token and chat_ids):
        log.error("TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID 없음")
        return 1

    constants.set_active_quarter(quarter)
    state_path = Path(args.state)
    state = load_state(state_path)
    seen = set(state["seen"])

    quality: list = []
    raw = qd.fetch_dart_disclosures(api_key, args.start, args.end, "all", log, quality, False)
    errors = [q for q in quality if q.get("상태") == "ERROR"]
    if not raw and errors:
        log.error("DART 공시검색 실패: %s", errors)
        return 1
    targets = qd.filter_common_stocks(qd.filter_report_names(raw), [])
    new = [d for d in targets if d.rcp_no not in seen]
    if os.getenv("INCLUDE_PERIODIC", "1") == "0":  # 정기보고서 알림 끄기 스위치
        new = [d for d in new if d.report_type not in PERIODIC_TYPES]
    # 잠정실적·손익구조변동을 먼저 처리(속보성), 정기보고서는 뒤로.
    # 같은 회사가 연결·별도를 따로 내면 연결 기준이 먼저 오게(아래에서 회사·날짜당 1회만 알림).
    new.sort(key=lambda d: (d.report_type in PERIODIC_TYPES, "연결" not in d.report_nm, d.rcp_no))
    alerted = set(state.setdefault("alerted", []))

    def mark_alerted(d) -> None:
        key = f"{d.stock_code}|{d.rcept_dt}"
        if key not in alerted:
            alerted.add(key)
            state["alerted"].append(key)
    log.info("공시 %d건 중 실적 대상 %d건, 신규 %d건", len(raw), len(targets), len(new))

    def deliver(message: str) -> None:
        if args.dry_run:
            print("-" * 40 + "\n" + message)
        else:
            # 한 사람이 실패(봇 차단 등)해도 나머지에게는 보낸다. 전원 실패일 때만 오류 → 다음 실행에서 재시도.
            failed = []
            for cid in chat_ids:
                try:
                    send_telegram(token, cid, message)
                except Exception as exc:
                    failed.append(cid)
                    log.warning("텔레그램 전송 실패 (받는 사람 %s번째): %s", chat_ids.index(cid) + 1, exc)
            if len(failed) == len(chat_ids):
                raise RuntimeError("모든 받는 사람에게 전송 실패")
            time.sleep(1)

    sent = 0
    digest: list[tuple[object, str]] = []  # 정기보고서 통과 종목 (공시, 한 줄 요약) — 실행 끝에 묶어서 보냄
    started = time.monotonic()
    try:
        for d in new:
            # 마감일(공시 수백 건)에는 한 번에 다 못 돈다. 나머지는 다음 실행이 이어서 처리.
            if time.monotonic() - started > RUN_BUDGET_SEC:
                log.info("시간 한도 도달 — 남은 %d건은 다음 실행에서 처리", len(new) - new.index(d))
                break
            # 정정공시는 숫자가 바뀌므로 다시 알린다. 그 외 같은 회사·같은 날 두 번째 공시(별도 기준 등)는 생략.
            if f"{d.stock_code}|{d.rcept_dt}" in alerted and not d.is_correction:
                log.info("같은 날 이미 알림, 생략: %s (%s)", d.corp_name, d.report_nm)
                state["seen"].append(d.rcp_no)
                continue
            try:
                result, skip = evaluate(api_key, d, quarter)
            except Exception as exc:  # 한 건이 실패해도 나머지는 계속, 다음 실행 때 재시도
                log.warning("판정 실패 %s %s: %s", d.corp_name, d.rcp_no, exc)
                continue
            if result is None:
                log.info("건너뜀 %s (%s): %s", d.corp_name, d.report_nm, skip)
            elif d.report_type in PERIODIC_TYPES:
                if result.passed:
                    digest.append((d, build_line(result)))
                    continue  # 묶음 전송 후에 '본 것'으로 기록
                log.info("정기보고서 조건 미달, 알림 생략: %s", d.corp_name)
            elif args.passed_only and not result.passed:
                log.info("조건 미달, 알림 생략: %s", d.corp_name)
            else:
                deliver(build_message(result, quarter_label))
                mark_alerted(d)
                sent += 1
            state["seen"].append(d.rcp_no)
            save_state(state_path, state)  # 한 건마다 저장 — 중간에 죽어도 중복 알림 없음

        # 정기보고서 통과 종목: 메시지 하나에 여러 종목(텔레그램 한도 4096자 안으로 나눔)
        chunk: list[tuple[object, str]] = []
        for i, item in enumerate(digest):
            chunk.append(item)
            body = "\n".join(m for _, m in chunk)
            if len(body) > 3300 or i == len(digest) - 1:
                deliver(f"📑 정기보고서 조건 통과 {len(chunk)}건 ({quarter_label})\n{body}")
                sent += len(chunk)
                for item_d, _ in chunk:  # 전송이 끝난 뒤에야 기록 — 전송 전에 죽으면 다음 실행이 다시 보냄
                    state["seen"].append(item_d.rcp_no)
                    mark_alerted(item_d)
                save_state(state_path, state)
                chunk = []
    finally:
        # 날짜는 하루 한 번 바뀌므로 공시가 없는 날에도 상태파일 커밋이 생긴다 → 저장소가 '활동 없음'으로 꺼지지 않음.
        state["date"] = now.strftime("%Y-%m-%d")
        save_state(state_path, state)
    log.info("알림 %d건 전송%s", sent, " (dry-run)" if args.dry_run else "")
    return 0


if __name__ == "__main__":
    sys.exit(main())
