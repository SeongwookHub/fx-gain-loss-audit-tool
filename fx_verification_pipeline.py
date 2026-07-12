# -*- coding: utf-8 -*-
"""
외환차손익 / 외화환산손익 검증 파이프라인 (1단계 스크리닝 + 2단계 증빙 OCR)
================================================================

[전체 구조]

  A. 외환차손익(결제 건, 분개장의 "결제" 라인)
     1단계: 전수 스크리닝 - 회사 적용환율 vs 그날 공식 매매기준율(수출입은행 API)
            괴리율이 5%를 넘는 건만 자동 추출 (전수조사가 불가능한 현실을 반영,
            "어떤 걸 표본으로 볼지"를 사람이 감으로 고르지 않고 수치 기준으로 자동 추출)
     2단계: 1단계에서 걸린 건만 증빙(은행 외환거래확인서/SWIFT/외화예금 거래명세표 등
            스캔 이미지)을 Claude Vision으로 읽어 실제 적용환율을 추출하고,
            그 환율로 재계산한 외환차손익을 회사 계상액과 최종 비교

  B. 외화환산손익(기말평가 건, 분개장의 "기말평가" 라인)
     결산일 하나의 공식 환율로 전수 재계산 (거래별로 증빙이 다를 이유가 없으므로
     OCR 2단계가 필요 없음 - 전수 자동 검증)

[사용 전 준비]
  1. data.go.kr에서 한국수출입은행 환율 API 인증키 발급 (EXIM_AUTH_KEY 환경변수)
  2. 증빙 이미지가 있다면 evidence/ 폴더에 "{거래ID}.png" 또는 "{거래ID}.jpg" 형태로 저장
     (예: evidence/TXN002.png) - 없으면 2단계는 "증빙 요청 필요"로 표시만 하고 넘어감
  3. Claude API 사용을 위해 ANTHROPIC_API_KEY 환경변수 설정 (2단계 OCR용)
  4. pip install anthropic pandas openpyxl requests --break-system-packages
"""

import os
import re
import json
import time
import base64
from datetime import datetime, timedelta

import pandas as pd
import requests

# ------------------------------------------------------------------
# 0. 설정
# ------------------------------------------------------------------

EXIM_AUTH_KEY = os.environ.get("EXIM_AUTH_KEY", "여기에_발급받은_인증키_입력")
EXIM_BASE_URL = "https://oapi.koreaexim.go.kr/site/program/financial/exchangeJSON"
# 참고: 2025.6.25부로 요청 URL 도메인이 www.koreaexim.go.kr -> oapi.koreaexim.go.kr로 변경됨.
# 기존 도메인(www.koreaexim.go.kr)은 점진적으로 종료 예정이라 응답이 없거나 타임아웃 날 수 있음.

TOLERANCE_PCT = 0.05          # 1단계 이상치 판단 기준: 괴리율 5%
GAIN_LOSS_TOLERANCE_KRW = 1000  # 재계산 금액과 회사 계상액의 허용 오차(원 단위 반올림 차이)
AMPT = float(os.environ.get("FX_AMPT", 3000000))  # 허용가능 오류금액(Tolerable Misstatement) - 감사팀 산정치로 교체

CUR_UNIT_MAP = {"USD": "USD", "JPY": "JPY(100)", "EUR": "EUR", "CNH": "CNH"}

EVIDENCE_DIR = "evidence"      # 증빙 이미지 폴더

_rate_cache = {}


# ------------------------------------------------------------------
# 1. 수출입은행 환율 API
# ------------------------------------------------------------------

def fetch_rates_for_date(date_str: str, _retries: int = 2) -> dict:
    """특정 날짜(YYYYMMDD)의 전체 통화 매매기준율표를 API에서 가져옴. RESULT 코드 체크 포함.
    일시적 타임아웃/연결 오류는 짧게 대기 후 재시도(최대 2회)."""
    if date_str in _rate_cache:
        return _rate_cache[date_str]

    params = {"authkey": EXIM_AUTH_KEY, "data": "AP01", "searchdate": date_str}

    for attempt in range(_retries + 1):
        try:
            resp = requests.get(EXIM_BASE_URL, params=params, timeout=10)
            resp.raise_for_status()
            data = resp.json()
            break
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
            if attempt < _retries:
                time.sleep(1.5 * (attempt + 1))  # 1.5초, 3초 순으로 대기 후 재시도
                continue
            raise

    if not data:
        # 주말/공휴일 등으로 데이터가 없는 경우 (빈 리스트) - 폴백 대상
        _rate_cache[date_str] = {}
        return {}

    rate_table = {}
    for row in data:
        result = str(row.get("result", ""))
        if result == "2":
            raise ValueError("수출입은행 API 오류(RESULT=2): DATA 코드 오류")
        elif result == "3":
            raise ValueError("수출입은행 API 오류(RESULT=3): 인증코드(authkey) 오류 - 키를 확인하세요")
        elif result == "4":
            raise ValueError("수출입은행 API 오류(RESULT=4): 일일 호출 제한 초과 - 내일 다시 시도하거나 캐시를 활용하세요")
        try:
            rate_table[row["cur_unit"]] = float(row["deal_bas_r"].replace(",", ""))
        except (KeyError, ValueError):
            continue

    _rate_cache[date_str] = rate_table
    time.sleep(0.15)  # 짧은 pacing 지연 - 짧은 시간에 요청이 몰려 서버가 지연/거부하는 것을 예방
    return rate_table


def get_official_rate(date_obj: datetime, cur_unit: str, max_lookback_days: int = 5) -> tuple:
    """해당일 공식 매매기준율. 주말/공휴일이면 직전 영업일로 폴백.
    반환값: (환율, 실제 적용된 날짜) - 폴백된 경우 실제 날짜를 알아야 리포트에 남길 수 있음"""
    for i in range(max_lookback_days + 1):
        d = date_obj - timedelta(days=i)
        rate_table = fetch_rates_for_date(d.strftime("%Y%m%d"))
        if cur_unit in rate_table:
            return rate_table[cur_unit], d
    raise ValueError(f"{date_obj.date()} 기준 {cur_unit} 환율을 찾을 수 없습니다 (휴장일 {max_lookback_days}일 초과)")


def normalize_rate(raw_rate: float, currency: str) -> float:
    """JPY는 100엔당 고시이므로 1엔당 단가로 환산."""
    return raw_rate / 100 if currency == "JPY" else raw_rate


# ------------------------------------------------------------------
# 2. 분개장 파싱 - 결제 건 / 기말평가 건 추출
# ------------------------------------------------------------------

def load_journal(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    df["일자"] = pd.to_datetime(df["일자"])
    return df


def extract_settlement_transactions(journal_df: pd.DataFrame) -> pd.DataFrame:
    """거래ID별로 '결제' 구분 라인 중 외화금액이 있는 라인(= 외화예금 또는 채무 상대계정)을
    골라 거래 단위로 요약. 회사가 계상한 외환차익/차손 금액도 같은 거래ID에서 찾아 붙임."""
    settle_rows = journal_df[(journal_df["구분"] == "결제") & (journal_df["외화금액"].notna())].copy()

    records = []
    for txn_id, group in settle_rows.groupby("전표번호(거래ID)"):
        # 채권/채무 결제 모두 '외화예금' 라인이 결제 시점 실제 자금 흐름의 기준
        cash_row = group[group["계정과목"] == "외화예금"]
        row = cash_row.iloc[0] if not cash_row.empty else group.iloc[0]
        settle_date = row["일자"]
        currency = row["통화"]
        fc_amount = row["외화금액"]

        # 적용환율 '컬럼값'을 그대로 믿지 않고, 실제 분개된 원화금액에서 환율을 역산.
        # (라벨링 자체가 틀렸거나 단위 착오가 있어도 실제 찍힌 금액 기준으로 비교하기 위함)
        booked_krw = row[["차변금액", "대변금액"]].fillna(0).abs().max()
        implied_rate = booked_krw / fc_amount if fc_amount else None

        # 같은 거래ID의 결제 라인 중 907(외환차익)/957(외환차손) 계정 금액을 찾음
        same_txn = journal_df[(journal_df["전표번호(거래ID)"] == txn_id) & (journal_df["구분"] == "결제")]
        gain_row = same_txn[same_txn["계정과목"] == "외환차익"]
        loss_row = same_txn[same_txn["계정과목"] == "외환차손"]

        booked_gain_loss = 0
        if not gain_row.empty:
            booked_gain_loss = gain_row.iloc[0][["차변금액", "대변금액"]].fillna(0).abs().max()
        elif not loss_row.empty:
            booked_gain_loss = -loss_row.iloc[0][["차변금액", "대변금액"]].fillna(0).abs().max()

        # 채권/채무 구분과 발생 시점 장부인식금액(occur_krw) - 합계 중요성 검증에 필요
        occur_rows = journal_df[(journal_df["전표번호(거래ID)"] == txn_id) &
                                 (journal_df["구분"] == "발생") &
                                 (journal_df["계정코드"].isin([108, 251]))]
        side = "채권" if not occur_rows.empty and occur_rows.iloc[0]["계정코드"] == 108 else "채무"
        occur_krw = occur_rows.iloc[0][["차변금액", "대변금액"]].fillna(0).abs().max() if not occur_rows.empty else None

        records.append({
            "거래ID": txn_id, "결제일": settle_date, "통화": currency, "구분": side,
            "외화금액": fc_amount, "결제원화금액": booked_krw, "발생원화금액": occur_krw,
            "회사적용환율(내재)": implied_rate,
            "회사계상_외환차익차손": booked_gain_loss,
        })

    return pd.DataFrame(records)


def extract_yearend_transactions(journal_df: pd.DataFrame) -> pd.DataFrame:
    """'기말평가' 구분 라인에서 거래ID별 재평가 내역 추출."""
    ye_rows = journal_df[(journal_df["구분"] == "기말평가") & (journal_df["외화금액"].notna())].copy()

    records = []
    for txn_id, group in ye_rows.groupby("전표번호(거래ID)"):
        row = group.iloc[0]
        records.append({
            "거래ID": txn_id, "결산일": row["일자"], "통화": row["통화"],
            "외화금액": row["외화금액"], "회사적용환율(기말)": row["적용환율"],
        })
    return pd.DataFrame(records)


def get_unsettled_yearend_candidates(journal_df: pd.DataFrame, schedule_df: pd.DataFrame) -> pd.DataFrame:
    """명세서(외화자산부채명세서) 기준 기말 미결제 건 중, 분개장에 기말평가 라인이
    아예 없는 거래를 찾아냄 (재평가 누락 탐지)."""
    outstanding = schedule_df[schedule_df["기말미결제외화잔액"] > 0]
    ye_done_ids = set(extract_yearend_transactions(journal_df)["거래ID"]) if not journal_df.empty else set()
    missing = outstanding[~outstanding["전표번호(거래ID)"].isin(ye_done_ids)]
    return missing


# ------------------------------------------------------------------
# 3. 1단계: 외환차손익 스크리닝
# ------------------------------------------------------------------

def screen_fx_settlements(settlements: pd.DataFrame) -> pd.DataFrame:
    """각 결제 건에 대해:
    1) 공식 매매기준율과의 '괴리율'을 계산해 5% 초과 건을 개별 플래그
    2) 공식환율 기준으로 재계산한 외환차손익과 회사 계상액의 '금액 차이(KRW)'도 함께 산출
       (개별로는 5% 이내라도, 이 금액 차이들을 합산해 중요성 검토를 하기 위함)
    회사측 환율은 실제 분개 금액에서 역산한 내재환율(원 단위, 정규화 불필요)을 사용."""
    results = []
    for _, row in settlements.iterrows():
        cur_unit = CUR_UNIT_MAP.get(row["통화"], row["통화"])
        official_raw, actual_date = get_official_rate(row["결제일"], cur_unit)
        official_rate = normalize_rate(official_raw, row["통화"])
        company_rate = row["회사적용환율(내재)"]

        deviation_pct = abs(company_rate - official_rate) / official_rate
        flagged = deviation_pct > TOLERANCE_PCT

        # 공식환율로 재계산한 결제원화금액 및 외환차손익 (부호는 booked_gain_loss와 동일 규약:
        # 채권은 (결제-발생)이 이익, 채무는 (발생-결제)가 이익)
        recalculated_settle_krw = row["외화금액"] * official_rate
        if row["구분"] == "채권":
            recalculated_gain_loss = recalculated_settle_krw - row["발생원화금액"]
        else:
            recalculated_gain_loss = row["발생원화금액"] - recalculated_settle_krw

        diff_krw = recalculated_gain_loss - row["회사계상_외환차익차손"]

        results.append({
            **row.to_dict(),
            "공식매매기준율": round(official_rate, 4),
            "공식환율기준일": actual_date.strftime("%Y-%m-%d"),
            "괴리율(%)": round(deviation_pct * 100, 2),
            "1차플래그": "이상치(정밀검증필요)" if flagged else "적정(스크리닝통과)",
            "재계산_외환차손익(공식환율기준)": round(recalculated_gain_loss),
            "회사계상액과의차이(KRW)": round(diff_krw),
        })
    return pd.DataFrame(results)


def check_aggregate_materiality(screened: pd.DataFrame, ampt: float) -> dict:
    """개별 건이 전부 5% 이내로 통과했더라도, 재계산액과 회사계상액의 차이를
    전체 합산했을 때 AMPT(허용가능 오류금액)를 넘는지 확인.
    순합계(넷)와 절대값합계(그로스) 둘 다 보여줌 - 방향이 서로 반대인 오류가
    상쇄되어 순합계는 작아 보여도 그로스 기준으로는 클 수 있기 때문."""
    net_sum = screened["회사계상액과의차이(KRW)"].sum()
    gross_sum = screened["회사계상액과의차이(KRW)"].abs().sum()
    breach = abs(net_sum) > ampt or gross_sum > ampt

    return {
        "AMPT": ampt,
        "순차이합계(KRW)": round(net_sum),
        "절대값차이합계(KRW)": round(gross_sum),
        "중요성초과여부": breach,
        "판정": "전체 재검토 필요(합계 중요성 초과)" if breach else "전체 적정(합계 중요성 이내)",
    }


# ------------------------------------------------------------------
# 3-보조. 결제일-환율기준일 불일치 탐지
# ------------------------------------------------------------------
#
# 왜 필요한가: 5% 괴리율 기준(크기 기준)은 "전월말 환율을 잘못 쓴" 것 같은 오류를
# 놓칠 수 있다 (원/달러가 한 달 새 5% 넘게 안 움직이면 개별 판정을 통과해버림 -
# TXN002 사례에서 실측). 그래서 크기가 아니라 "회사가 쓴 환율이 실제로는 다른
# 날짜의 공식 환율과 정확히 일치하는가"를 별도로 확인한다. 소액이라도 날짜
# 자체를 잘못 쓴 경우라면 이 체크에서 걸린다.

REF_DATE_MATCH_TOLERANCE = 0.0005  # 환율이 '일치'한다고 볼 허용오차(0.05%) - API 반올림 오차 흡수용


def _candidate_wrong_dates(settle_date: datetime) -> list:
    """실무에서 실제로 자주 발생하는 '기준일 착오' 패턴 후보만 골라서 반환.
    (임의로 40일을 다 훑는 대신, 흔한 실수 패턴만 확인 -> API 호출을 대폭 절감)"""
    candidates = []

    # 전월말(직전월 마지막 날)
    first_of_this_month = settle_date.replace(day=1)
    prev_month_end = first_of_this_month - timedelta(days=1)
    candidates.append(prev_month_end)

    # 전영업일들 (T-1 ~ T-3, 주말 포함해서 최대 5일 전까지)
    for i in range(1, 6):
        candidates.append(settle_date - timedelta(days=i))

    # 1주일 전, 2주일 전 (담당자가 "지난주 환율로" 착각하는 경우)
    candidates.append(settle_date - timedelta(days=7))
    candidates.append(settle_date - timedelta(days=14))

    # 중복 제거, 결제일 이후 날짜는 제외
    seen = set()
    result = []
    for d in candidates:
        key = d.strftime("%Y%m%d")
        if key not in seen and d < settle_date:
            seen.add(key)
            result.append(d)
    return result


def detect_reference_date_mismatch(settlements: pd.DataFrame) -> pd.DataFrame:
    """회사가 사용한 환율이 결제일이 아닌 다른 '흔히 착각하는' 날짜의 공식환율과
    더 정확히 일치하는지 확인. 일치하는 과거 날짜를 찾으면 '기준일 오류 의심'으로
    표시하고 그 날짜를 함께 보여준다. 5% 이내로 통과한 건도 전부 검사 대상.
    (거래당 API 호출을 최대 8회 내외로 제한 - 이전 버전은 최대 40회까지 순차 호출해서
    수출입은행 서버의 사실상 레이트리밋에 걸렸었음)"""
    results = []
    for _, row in settlements.iterrows():
        cur_unit = CUR_UNIT_MAP.get(row["통화"], row["통화"])
        company_rate = row["회사적용환율(내재)"]
        settle_date = row["결제일"]

        matched_date = None
        for candidate_date in _candidate_wrong_dates(settle_date):
            rate_table = fetch_rates_for_date(candidate_date.strftime("%Y%m%d"))
            if cur_unit not in rate_table:
                continue
            candidate_rate = normalize_rate(rate_table[cur_unit], row["통화"])
            if candidate_rate == 0:
                continue
            if abs(company_rate - candidate_rate) / candidate_rate <= REF_DATE_MATCH_TOLERANCE:
                matched_date = candidate_date
                break

        official_settle_raw, _ = get_official_rate(settle_date, cur_unit)
        official_settle_rate = normalize_rate(official_settle_raw, row["통화"])
        already_matches_settle_date = (
            official_settle_rate != 0 and
            abs(company_rate - official_settle_rate) / official_settle_rate <= REF_DATE_MATCH_TOLERANCE
        )

        if already_matches_settle_date:
            verdict = "정상(결제일 환율 사용 확인)"
            matched_date_str = None
        elif matched_date is not None:
            verdict = f"기준일 오류 의심 - {matched_date.strftime('%Y-%m-%d')} 환율을 사용한 것으로 보임"
            matched_date_str = matched_date.strftime("%Y-%m-%d")
        else:
            verdict = "불일치(흔한 패턴과 매칭 안 됨 - 은행 우대환율 등 개별 사유 가능, 증빙 확인 필요)"
            matched_date_str = None

        results.append({
            "거래ID": row["거래ID"], "결제일": settle_date.strftime("%Y-%m-%d"),
            "회사적용환율(내재)": company_rate, "기준일판정": verdict, "추정사용일자": matched_date_str,
        })

    return pd.DataFrame(results)


# ------------------------------------------------------------------
# 4. 2단계: 증빙 OCR (Claude Vision)
# ------------------------------------------------------------------

def _find_evidence_file(txn_id: str) -> str | None:
    for ext in (".png", ".jpg", ".jpeg", ".pdf"):
        candidate = os.path.join(EVIDENCE_DIR, f"{txn_id}{ext}")
        if os.path.exists(candidate):
            return candidate
    return None


def extract_rate_from_evidence(image_path: str) -> dict:
    """증빙 이미지(은행 외환거래확인서 등)에서 실제 적용환율/금액을 Claude Vision으로 추출.
    반환: {"거래일자":..., "통화":..., "적용환율":..., "외화금액":..., "원화금액":...}
    파싱 실패 시 raw_text에 원본 응답을 담아 반환하니, 로그로 확인해서 프롬프트를 조정하세요."""
    import anthropic  # 로컬 실행 시 pip install anthropic 필요

    client = anthropic.Anthropic()  # ANTHROPIC_API_KEY 환경변수 사용

    with open(image_path, "rb") as f:
        image_b64 = base64.standard_b64encode(f.read()).decode("utf-8")

    ext = os.path.splitext(image_path)[1].lower()
    media_type = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg"}.get(ext.strip("."), "image/png")

    if ext == ".pdf":
        content_block = {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": image_b64}}
    else:
        content_block = {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": image_b64}}

    prompt = (
        "이 이미지는 은행의 외환거래확인서, SWIFT 통지서, 또는 외화예금 거래명세표입니다. "
        "여기서 실제 적용된 환율과 거래 정보를 추출하세요. "
        "다른 설명 없이 아래 형식의 JSON 객체 하나만 답하세요:\n"
        '{"거래일자": "YYYY-MM-DD", "통화": "USD", "적용환율": 1234.5, '
        '"외화금액": 10000, "원화금액": 12345000}\n'
        "값을 찾을 수 없는 항목은 null로 표기하세요."
    )

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=500,
        messages=[{"role": "user", "content": [content_block, {"type": "text", "text": prompt}]}],
    )

    raw_text = "".join(b.text for b in message.content if b.type == "text")
    cleaned = re.sub(r"```json|```", "", raw_text).strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        return {"error": "JSON 파싱 실패", "raw_text": raw_text}


def verify_with_evidence(flagged_df: pd.DataFrame) -> pd.DataFrame:
    """1단계에서 플래그된 건에 대해 증빙 파일이 있으면 OCR로 검증, 없으면 '증빙요청필요' 표시."""
    rows = []
    for _, row in flagged_df.iterrows():
        record = row.to_dict()
        evidence_path = _find_evidence_file(row["거래ID"])

        if evidence_path is None:
            record["증빙상태"] = "증빙 요청 필요"
            record["증빙확인환율"] = None
            record["최종판정"] = "미확인(증빙 미확보)"
            rows.append(record)
            continue

        extracted = extract_rate_from_evidence(evidence_path)
        if "error" in extracted:
            record["증빙상태"] = "OCR 인식 실패 - 수기 확인 필요"
            record["증빙확인환율"] = None
            record["최종판정"] = "미확인(OCR 실패)"
            rows.append(record)
            continue

        confirmed_rate = extracted.get("적용환율")
        record["증빙상태"] = "증빙 확인 완료"
        record["증빙확인환율"] = confirmed_rate

        if confirmed_rate is not None:
            currency = row["통화"]
            fc = row["외화금액"]
            # 재계산: 발생시점 원가와의 차이는 별도 원장 대조가 필요하므로,
            # 여기서는 "증빙 환율 vs 회사 계상 환율" 자체의 일치 여부를 우선 확인
            company_rate = row["회사적용환율(내재)"]
            rate_match = abs(confirmed_rate - company_rate) <= max(company_rate * 0.001, 0.01)
            record["최종판정"] = "적정(증빙과 일치)" if rate_match else "부적정(증빙과 불일치 - 재계산 필요)"
        else:
            record["최종판정"] = "미확인(증빙에서 환율 추출 실패)"

        rows.append(record)

    return pd.DataFrame(rows)


# ------------------------------------------------------------------
# 5. 외화환산손익 (기말평가) - 전수 자동 검증
# ------------------------------------------------------------------

def verify_yearend_translation(ye_df: pd.DataFrame, missing_df: pd.DataFrame, year_end_date: str) -> pd.DataFrame:
    rows = []
    ye_dt = pd.to_datetime(year_end_date)

    for _, row in ye_df.iterrows():
        cur_unit = CUR_UNIT_MAP.get(row["통화"], row["통화"])
        official_raw, actual_date = get_official_rate(ye_dt, cur_unit)
        official_rate = normalize_rate(official_raw, row["통화"])
        company_rate = normalize_rate(row["회사적용환율(기말)"], row["통화"]) if row["통화"] == "JPY" else row["회사적용환율(기말)"]

        rate_match = abs(company_rate - official_rate) <= official_rate * 0.001
        rows.append({
            **row.to_dict(),
            "공식결산환율": round(official_rate, 4),
            "일치여부": "적정" if rate_match else "부적정(환율 오류)",
        })

    for _, row in missing_df.iterrows():
        rows.append({
            "거래ID": row["전표번호(거래ID)"], "결산일": year_end_date, "통화": row["통화"],
            "외화금액": row["기말미결제외화잔액"], "회사적용환율(기말)": None,
            "공식결산환율": None, "일치여부": "부적정(기말평가 누락)",
        })

    return pd.DataFrame(rows)


# ------------------------------------------------------------------
# 6. 실행
# ------------------------------------------------------------------

if __name__ == "__main__":
    journal = load_journal("분개장.xlsx")
    schedule = pd.read_excel("명세서_외화자산부채명세서.xlsx")

    print("=== A. 외환차손익 1단계 스크리닝 ===")
    settlements = extract_settlement_transactions(journal)
    screened = screen_fx_settlements(settlements)
    print(screened[["거래ID", "통화", "회사적용환율(내재)", "공식매매기준율", "괴리율(%)",
                     "1차플래그", "회사계상액과의차이(KRW)"]])

    print("\n=== A-보조. 합계 중요성 검증 (개별 통과 건 포함 전체 합산) ===")
    agg = check_aggregate_materiality(screened, AMPT)
    print(agg)

    print("\n=== A-보조2. 결제일-환율기준일 불일치 탐지 (5% 통과 건 포함 전수 검사) ===")
    ref_check = detect_reference_date_mismatch(settlements)
    print(ref_check[["거래ID", "결제일", "기준일판정", "추정사용일자"]])

    print("\n=== A. 외환차손익 2단계 증빙검증 ===")
    if agg["중요성초과여부"]:
        print("↳ 합계 중요성 초과 - 개별 통과 건까지 포함해 전체를 정밀검증 대상으로 확장합니다.")
        to_verify = screened
    else:
        to_verify = screened[screened["1차플래그"] == "이상치(정밀검증필요)"]
    verified = verify_with_evidence(to_verify)
    print(verified[["거래ID", "증빙상태", "증빙확인환율", "최종판정"]])

    print("\n=== B. 외화환산손익 전수 검증 ===")
    ye_txns = extract_yearend_transactions(journal)
    missing_ye = get_unsettled_yearend_candidates(journal, schedule)
    ye_result = verify_yearend_translation(ye_txns, missing_ye, "2025-12-31")
    print(ye_result[["거래ID", "통화", "일치여부"]])
