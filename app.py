# -*- coding: utf-8 -*-
"""
숨고 SA 통합 리포트 성과 코멘트 생성기
=====================================
네이버/구글 x 요청/고수 4개 매체-유형 조합에 대해
RAW / RAW2 시트를 파싱하여 자동으로 성과 코멘트 초안을 생성하는 Streamlit 앱.

사용법:
    streamlit run app.py

입력:
    - 숨고 SA 통합 리포트 (.xlsb) 파일 1개
    - 기준일 (= '전일', 리포트 기준 가장 최신 일자)

데이터 소스:
    - RAW  시트: 매체 x 유형 단위 일자별 원본 데이터 (전체 탑라인 계산용)
    - RAW2 시트: 매체 x 유형 x 카테고리1 x 카테고리2 x 서비스명 단위
                 일자별 원본 데이터 (카테고리/서비스 상세 분석용, 숨김 시트)

    기준일(전일)과 기준일-7일(전주 동요일) 두 날짜의 데이터를 실제 '날짜' 컬럼으로
    직접 필터링하여 비교합니다. (기존 '일_매체_유형' 시트의 사전 계산된
    전일/전주 스냅샷 대신, RAW/RAW2의 원본 일자 데이터를 기준으로 매번 재계산)

출력:
    - 매체 x 유형 4종 각각에 대한 코멘트 초안 (텍스트, 다운로드 가능)

주의:
    - 키워드 레벨 데이터는 통합 리포트에 없으므로, 성과 변동폭이 큰 서비스에는
      '*키워드 성과 확인 필요' 문구를 자동으로 삽입합니다.
      (추후 네이버 검색광고 API 연동 시 이 부분을 실제 키워드 데이터로 대체 예정)
"""

import tempfile
from datetime import date, timedelta

import pandas as pd
import streamlit as st
from pyxlsb import open_workbook


# ----------------------------------------------------------------------------
# 0. 설정값
# ----------------------------------------------------------------------------

MEDIA_TYPE_COMBOS = [
    ("네이버", "요청"),
    ("네이버", "고수"),
    ("구글", "요청"),
    ("구글", "고수"),
]

WEEKDAY_KR = ["월", "화", "수", "목", "금", "토", "일"]

EXCEL_EPOCH = date(1899, 12, 30)  # 엑셀 날짜 시리얼 넘버 기준일 (1900 윤년 버그 보정 포함)

RAW_SHEET_NAME = "RAW"
RAW2_SHEET_NAME = "RAW2"

# RAW 시트 숫자 컬럼
RAW_NUMERIC_COLS = [
    "노출", "클릭", "광고비", "UA", "REQ",
    "UA_요청 (AB)", "UA_고수 (AB)", "REQ (AB)",
    "결제고수 (AB)", "견적서발송 (AB)", "RR", "CAC", "CASH", "REWARD",
]

# RAW2 시트 숫자 컬럼
RAW2_NUMERIC_COLS = [
    "CAP", "노출", "클릭", "광고비",
    "UA_요청 (AB)", "UA_고수 (AB)", "REQ (AB)", "UA", "REQ", "RR", "CASH", "REWARD",
]


# ----------------------------------------------------------------------------
# 1. 날짜 유틸
# ----------------------------------------------------------------------------

def excel_serial_to_date(serial) -> date | None:
    """엑셀 날짜 시리얼 넘버 -> python date. 값이 없으면 None."""
    if serial is None or pd.isna(serial):
        return None
    return EXCEL_EPOCH + timedelta(days=int(serial))


def date_to_excel_serial(d: date) -> int:
    return (d - EXCEL_EPOCH).days


def weekday_label(d: date) -> str:
    return WEEKDAY_KR[d.weekday()]


# ----------------------------------------------------------------------------
# 2. 시트 로딩 (RAW / RAW2)
# ----------------------------------------------------------------------------

def _read_sheet_rows(file_bytes: bytes, sheet_name: str):
    with tempfile.NamedTemporaryFile(suffix=".xlsb", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    with open_workbook(tmp_path) as wb:
        if sheet_name not in wb.sheets:
            return None
        with wb.get_sheet(sheet_name) as sheet:
            return list(sheet.rows())


@st.cache_data(show_spinner=False)
def load_raw_sheet(file_bytes: bytes) -> pd.DataFrame | None:
    """RAW 시트: 매체 x 유형 단위 일자별 데이터. 탑라인(전체 합계) 계산에 사용."""
    rows = _read_sheet_rows(file_bytes, RAW_SHEET_NAME)
    if rows is None or len(rows) < 2:
        return None

    header = [c.v for c in rows[0]]
    data = [[c.v for c in r] for r in rows[1:]]
    df = pd.DataFrame(data, columns=header)

    df = df.dropna(subset=["매체", "유형", "날짜"]).reset_index(drop=True)
    df["날짜_date"] = df["날짜"].apply(excel_serial_to_date)

    for c in RAW_NUMERIC_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    return df


@st.cache_data(show_spinner=False)
def load_raw2_sheet(file_bytes: bytes) -> pd.DataFrame | None:
    """RAW2 시트: 매체 x 유형 x 카테고리1 x 카테고리2 x 서비스명 단위 일자별 데이터.
    카테고리/서비스 상세 분석에 사용."""
    rows = _read_sheet_rows(file_bytes, RAW2_SHEET_NAME)
    if rows is None or len(rows) < 2:
        return None

    header = [c.v for c in rows[0]]
    # 마지막 2개 컬럼처럼 헤더명이 None인 경우 임시 이름 부여
    header = [h if h not in (None, "") else f"unnamed_{i}" for i, h in enumerate(header)]
    data = [[c.v for c in r] for r in rows[1:]]
    df = pd.DataFrame(data, columns=header)

    df = df.rename(columns={
        "카테고리1": "cate1",
        "카테고리2": "cate2",
        "서비스명": "서비스",
    })

    df = df.dropna(subset=["매체", "유형", "날짜"]).reset_index(drop=True)
    df["날짜_date"] = df["날짜"].apply(excel_serial_to_date)

    for c in RAW2_NUMERIC_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    return df


def available_date_range(raw2_df: pd.DataFrame):
    dates = raw2_df["날짜_date"].dropna()
    if dates.empty:
        return None, None
    return dates.min(), dates.max()


# ----------------------------------------------------------------------------
# 3. 매체/유형/날짜 필터링 + 비교용 데이터프레임 생성
# ----------------------------------------------------------------------------

def pct_change(new: float, old: float):
    """old 대비 new 의 증감률. old가 0/NaN이면 None 반환."""
    if old in (0, None) or pd.isna(old) or pd.isna(new):
        return None
    return (new - old) / old


def compute_topline_from_raw(raw_df: pd.DataFrame, media: str, type_: str,
                              base_date: date, prev_date: date) -> dict:
    """RAW 시트 기준으로 기준일 vs 전주 동요일 탑라인 합계 계산."""
    sub = raw_df[(raw_df["매체"] == media) & (raw_df["유형"] == type_)]

    today = sub[sub["날짜_date"] == base_date]
    prev = sub[sub["날짜_date"] == prev_date]

    ad_today = today["광고비"].sum()
    ad_prev = prev["광고비"].sum()
    req_today = today["REQ (AB)"].sum()
    req_prev = prev["REQ (AB)"].sum()

    cpr_today = ad_today / req_today if req_today else None
    cpr_prev = ad_prev / req_prev if req_prev else None

    return {
        "ad_today": ad_today, "ad_prev": ad_prev,
        "req_today": req_today, "req_prev": req_prev,
        "cpr_today": cpr_today, "cpr_prev": cpr_prev,
        "ad_gap_pct": pct_change(ad_today, ad_prev),
        "req_gap_pct": pct_change(req_today, req_prev),
        "cpr_gap_pct": pct_change(cpr_today, cpr_prev),
        "today_has_data": not today.empty,
        "prev_has_data": not prev.empty,
    }


def build_compat_df(raw2_df: pd.DataFrame, media: str, type_: str,
                     base_date: date, prev_date: date) -> pd.DataFrame:
    """RAW2에서 기준일/전주 동요일 두 날짜를 cate1/cate2/서비스 단위로 집계 후 병합.
    이후 집계 로직(cate1_summary, cate2_summary, service_detail 등)이 기대하는
    '전일_광고비', '전주_광고비', '전일_AB_REQ', '전주_AB_REQ', '전일_CPR', '전주_CPR'
    컬럼 구조를 그대로 생성한다."""
    sub = raw2_df[(raw2_df["매체"] == media) & (raw2_df["유형"] == type_)]

    today = sub[sub["날짜_date"] == base_date]
    prev = sub[sub["날짜_date"] == prev_date]

    group_cols = ["cate1", "cate2", "서비스"]

    today_g = today.groupby(group_cols, dropna=False).agg(
        전일_광고비=("광고비", "sum"),
        전일_AB_REQ=("REQ (AB)", "sum"),
    ).reset_index()

    prev_g = prev.groupby(group_cols, dropna=False).agg(
        전주_광고비=("광고비", "sum"),
        전주_AB_REQ=("REQ (AB)", "sum"),
    ).reset_index()

    merged = pd.merge(today_g, prev_g, on=group_cols, how="outer")
    for c in ["전일_광고비", "전일_AB_REQ", "전주_광고비", "전주_AB_REQ"]:
        merged[c] = merged[c].fillna(0)

    merged["전일_CPR"] = merged.apply(
        lambda r: (r["전일_광고비"] / r["전일_AB_REQ"]) if r["전일_AB_REQ"] else None, axis=1)
    merged["전주_CPR"] = merged.apply(
        lambda r: (r["전주_광고비"] / r["전주_AB_REQ"]) if r["전주_AB_REQ"] else None, axis=1)

    return merged


# ----------------------------------------------------------------------------
# 4. 집계 함수 (카테고리1 / 카테고리2 / 서비스)
# ----------------------------------------------------------------------------

def cate1_summary(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("cate1").agg(
        광고비_전일=("전일_광고비", "sum"),
        광고비_전주=("전주_광고비", "sum"),
        REQ_전일=("전일_AB_REQ", "sum"),
        REQ_전주=("전주_AB_REQ", "sum"),
    ).reset_index()
    g["광고비_gap_pct"] = g.apply(lambda r: pct_change(r["광고비_전일"], r["광고비_전주"]), axis=1)
    g["REQ_gap_pct"] = g.apply(lambda r: pct_change(r["REQ_전일"], r["REQ_전주"]), axis=1)
    return g


def cate2_summary(df: pd.DataFrame) -> pd.DataFrame:
    # cate2 -> 대표 cate1 매핑 (같은 cate2가 항상 같은 cate1에 속한다고 가정)
    cate1_map = df.groupby("cate2")["cate1"].first()

    g = df.groupby("cate2").agg(
        광고비_전일=("전일_광고비", "sum"),
        광고비_전주=("전주_광고비", "sum"),
        REQ_전일=("전일_AB_REQ", "sum"),
        REQ_전주=("전주_AB_REQ", "sum"),
    ).reset_index()

    g["cate1"] = g["cate2"].map(cate1_map)
    g["CPR_전일"] = g.apply(lambda r: (r["광고비_전일"] / r["REQ_전일"]) if r["REQ_전일"] else None, axis=1)
    g["CPR_전주"] = g.apply(lambda r: (r["광고비_전주"] / r["REQ_전주"]) if r["REQ_전주"] else None, axis=1)

    g["REQ_gap"] = g["REQ_전일"] - g["REQ_전주"]
    g["REQ_gap_pct"] = g.apply(lambda r: pct_change(r["REQ_전일"], r["REQ_전주"]), axis=1)
    g["광고비_gap_pct"] = g.apply(lambda r: pct_change(r["광고비_전일"], r["광고비_전주"]), axis=1)
    g["CPR_gap_pct"] = g.apply(lambda r: pct_change(r["CPR_전일"], r["CPR_전주"]), axis=1)

    return g


def service_detail(df: pd.DataFrame, cate2: str) -> pd.DataFrame:
    sub = df[df["cate2"] == cate2].copy()
    sub["REQ_gap"] = sub["전일_AB_REQ"] - sub["전주_AB_REQ"]
    sub["REQ_gap_pct"] = sub.apply(lambda r: pct_change(r["전일_AB_REQ"], r["전주_AB_REQ"]), axis=1)
    sub["광고비_gap_pct"] = sub.apply(lambda r: pct_change(r["전일_광고비"], r["전주_광고비"]), axis=1)
    sub["CPR_gap_pct"] = sub.apply(lambda r: pct_change(r["전일_CPR"], r["전주_CPR"]), axis=1)
    sub = sub.reindex(sub["REQ_gap"].abs().sort_values(ascending=False).index)
    return sub


def select_top_cate2(cate2_df: pd.DataFrame, top_n: int, min_req_volume: int) -> pd.DataFrame:
    """전주 혹은 전일 REQ가 min_req_volume 이상인 cate2 중, REQ 절대 변동폭이 큰 순으로 top_n개 선정."""
    filtered = cate2_df[
        (cate2_df["REQ_전주"].fillna(0) >= min_req_volume)
        | (cate2_df["REQ_전일"].fillna(0) >= min_req_volume)
    ].copy()
    filtered = filtered.reindex(filtered["REQ_gap"].abs().sort_values(ascending=False).index)
    return filtered.head(top_n)


def needs_keyword_check(req_gap_pct, req_today, req_prev, threshold: float, min_req: int) -> bool:
    if req_gap_pct is None:
        return False
    volume_ok = (req_today or 0) >= min_req or (req_prev or 0) >= min_req
    return abs(req_gap_pct) >= threshold and volume_ok


# ----------------------------------------------------------------------------
# 5. 포맷 헬퍼
# ----------------------------------------------------------------------------

def fmt_money_man(won) -> str:
    if won is None or pd.isna(won):
        return "N/A"
    man = won / 10000
    return f"{man:,.0f}만원"

def fmt_money_won(won) -> str:
    if won is None or pd.isna(won):
        return "N/A"
    return f"{won:,.0f}원"

def fmt_cpr_range(cpr) -> str:
    if cpr is None or pd.isna(cpr):
        return "N/A"
    hundred = int(cpr // 100 * 100)
    return f"{hundred:,}원대"

def fmt_pct_directional(x, up_word="증가", down_word="감소", digits=0) -> str:
    if x is None or pd.isna(x):
        return "변동 없음(비교 불가)"
    word = up_word if x > 0 else down_word
    return f"{abs(x)*100:.{digits}f}% {word}"

def fmt_int(x) -> str:
    if x is None or pd.isna(x):
        return "0"
    return f"{x:,.0f}"


# ----------------------------------------------------------------------------
# 6. 코멘트 생성
# ----------------------------------------------------------------------------

def build_topline_line(base_date: date, prev_date: date, topline: dict) -> str:
    date_str = f"{base_date.month}/{base_date.day}({weekday_label(base_date)})"

    warnings = []
    if not topline["today_has_data"]:
        warnings.append(f"⚠ {base_date.isoformat()}(기준일) 데이터가 RAW 시트에 없습니다.")
    if not topline["prev_has_data"]:
        warnings.append(f"⚠ {prev_date.isoformat()}(전주 동요일) 데이터가 RAW 시트에 없습니다.")

    ad_str = fmt_money_man(topline["ad_today"])
    cpr_str = fmt_cpr_range(topline["cpr_today"])
    cpr_change = fmt_pct_directional(topline["cpr_gap_pct"], up_word="상승", down_word="하락")
    req_str = fmt_int(topline["req_today"])
    ad_change = fmt_pct_directional(topline["ad_gap_pct"])
    req_change = fmt_pct_directional(topline["req_gap_pct"])

    line1 = f"{date_str} 광고비 {ad_str} 소진. REQ {req_str}건, CPR {cpr_str} (전주 동요일 대비 CPR {cpr_change})"
    line2 = f"- 전주 동요일({prev_date.month}/{prev_date.day}) 대비 광고비 {ad_change}, REQ {req_change}"

    lines = [line1, line2]
    lines.extend(warnings)
    return "\n".join(lines)


def build_cate2_block(row, cate1_row, service_df, kw_threshold: float, kw_min_req: int, top_services: int) -> str:
    cate2 = row["cate2"]
    lines = [f"카테고리2 : {cate2}"]

    ad_change = fmt_pct_directional(row["광고비_gap_pct"])
    req_change = fmt_pct_directional(row["REQ_gap_pct"], up_word="상승", down_word="하락")
    cpr_change = fmt_pct_directional(row["CPR_gap_pct"], up_word="상승", down_word="하락")
    cpr_from_to = ""
    if row["CPR_전일"] is not None and row["CPR_전주"] is not None and not pd.isna(row["CPR_전일"]) and not pd.isna(row["CPR_전주"]):
        cpr_from_to = f" ({fmt_money_won(row['CPR_전주'])} → {fmt_money_won(row['CPR_전일'])})"

    lines.append(f"전주 동요일 대비 광고비 {ad_change}, REQ {req_change}하며 CPR {cpr_change}{cpr_from_to}")

    # cate1 대비 역행/동행 여부 코멘트
    if cate1_row is not None and cate1_row["REQ_gap_pct"] is not None and row["REQ_gap_pct"] is not None:
        cate1_dir = "증가" if cate1_row["REQ_gap_pct"] > 0 else "감소"
        cate2_dir = "증가" if row["REQ_gap_pct"] > 0 else "감소"
        if cate1_dir != cate2_dir:
            lines.append(
                f"- 카테고리1({cate1_row['cate1']}) 전체는 REQ {fmt_pct_directional(cate1_row['REQ_gap_pct'], up_word='증가', down_word='감소')} "
                f"기조였음에도, 이 카테고리는 역행하여 {cate2_dir}"
            )

    # 서비스 레벨 top movers
    top_svc = service_df.head(top_services)
    for _, srow in top_svc.iterrows():
        svc_name = srow["서비스"]
        svc_ad_change = fmt_pct_directional(srow["광고비_gap_pct"])
        svc_req_change = fmt_pct_directional(srow["REQ_gap_pct"], up_word="상승", down_word="하락")

        if srow["전주_AB_REQ"] == 0 and srow["전일_AB_REQ"] > 0:
            req_part = f"전주 REQ 미발생 → {fmt_int(srow['전일_AB_REQ'])}건 신규 발생"
        else:
            req_part = f"REQ {svc_req_change}"

        lines.append(f"ㄴ [{svc_name}] 광고비 {svc_ad_change}, {req_part}")

        if needs_keyword_check(srow["REQ_gap_pct"], srow["전일_AB_REQ"], srow["전주_AB_REQ"], kw_threshold, kw_min_req):
            lines.append("  *키워드 성과 확인 필요")

    return "\n".join(lines)


def generate_comment(media: str, type_: str, base_date: date, prev_date: date,
                      raw_df: pd.DataFrame, raw2_df: pd.DataFrame,
                      top_n_cate2: int, top_n_service: int,
                      min_req_volume: int, kw_threshold: float, kw_min_req: int) -> str:
    topline = compute_topline_from_raw(raw_df, media, type_, base_date, prev_date)
    compat_df = build_compat_df(raw2_df, media, type_, base_date, prev_date)

    header = f"[{media} / {type_}]\n"
    body = [header + build_topline_line(base_date, prev_date, topline)]

    if compat_df.empty:
        body.append("\n(RAW2 시트에서 해당 매체/유형/날짜 조합의 카테고리 데이터를 찾을 수 없습니다.)")
        return "\n".join(body)

    c1 = cate1_summary(compat_df)
    c2 = cate2_summary(compat_df)

    top_cate2 = select_top_cate2(c2, top_n_cate2, min_req_volume)

    if top_cate2.empty:
        body.append("\n(변동폭이 유의미한 카테고리2가 없거나, 데이터 볼륨이 낮아 추가 분석 대상이 없습니다.)")
    else:
        for _, row in top_cate2.iterrows():
            cate1_row = c1[c1["cate1"] == row["cate1"]]
            cate1_row = cate1_row.iloc[0] if not cate1_row.empty else None
            svc_df = service_detail(compat_df, row["cate2"])
            block = build_cate2_block(row, cate1_row, svc_df, kw_threshold, kw_min_req, top_n_service)
            body.append("\n" + block)

    body.append(
        "\n※ 키워드 레벨 성과는 통합 리포트에 포함되어 있지 않습니다. "
        "'*키워드 성과 확인 필요' 표시된 항목은 네이버/구글 검색광고 관리자센터에서 별도 확인이 필요합니다. "
        "(추후 네이버 검색광고 API 연동 예정)"
    )

    return "\n".join(body)


# ----------------------------------------------------------------------------
# 7. Streamlit UI
# ----------------------------------------------------------------------------

def main():
    st.set_page_config(page_title="숨고 SA 성과 코멘트 생성기", layout="wide")
    st.title("📊 숨고 SA 통합 리포트 → 성과 코멘트 생성기")
    st.caption(
        "RAW(매체x유형 일자별) / RAW2(카테고리·서비스 일자별) 시트를 기준으로, "
        "기준일(전일)과 기준일-7일(전주 동요일)을 비교하여 성과 코멘트 초안을 자동 생성합니다."
    )

    with st.sidebar:
        st.header("⚙️ 설정")
        uploaded_file = st.file_uploader("숨고 SA 통합 리포트 업로드 (.xlsb)", type=["xlsb"])

        default_base_date = date.today() - timedelta(days=1)
        base_date = st.date_input("기준일 (= 리포트 기준 가장 최신 일자, '전일')", value=default_base_date)
        prev_date = base_date - timedelta(days=7)
        st.caption(f"→ 전주 동요일: **{prev_date.isoformat()} ({weekday_label(prev_date)})** 로 자동 계산됩니다.")

        st.divider()
        st.subheader("코멘트 상세 옵션")
        top_n_cate2 = st.slider("카테고리2 노출 개수", min_value=1, max_value=8, value=3)
        top_n_service = st.slider("카테고리2당 서비스 노출 개수", min_value=1, max_value=5, value=2)
        min_req_volume = st.number_input("최소 REQ 볼륨 (이 이하 카테고리는 분석에서 제외)", min_value=0, value=10)

        st.divider()
        st.subheader("키워드 확인 플래그 기준")
        kw_threshold = st.slider("REQ 변동률 임계값 (%)", min_value=10, max_value=100, value=30) / 100
        kw_min_req = st.number_input("최소 REQ 건수 (이 이상일 때만 플래그)", min_value=1, value=10)

        generate_btn = st.button("🚀 코멘트 생성", type="primary", use_container_width=True)

    if not uploaded_file:
        st.info("왼쪽에서 통합 리포트(.xlsb) 파일을 업로드해주세요.")
        return

    file_bytes = uploaded_file.getvalue()

    with st.spinner("RAW / RAW2 시트 로딩 중..."):
        raw_df = load_raw_sheet(file_bytes)
        raw2_df = load_raw2_sheet(file_bytes)

    if raw_df is None:
        st.error(f"'{RAW_SHEET_NAME}' 시트를 찾을 수 없습니다. 파일 구조를 확인해주세요.")
        return
    if raw2_df is None:
        st.error(f"'{RAW2_SHEET_NAME}' 시트를 찾을 수 없습니다. 숨김 시트가 포함된 원본 파일을 업로드해주세요.")
        return

    min_d, max_d = available_date_range(raw2_df)
    if min_d and max_d:
        st.caption(f"RAW2 시트 내 데이터 범위: {min_d.isoformat()} ~ {max_d.isoformat()}")
        if base_date > max_d or base_date < min_d:
            st.warning(
                f"선택한 기준일({base_date.isoformat()})이 RAW2 데이터 범위를 벗어납니다. "
                "카테고리/서비스 상세 분석이 비어있을 수 있습니다."
            )
        if prev_date < min_d:
            st.warning(
                f"전주 동요일({prev_date.isoformat()})이 RAW2 데이터 범위 이전입니다. "
                "월초 리포트의 경우 RAW2에 전월 데이터가 없어 카테고리/서비스 비교가 제한될 수 있습니다."
            )

    if not generate_btn:
        st.info("옵션을 확인한 뒤 '코멘트 생성' 버튼을 눌러주세요.")
        return

    tabs = st.tabs([f"{media} · {type_}" for (media, type_) in MEDIA_TYPE_COMBOS])
    all_comments = []

    for tab, (media, type_) in zip(tabs, MEDIA_TYPE_COMBOS):
        with tab:
            comment = generate_comment(
                media, type_, base_date, prev_date, raw_df, raw2_df,
                top_n_cate2=top_n_cate2,
                top_n_service=top_n_service,
                min_req_volume=min_req_volume,
                kw_threshold=kw_threshold,
                kw_min_req=kw_min_req,
            )
            all_comments.append(comment)

            st.text_area("코멘트 (수정 가능)", value=comment, height=500, key=f"comment_{media}_{type_}")
            st.download_button(
                label="📥 이 코멘트 다운로드 (.txt)",
                data=comment.encode("utf-8"),
                file_name=f"{media}_{type_}_코멘트_{base_date.isoformat()}.txt",
                mime="text/plain",
                key=f"download_{media}_{type_}",
            )

            with st.expander("카테고리/서비스 원본 비교 데이터 미리보기 (RAW2 기준)"):
                preview = build_compat_df(raw2_df, media, type_, base_date, prev_date)
                st.dataframe(preview.sort_values("전일_광고비", ascending=False).head(30))

    if all_comments:
        st.divider()
        combined = "\n\n" + ("=" * 60) + "\n\n"
        combined_text = combined.join(all_comments)
        st.download_button(
            label="📥 전체 4종 코멘트 통합 다운로드 (.txt)",
            data=combined_text.encode("utf-8"),
            file_name=f"숨고_SA_전체코멘트_{base_date.isoformat()}.txt",
            mime="text/plain",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
