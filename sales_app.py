import os
import re
import io
from datetime import datetime

import pandas as pd
import streamlit as st


st.set_page_config(page_title="윤우 영업 통합 시스템", layout="wide")

# =========================================================
# 기본 설정
# =========================================================
DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

USERS = {
    "admin": {"pw": "1234", "role": "admin", "name": "관리자"},
    "user1": {"pw": "1234", "role": "user", "name": "영업1"},
    "user2": {"pw": "1234", "role": "user", "name": "영업2"},
}

# =========================================================
# 사업별 구글시트 링크
# =========================================================
BUSINESS_CONFIG = {
    "아이센서": {
        "sheets": {
            "영업현황": "https://docs.google.com/spreadsheets/d/1CWTvHC1r6i5wjcZoFJa5kAm-6T0SKao1EdjI8jwVQG8/edit?gid=167508641#gid=167508641",
            "가능단지": "https://docs.google.com/spreadsheets/d/1CWTvHC1r6i5wjcZoFJa5kAm-6T0SKao1EdjI8jwVQG8/edit?gid=1108943027#gid=1108943027",
            "입찰공고": "https://docs.google.com/spreadsheets/d/1CWTvHC1r6i5wjcZoFJa5kAm-6T0SKao1EdjI8jwVQG8/edit?gid=243967548#gid=243967548",
            "계약단지": "https://docs.google.com/spreadsheets/d/1CWTvHC1r6i5wjcZoFJa5kAm-6T0SKao1EdjI8jwVQG8/edit?gid=2071693391#gid=2071693391",
        },
        "menus": [
            "대시보드",
            "데이터 가져오기",
            "영업현황",
            "가능단지",
            "입찰공고단지",
            "계약단지",
            "라우터 관리",
            "오늘 할 일",
            "일정 관리",
            "영업 알림",
        ],
    },
    "전기차 충전기": {
        "sheets": {
            "계약접수현황": "https://docs.google.com/spreadsheets/d/1Efuld8DUSHs8jR42R5XKIT3y-3P_pf--aP9ED_SYWhc/edit?gid=784400128#gid=784400128",
        },
        "menus": [
            "대시보드",
            "데이터 가져오기",
            "계약접수현황",
            "오늘 할 일",
            "일정 관리",
            "영업 알림",
        ],
    },
}

# =========================================================
# 세션 상태
# =========================================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""
if "role" not in st.session_state:
    st.session_state.role = ""
if "display_name" not in st.session_state:
    st.session_state.display_name = ""
if "business" not in st.session_state:
    st.session_state.business = "아이센서"


# =========================================================
# 공통 유틸
# =========================================================
def clean_text(value):
    if pd.isna(value):
        return ""
    if isinstance(value, str):
        return re.sub(r"[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F]", "", value)
    return value


def make_unique_columns(cols):
    seen = {}
    new_cols = []
    for c in cols:
        base = str(c).strip()
        if base == "":
            base = "빈컬럼"
        if base not in seen:
            seen[base] = 0
            new_cols.append(base)
        else:
            seen[base] += 1
            new_cols.append(f"{base}_{seen[base]}")
    return new_cols


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    cols = []
    for c in df.columns:
        c = str(c).replace("\n", " ").strip()
        if c.startswith("Unnamed"):
            c = ""
        cols.append(c)
    df.columns = make_unique_columns(cols)
    return df


def remove_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    keep_cols = []
    for col in df.columns:
        if str(col).strip() in ["", "빈컬럼"]:
            continue
        col_data = df[col]
        if isinstance(col_data, pd.DataFrame):
            col_data = col_data.iloc[:, 0]
        if not col_data.isna().all():
            keep_cols.append(col)
    return df[keep_cols].copy()


def normalize_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def format_date_value(x):
    if pd.isna(x) or x == "":
        return ""
    try:
        if isinstance(x, pd.Timestamp):
            return x.strftime("%Y-%m-%d")
        if isinstance(x, datetime):
            return x.strftime("%Y-%m-%d")
        s = str(x).strip()
        s = s.replace(".", "-").replace("/", "-")
        return s
    except Exception:
        return str(x)


def preprocess_df(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_columns(df)
    df = df.fillna("")

    for col in df.columns:
        col_data = df[col]
        if isinstance(col_data, pd.DataFrame):
            col_data = col_data.iloc[:, 0]

        if any(key in str(col) for key in ["날짜", "일자", "마감", "공고", "접수", "실사", "발행", "입금", "계약"]):
            df[col] = col_data.apply(format_date_value)
        else:
            df[col] = col_data.apply(lambda x: x if isinstance(x, (int, float)) else normalize_text(x))
    return df


def business_prefix():
    return "sensor" if st.session_state.business == "아이센서" else "ev"


def local_file_path(sheet_key: str) -> str:
    prefix = business_prefix()
    safe_name = re.sub(r"[^가-힣a-zA-Z0-9_]", "_", sheet_key)
    return os.path.join(DATA_DIR, f"{prefix}_{safe_name}.csv")


def save_df(sheet_key: str, df: pd.DataFrame):
    path = local_file_path(sheet_key)
    df.to_csv(path, index=False, encoding="utf-8-sig")


def load_local_df(sheet_key: str) -> pd.DataFrame:
    path = local_file_path(sheet_key)
    if os.path.exists(path):
        try:
            return pd.read_csv(path, encoding="utf-8-sig").fillna("")
        except Exception:
            return pd.read_csv(path).fillna("")
    return pd.DataFrame()


def get_current_sheet_urls():
    return BUSINESS_CONFIG[st.session_state.business]["sheets"]


# =========================================================
# 구글 시트 연동
# =========================================================
import gspread
from google.oauth2.service_account import Credentials

def get_gsheet_client():
    import gspread
    from google.oauth2.service_account import Credentials
    import streamlit as st

    creds_dict = st.secrets["gcp_service_account"]

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client

    
def append_to_gsheet(sheet_url, row_data, worksheet_index=0):
    try:
        client = get_gsheet_client()

        sheet_id = re.search(r"/d/([a-zA-Z0-9-_]+)", sheet_url).group(1)
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.get_worksheet(worksheet_index)
        worksheet.append_row(row_data)

        return True
    except Exception as e:
        st.error(f"구글 저장 오류: {e}")
        return False

def update_billing_status_in_gsheet(sheet_url, 기준월, 단지명, 담당자, 청구금액, worksheet_index=0):
    try:
        client = get_gsheet_client()
        sheet_id = re.search(r"/d/([a-zA-Z0-9-_]+)", sheet_url).group(1)
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.get_worksheet(worksheet_index)

        values = worksheet.get_all_values()
        if not values or len(values) < 2:
            st.session_state["google_update_msg"] = "구글 시트에 데이터가 없습니다."
            return False

        target_ym = str(기준월).strip()
        target_name = str(단지명).strip()
        target_manager = str(담당자).strip()

        for i, row in enumerate(values[1:], start=2):
            row_기준월 = str(row[0]).strip() if len(row) > 0 else ""
            row_단지명 = str(row[1]).strip() if len(row) > 1 else ""
            row_담당자 = str(row[2]).strip() if len(row) > 2 else ""

            # 청구금액은 비교하지 않음
            if (
                row_기준월 == target_ym and
                row_단지명 == target_name and
                row_담당자 == target_manager
            ):
                worksheet.update_cell(i, 5, "입금")   # E열
                worksheet.update_cell(i, 6, 0)       # F열 미수금 0
                st.session_state["google_update_msg"] = (
                    f"구글 시트 업데이트 완료: {target_name} / 행={i}"
                )
                return True

        st.session_state["google_update_msg"] = (
            f"일치 행을 찾지 못했습니다: 기준월={target_ym}, 단지명={target_name}, 담당자={target_manager}"
        )
        return False

    except Exception as e:
        st.session_state["google_update_msg"] = f"구글 수금관리 업데이트 오류: {type(e).__name__} / {e}"
        return False

def convert_google_sheet_url_to_csv(url: str) -> str:
    sheet_match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    gid_match = re.search(r"gid=(\d+)", url)

    if not sheet_match:
        raise ValueError("구글 스프레드시트 문서 ID를 찾을 수 없습니다.")

    sheet_id = sheet_match.group(1)
    gid = gid_match.group(1) if gid_match else "0"
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}"


def detect_header_row(raw: pd.DataFrame) -> int:
    header_keywords = [
        "관리코드", "관리번호", "아파트명", "아파트 명", "단지명", "현장명",
        "연락처", "지역", "주소", "상품", "판매형태", "운영계약", "결제방식",
        "영업담당", "담당자", "실사담당", "수량", "판매가격", "계약날짜",
        "행위신고", "시공요청서", "시공", "시공여부", "세금계산서 발행",
        "세금계산서발행", "입금일", "영업수수료", "운영사", "구분",
        "계약서 유무", "서류 풀 세팅 완료 표시", "추가요금", "설치업체", "설치유무"
    ]

    if raw.empty:
        return 0

    scan_rows = min(len(raw), 10)
    best_score = -1
    best_idx = 0

    for i in range(scan_rows):
        row_values = raw.iloc[i].astype(str).str.strip().tolist()
        score = sum(1 for v in row_values if v in header_keywords)
        if score > best_score:
            best_score = score
            best_idx = i

    return best_idx


def force_fix_quantity_column(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    df = df.copy()

    def is_currency_only_name(name: str) -> bool:
        return str(name).strip() in ["₩", "￦"]

    def is_effective_blank(series: pd.Series) -> bool:
        s = series.astype(str).str.strip()
        return (s == "").all()

    def rename_sensor_contract_columns(df_inner: pd.DataFrame) -> pd.DataFrame:
        cols = list(df_inner.columns)

        if "영업담당" not in cols:
            return df_inner

        manager_idx = cols.index("영업담당")
        right_cols = cols[manager_idx + 1:]

        usable_cols = []
        for c in right_cols:
            col_series = df_inner[c]
            if isinstance(col_series, pd.DataFrame):
                col_series = col_series.iloc[:, 0]

            col_name = str(c).strip()

            if is_currency_only_name(col_name):
                continue

            if is_effective_blank(col_series):
                continue

            usable_cols.append(c)

        expected = [
            "수량",
            "판매가격",
            "계약날짜",
            "행위신고",
            "시공요청서",
            "시공",
            "시공여부",
            "세금계산서 발행",
            "입금일",
            "영업수수료",
        ]

        rename_map = {}
        for old, new in zip(usable_cols, expected):
            rename_map[old] = new

        if rename_map:
            df_inner = df_inner.rename(columns=rename_map)
            df_inner.columns = make_unique_columns(df_inner.columns)

        drop_cols = [c for c in df_inner.columns if is_currency_only_name(c)]
        if drop_cols:
            df_inner = df_inner.drop(columns=drop_cols, errors="ignore")  
            
        # -----------------------------------------
        # 라우터 컬럼명 보정
        # -----------------------------------------
        cols_after = list(df_inner.columns)

        if "라우터청구대상" in cols_after:
            billing_idx = cols_after.index("라우터청구대상")

            if billing_idx + 1 < len(cols_after):
                next_col = cols_after[billing_idx + 1]
                if str(next_col).startswith("빈컬럼"):
                    df_inner = df_inner.rename(columns={next_col: "라우터월비용"})

            cols_after = list(df_inner.columns)
            if billing_idx + 2 < len(cols_after):
                next_col = cols_after[billing_idx + 2]
                if str(next_col).startswith("빈컬럼"):
                    df_inner = df_inner.rename(columns={next_col: "라우터청구시작월"})

            cols_after = list(df_inner.columns)
            if billing_idx + 3 < len(cols_after):
                next_col = cols_after[billing_idx + 3]
                if str(next_col).startswith("빈컬럼"):
                    df_inner = df_inner.rename(columns={next_col: "라우터청구종료월"})

        df_inner.columns = make_unique_columns(df_inner.columns)
        return df_inner

    def rename_ev_contract_columns(df_inner: pd.DataFrame) -> pd.DataFrame:
        cols = list(df_inner.columns)

        if "주소" not in cols:
            return df_inner

        addr_idx = cols.index("주소")
        rename_map = {}

        # 주소 다음 컬럼들 강제 보정
        if addr_idx + 1 < len(cols):
            rename_map[cols[addr_idx + 1]] = "설치대수"
        if addr_idx + 2 < len(cols):
            rename_map[cols[addr_idx + 2]] = "주차면"
        if addr_idx + 3 < len(cols):
            rename_map[cols[addr_idx + 3]] = "계약서 유무"
        if addr_idx + 4 < len(cols):
            rename_map[cols[addr_idx + 4]] = "서류 풀 세팅 완료 표시"
        if addr_idx + 5 < len(cols):
            rename_map[cols[addr_idx + 5]] = "추가요금"
        if addr_idx + 6 < len(cols):
            rename_map[cols[addr_idx + 6]] = "설치업체"
        if addr_idx + 7 < len(cols):
            rename_map[cols[addr_idx + 7]] = "설치유무"
        if addr_idx + 8 < len(cols):
            rename_map[cols[addr_idx + 8]] = "계약날짜"
        if addr_idx + 9 < len(cols):
            rename_map[cols[addr_idx + 9]] = "원본 등기 발송"
        if addr_idx + 10 < len(cols):
            rename_map[cols[addr_idx + 10]] = "계약기간"
        if addr_idx + 11 < len(cols):
            rename_map[cols[addr_idx + 11]] = "운영사 접수"
        if addr_idx + 12 < len(cols):
            rename_map[cols[addr_idx + 12]] = "기타"
        if addr_idx + 13 < len(cols):
            rename_map[cols[addr_idx + 13]] = "전력인입"
        if addr_idx + 14 < len(cols):
            rename_map[cols[addr_idx + 14]] = "특이사항"

        df_inner = df_inner.rename(columns=rename_map)
        df_inner.columns = make_unique_columns(df_inner.columns)

        # 숫자 컬럼 정리
        for num_col in ["설치대수", "주차면"]:
            if num_col in df_inner.columns:
                num = pd.to_numeric(
                    df_inner[num_col].astype(str).str.replace(",", "", regex=False).str.strip(),
                    errors="coerce"
                )
                if num.notna().sum() > 0:
                    df_inner[num_col] = num.apply(
                        lambda x: int(x) if pd.notna(x) and float(x).is_integer() else x
                    )

        return df_inner

    # 아이센서 계약단지 전용
    if st.session_state.business == "아이센서" and sheet_name == "계약단지":
        df = rename_sensor_contract_columns(df)

        if "수량" in df.columns:
            qty = pd.to_numeric(
                df["수량"].astype(str).str.replace(",", "", regex=False).str.strip(),
                errors="coerce"
            )
            if qty.notna().sum() > 0:
                df["수량"] = qty.apply(
                    lambda x: int(x) if pd.notna(x) and float(x).is_integer() else x
                )

        if "판매가격" in df.columns:
            price = pd.to_numeric(
                df["판매가격"].astype(str)
                .str.replace(",", "", regex=False)
                .str.replace("₩", "", regex=False)
                .str.replace("￦", "", regex=False)
                .str.strip(),
                errors="coerce"
            )
            if price.notna().sum() > 0:
                df["판매가격"] = price.apply(
                    lambda x: int(x) if pd.notna(x) and float(x).is_integer() else x
                )

        if "영업수수료" in df.columns:
            fee = pd.to_numeric(
                df["영업수수료"].astype(str)
                .str.replace(",", "", regex=False)
                .str.replace("₩", "", regex=False)
                .str.replace("￦", "", regex=False)
                .str.strip(),
                errors="coerce"
            )
            if fee.notna().sum() > 0:
                df["영업수수료"] = fee.apply(
                    lambda x: int(x) if pd.notna(x) and float(x).is_integer() else x
                )

        return df

    # 전기차 계약접수현황 전용
    if st.session_state.business == "전기차 충전기" and sheet_name == "계약접수현황":
        df = rename_ev_contract_columns(df)
        return df

    # 기타 공통 처리
    if "수량" in df.columns:
        qty = pd.to_numeric(
            df["수량"].astype(str).str.replace(",", "", regex=False).str.strip(),
            errors="coerce"
        )
        if qty.notna().sum() > 0:
            df["수량"] = qty.apply(
                lambda x: int(x) if pd.notna(x) and float(x).is_integer() else x
            )

    return df


@st.cache_data(ttl=60)
def load_google_sheet_data(business_name: str, sheet_name: str, url: str) -> pd.DataFrame:
    if not url or "여기에_" in url:
        return pd.DataFrame()

    csv_url = convert_google_sheet_url_to_csv(url)
    raw = pd.read_csv(csv_url, header=None, dtype=str).fillna("")

    if raw.empty:
        return pd.DataFrame()

    header_row_idx = detect_header_row(raw)

    headers = raw.iloc[header_row_idx].astype(str).str.strip().tolist()
    data = raw.iloc[header_row_idx + 1:].reset_index(drop=True).copy()
    data.columns = headers

    df = data.copy()
    df = preprocess_df(df)
    df = force_fix_quantity_column(df, sheet_name)
    return df


def load_df(sheet_key: str) -> pd.DataFrame:
    sheet_urls = get_current_sheet_urls()

    if sheet_key in sheet_urls:
        try:
            df = load_google_sheet_data(st.session_state.business, sheet_key, sheet_urls[sheet_key])
            if not df.empty:
                save_df(sheet_key, df)
                return df
        except Exception as e:
            st.warning(f"{sheet_key} 구글시트 불러오기 실패, 로컬 백업 사용: {e}")

    return load_local_df(sheet_key)


# =========================================================
# 공통 보조 파일 (사업 공통)
# =========================================================
COMMON_FILE_MAP = {
    "할일": os.path.join(DATA_DIR, "common_tasks.csv"),
    "일정": os.path.join(DATA_DIR, "common_schedule.csv"),
    "세금알림": os.path.join(DATA_DIR, "common_tax_alerts.csv"),
    "입대의알림": os.path.join(DATA_DIR, "common_meeting_alerts.csv"),
}


def save_common_df(key: str, df: pd.DataFrame):
    df.to_csv(COMMON_FILE_MAP[key], index=False, encoding="utf-8-sig")


def load_common_df(key: str) -> pd.DataFrame:
    path = COMMON_FILE_MAP[key]
    if os.path.exists(path):
        try:
            return pd.read_csv(path, encoding="utf-8-sig").fillna("")
        except Exception:
            return pd.read_csv(path).fillna("")
    return pd.DataFrame()


def ensure_common_file(key: str, columns: list[str]):
    if not os.path.exists(COMMON_FILE_MAP[key]):
        save_common_df(key, pd.DataFrame(columns=columns))


def init_files():
    ensure_common_file("할일", ["등록일시", "작성자", "사업", "할일"])
    ensure_common_file("일정", ["등록일시", "작성자", "사업", "일정명", "날짜"])
    ensure_common_file("세금알림", ["등록일시", "작성자", "사업", "단지명", "예정일", "상태", "비고"])
    ensure_common_file("입대의알림", ["등록일시", "작성자", "사업", "단지명", "입대의일자", "상태", "비고"])
    ensure_money_files()

def load_tasks_df():
    df = load_common_df("할일")
    for col in ["등록일시", "작성자", "사업", "할일"]:
        if col not in df.columns:
            df[col] = ""
    return df[["등록일시", "작성자", "사업", "할일"]].copy()


def save_tasks_df(df: pd.DataFrame):
    save_common_df("할일", df[["등록일시", "작성자", "사업", "할일"]].copy())


def load_schedule_df():
    df = load_common_df("일정")
    for col in ["등록일시", "작성자", "사업", "일정명", "날짜"]:
        if col not in df.columns:
            df[col] = ""
    return df[["등록일시", "작성자", "사업", "일정명", "날짜"]].copy()


def save_schedule_df(df: pd.DataFrame):
    save_common_df("일정", df[["등록일시", "작성자", "사업", "일정명", "날짜"]].copy())


def load_tax_alert_df():
    df = load_common_df("세금알림")
    for col in ["등록일시", "작성자", "사업", "단지명", "예정일", "상태", "비고"]:
        if col not in df.columns:
            df[col] = ""
    return df[["등록일시", "작성자", "사업", "단지명", "예정일", "상태", "비고"]].copy()


def save_tax_alert_df(df: pd.DataFrame):
    save_common_df("세금알림", df[["등록일시", "작성자", "사업", "단지명", "예정일", "상태", "비고"]].copy())


def load_meeting_alert_df():
    df = load_common_df("입대의알림")
    for col in ["등록일시", "작성자", "사업", "단지명", "입대의일자", "상태", "비고"]:
        if col not in df.columns:
            df[col] = ""
    return df[["등록일시", "작성자", "사업", "단지명", "입대의일자", "상태", "비고"]].copy()


def save_meeting_alert_df(df: pd.DataFrame):
    save_common_df("입대의알림", df[["등록일시", "작성자", "사업", "단지명", "입대의일자", "상태", "비고"]].copy())

# =========================================================
# 수금관리 / 미입금관리 / 담당자별현황
# =========================================================
MONEY_FILE_MAP = {
    "수금관리": os.path.join(DATA_DIR, "sensor_billing_master.csv"),
    "미입금관리": os.path.join(DATA_DIR, "sensor_unpaid.csv"),
    "담당자별현황": os.path.join(DATA_DIR, "sensor_manager_summary.csv"),
}

# 실사용 구글 수금관리 시트 URL
BILLING_SHEET_URL = "https://docs.google.com/spreadsheets/d/1QDf1No9Nz5CVu3BGxLr7omWNSLi8Pth7FKTBChh6ds0/edit?gid=970462186#gid=970462186"


def save_money_df(key: str, df: pd.DataFrame):
    path = MONEY_FILE_MAP[key]
    df.to_csv(path, index=False, encoding="utf-8-sig")


def load_money_df(key: str) -> pd.DataFrame:
    path = MONEY_FILE_MAP[key]
    if os.path.exists(path):
        try:
            return pd.read_csv(path, encoding="utf-8-sig").fillna("")
        except Exception:
            return pd.read_csv(path).fillna("")
    return pd.DataFrame()


def ensure_money_files():
    if not os.path.exists(MONEY_FILE_MAP["수금관리"]):
        save_money_df(
            "수금관리",
            pd.DataFrame(columns=[
                "기준월", "단지명", "담당자", "청구금액", "입금여부", "미수금"
            ])
        )

    if not os.path.exists(MONEY_FILE_MAP["미입금관리"]):
        save_money_df(
            "미입금관리",
            pd.DataFrame(columns=[
                "기준월", "단지명", "담당자", "청구금액", "입금여부", "미수금"
            ])
        )

    if not os.path.exists(MONEY_FILE_MAP["담당자별현황"]):
        save_money_df(
            "담당자별현황",
            pd.DataFrame(columns=[
                "담당자", "미수금합계"
            ])
        )


def normalize_payment_status(value):
    s = str(value).strip()
    if s in ["입금", "입금완료", "완료", "수금완료", "Y", "y", "yes", "Yes"]:
        return "입금"
    return ""


def safe_int(value, default=0):
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return default
    return int(num)


def load_billing_from_gsheet(sheet_url, worksheet_index=0) -> pd.DataFrame:
    try:
        client = get_gsheet_client()
        sheet_id = re.search(r"/d/([a-zA-Z0-9-_]+)", sheet_url).group(1)
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.get_worksheet(worksheet_index)

        values = worksheet.get_all_values()
        if not values or len(values) < 2:
            return pd.DataFrame(columns=["기준월", "단지명", "담당자", "청구금액", "입금여부", "미수금"])

        headers = [str(x).strip() for x in values[0]]
        rows = values[1:]
        df = pd.DataFrame(rows, columns=headers)

        for col in ["기준월", "단지명", "담당자", "청구금액", "입금여부", "미수금"]:
            if col not in df.columns:
                df[col] = ""

        df = df[["기준월", "단지명", "담당자", "청구금액", "입금여부", "미수금"]].copy()

        df["기준월"] = df["기준월"].astype(str).str.strip()
        df["단지명"] = df["단지명"].astype(str).str.strip()
        df["담당자"] = df["담당자"].astype(str).str.strip()
        df["입금여부"] = df["입금여부"].astype(str).apply(normalize_payment_status)

        df["청구금액"] = pd.to_numeric(df["청구금액"], errors="coerce").fillna(0)
        df["미수금"] = pd.to_numeric(df["미수금"], errors="coerce").fillna(0)

        df.loc[(df["청구금액"] <= 0) & (df["미수금"] > 0), "청구금액"] = df["미수금"]
        df.loc[(df["미수금"] <= 0) & (df["청구금액"] > 0) & (df["입금여부"] != "입금"), "미수금"] = df["청구금액"]
        df.loc[df["입금여부"] == "입금", "미수금"] = 0

        df["청구금액"] = df["청구금액"].astype(int)
        df["미수금"] = df["미수금"].astype(int)

        return df

    except Exception as e:
        st.session_state["google_update_msg"] = f"수금관리 구글시트 불러오기 오류: {type(e).__name__} / {e}"
        return pd.DataFrame(columns=["기준월", "단지명", "담당자", "청구금액", "입금여부", "미수금"])

def build_billing_rows_from_router_claim_df(claim_df: pd.DataFrame) -> pd.DataFrame:
    if claim_df is None or claim_df.empty:
        return pd.DataFrame(columns=[
            "기준월", "단지명", "담당자", "청구금액", "입금여부", "미수금"
        ])

    df = claim_df.copy()

    if "청구년월" not in df.columns:
        df["청구년월"] = ""

    if "단지명" not in df.columns:
        df["단지명"] = ""

    if "담당자" not in df.columns:
        df["담당자"] = ""

    if "라우터월비용" not in df.columns:
        df["라우터월비용"] = 0

    df["청구금액"] = pd.to_numeric(df["라우터월비용"], errors="coerce").fillna(0).astype(int)
    df["입금여부"] = ""
    df["미수금"] = df["청구금액"]

    out = df[["청구년월", "단지명", "담당자", "청구금액", "입금여부", "미수금"]].copy()
    out = out.rename(columns={"청구년월": "기준월"})

    return out


def add_monthly_billing_data(claim_df: pd.DataFrame):
    """
    계약단지에서 추출한 이번달 청구대상을
    구글 실사용시트(BILLING_SHEET_URL)에 직접 추가한다.
    이미 같은 기준월 + 단지명 + 담당자 행이 있으면 중복 생성하지 않는다.
    """
    try:
        if claim_df is None or claim_df.empty:
            return {
                "added_count": 0,
                "duplicate_count": 0,
                "total_count": 0
            }

        client = get_gsheet_client()
        sheet_id = re.search(r"/d/([a-zA-Z0-9-_]+)", BILLING_SHEET_URL).group(1)
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.get_worksheet(0)

        values = worksheet.get_all_values()
        if not values:
            # 헤더가 아예 없으면 생성
            worksheet.append_row(["기준월", "단지명", "담당자", "청구금액", "입금여부", "미수금"])
            values = [["기준월", "단지명", "담당자", "청구금액", "입금여부", "미수금"]]

        headers = [str(x).strip() for x in values[0]]
        rows = values[1:] if len(values) > 1 else []

        # 기존 행의 중복키 수집
        existing_keys = set()
        for row in rows:
            row_기준월 = str(row[0]).strip() if len(row) > 0 else ""
            row_단지명 = str(row[1]).strip() if len(row) > 1 else ""
            row_담당자 = str(row[2]).strip() if len(row) > 2 else ""
            key = f"{row_기준월}||{row_단지명}||{row_담당자}"
            existing_keys.add(key)

        add_rows = []
        duplicate_count = 0

        work_df = claim_df.copy()

        if "청구년월" not in work_df.columns:
            work_df["청구년월"] = ""

        if "단지명" not in work_df.columns:
            work_df["단지명"] = ""

        if "담당자" not in work_df.columns:
            work_df["담당자"] = ""

        if "라우터월비용" not in work_df.columns:
            work_df["라우터월비용"] = 0

        for _, r in work_df.iterrows():
            기준월 = str(r.get("청구년월", "")).strip()
            단지명 = str(r.get("단지명", "")).strip()
            담당자 = str(r.get("담당자", "")).strip()
            청구금액 = int(pd.to_numeric(r.get("라우터월비용", 0), errors="coerce") or 0)

            if not 기준월 or not 단지명:
                continue

            if 청구금액 <= 0:
                continue

            key = f"{기준월}||{단지명}||{담당자}"

            if key in existing_keys:
                duplicate_count += 1
                continue

            add_rows.append([
                기준월,
                단지명,
                담당자,
                청구금액,
                "",         # 입금여부
                청구금액    # 미수금
            ])
            existing_keys.add(key)

        if add_rows:
            worksheet.append_rows(add_rows, value_input_option="USER_ENTERED")

        total_count = len(rows) + len(add_rows)

        st.session_state["google_update_msg"] = (
            f"자동청구 생성 완료 / 신규 {len(add_rows)}건 / 중복 제외 {duplicate_count}건"
        )

        return {
            "added_count": len(add_rows),
            "duplicate_count": duplicate_count,
            "total_count": total_count
        }

    except Exception as e:
        st.session_state["google_update_msg"] = f"자동청구 생성 오류: {type(e).__name__} / {e}"
        return {
            "added_count": 0,
            "duplicate_count": 0,
            "total_count": 0
        }


def rebuild_billing_views():
    """
    기존 로컬 기반 함수는 더 이상 핵심 로직에서 사용하지 않음.
    남겨두되 동작은 최소화.
    """
    return


def load_billing_dashboard_data():
    """
    앱 수금관리 화면은 이제 구글 실사용시트를 원본으로 사용
    """
    billing_df = load_billing_from_gsheet(BILLING_SHEET_URL)

    if billing_df.empty:
        unpaid_df = pd.DataFrame(columns=["기준월", "단지명", "담당자", "청구금액", "입금여부", "미수금"])
        manager_df = pd.DataFrame(columns=["담당자", "미수금합계"])
        return billing_df, unpaid_df, manager_df

    unpaid_df = billing_df[billing_df["미수금"] > 0].copy()
    unpaid_df = unpaid_df[["기준월", "단지명", "담당자", "청구금액", "입금여부", "미수금"]].copy()

    if unpaid_df.empty:
        manager_df = pd.DataFrame(columns=["담당자", "미수금합계"])
    else:
        manager_df = (
            unpaid_df.groupby("담당자", dropna=False)["미수금"]
            .sum()
            .reset_index()
            .rename(columns={"미수금": "미수금합계"})
            .sort_values("미수금합계", ascending=False)
        )

    return billing_df, unpaid_df, manager_df


def mark_billing_paid(기준월, 단지명, 담당자, 청구금액):
    """
    입금 처리는 구글 실사용시트만 수정
    """
    amount_num = pd.to_numeric(청구금액, errors="coerce")
    if pd.isna(amount_num):
        st.session_state["google_update_msg"] = f"청구금액이 비어 있거나 숫자가 아닙니다: {청구금액}"
        return False

    return update_billing_status_in_gsheet(
        BILLING_SHEET_URL,
        str(기준월).strip(),
        str(단지명).strip(),
        str(담당자).strip(),
        int(amount_num)
    )

# =========================================================
# 로그인
# =========================================================
def login():
    st.title("🔐 윤우 영업 통합 시스템 로그인")
    user_id = st.text_input("아이디")
    password = st.text_input("비밀번호", type="password")

    if st.button("로그인"):
        if user_id in USERS and USERS[user_id]["pw"] == password:
            st.session_state.logged_in = True
            st.session_state.username = user_id
            st.session_state.role = USERS[user_id]["role"]
            st.session_state.display_name = USERS[user_id]["name"]
            st.rerun()
        else:
            st.error("아이디 또는 비밀번호가 맞지 않습니다.")


def logout():
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.session_state.display_name = ""
    st.rerun()


def is_admin():
    return st.session_state.role == "admin"


def current_user_name():
    return st.session_state.display_name or st.session_state.username


# =========================================================
# 컬럼 도우미
# =========================================================
def get_best_column(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def get_site_column(df):
    return get_best_column(df, ["아파트 명", "아파트명", "단지명", "현장명", "주소"])


def get_manager_column(df):
    return get_best_column(df, ["담당자", "영업담당", "실사담당", "작성자"])


def get_code_column(df):
    return get_best_column(df, ["관리코드", "관리 코드"])


# =========================================================
# 권한 필터
# =========================================================
def apply_role_filter(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or is_admin():
        return df

    manager_col = get_manager_column(df)
    if not manager_col:
        return df

    user_name = current_user_name().strip()
    return df[df[manager_col].astype(str).str.strip() == user_name].copy()


def apply_author_filter(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    business_name = st.session_state.business
    if "사업" in df.columns:
        df = df[df["사업"].astype(str).str.strip() == business_name].copy()

    if is_admin():
        return df

    if "작성자" not in df.columns:
        return df

    user_name = current_user_name().strip()
    return df[df["작성자"].astype(str).str.strip() == user_name].copy()


# =========================================================
# 화면 공통
# =========================================================
def filtered_df(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    result = df.copy()
    for col, value in filters.items():
        if col in result.columns and value and value != "전체":
            result = result[result[col].astype(str) == str(value)]
    return result


def keyword_filter(df: pd.DataFrame, keyword: str) -> pd.DataFrame:
    if not keyword or not keyword.strip():
        return df
    keyword = keyword.strip()
    mask = df.astype(str).apply(
        lambda row: row.str.contains(keyword, case=False, na=False).any(),
        axis=1,
    )
    return df[mask]


def style_status_value(val):
    s = str(val).strip()
    if s in ["계약", "완료", "시공완료", "유찰", "낙찰", "진행완료", "접수완료"]:
        return "background-color: #dff0d8; color: #1b5e20; font-weight: bold;"
    if s in ["진행중", "상담중", "검토중", "운영사 변경중", "필", "상", "접수중", "시공전"]:
        return "background-color: #fff4cc; color: #7a5c00; font-weight: bold;"
    if s in ["부결", "실패", "보류", "미진행", "하", "불가", "미접수"]:
        return "background-color: #f8d7da; color: #8a1f2d; font-weight: bold;"
    return ""


def convert_number_display(value, col_name=""):
    try:
        if value is None or pd.isna(value):
            return ""

        s = str(value).strip()
        if s == "":
            return ""

        # 날짜형 값은 그대로 유지
        if re.match(r"^\d{2,4}[.\-/]\d{1,2}([.\-/]\d{1,2})?$", s):
            return s

        numeric_cols = ["수량", "판매가격", "영업수수료", "설치대수", "주차면"]
        if col_name in numeric_cols:
            s2 = s.replace("₩", "").replace("￦", "").replace(",", "").replace(" ", "")
            num = pd.to_numeric(s2, errors="coerce")

            if pd.notna(num):
                # 정수면 문자열 정수로 반환
                if float(num).is_integer():
                    return f"{int(num):,}"

                # 소수점이 실제 있을 때만 표시
                return f"{float(num):,.2f}".rstrip("0").rstrip(".")

        return s
    except Exception:
        return str(value) if value is not None else ""
        return value


def prepare_display_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()
    df.columns = make_unique_columns(df.columns)

    # 통화 전용 컬럼 제거
    drop_cols = [c for c in df.columns if str(c).strip() in ["₩", "￦"]]
    if drop_cols:
        df = df.drop(columns=drop_cols, errors="ignore")

    # 완전히 비어 있는 빈컬럼 제거
    keep_cols = []
    for col in df.columns:
        col_name = str(col).strip()
        col_data = df[col]

        if isinstance(col_data, pd.DataFrame):
            col_data = col_data.iloc[:, 0]

        if col_name.startswith("빈컬럼"):
            if not col_data.astype(str).replace("", pd.NA).isna().all():
                keep_cols.append(col)
        else:
            keep_cols.append(col)

    df = df[keep_cols].copy()

    for col in df.columns:
        col_data = df[col]
        if isinstance(col_data, pd.DataFrame):
            col_data = col_data.iloc[:, 0]
        df[col] = col_data.apply(lambda x: convert_number_display(x, col))

    return df


def styled_dataframe(df: pd.DataFrame):
    df = prepare_display_df(df)

    target_cols = [
        c for c in df.columns
        if c in ["진행여부", "결과", "시공여부", "확인", "낙찰여부", "단지 반응도", "가능성", "구분", "상태", "설치유무"]
    ]

    if not target_cols:
        st.dataframe(df, use_container_width=True, height=500, hide_index=True)
        return

    styled = df.style
    for col in target_cols:
        styled = styled.map(style_status_value, subset=[col])

    st.dataframe(styled, use_container_width=True, height=500, hide_index=True)


def is_done_status(value):
    v = str(value).strip().lower()
    done_values = ["완료", "발행완료", "발행", "처리완료", "ok", "o", "y", "yes", "완", "끝"]
    return v in [x.lower() for x in done_values]


def get_d_day_label(target_date):
    today = pd.to_datetime(datetime.today().date())
    if pd.isna(target_date):
        return ""
    diff = (target_date - today).days
    if diff < 0:
        return f"D+{abs(diff)}"
    elif diff == 0:
        return "D-Day"
    else:
        return f"D-{diff}"


def make_alert_status(target_date, done_value):
    if is_done_status(done_value):
        return "완료"

    today = pd.to_datetime(datetime.today().date())
    if pd.isna(target_date):
        return "날짜없음"

    diff = (target_date - today).days
    if diff < 0:
        return "지남"
    elif diff == 0:
        return "오늘"
    elif diff <= 3:
        return "긴급"
    elif diff <= 7:
        return "임박"
    else:
        return "예정"


def style_alert_value(val):
    s = str(val).strip()
    if s == "완료":
        return "background-color: #dff0d8; color: #1b5e20; font-weight: bold;"
    if s in ["오늘", "긴급", "지남"]:
        return "background-color: #f8d7da; color: #8a1f2d; font-weight: bold;"
    if s in ["임박", "예정"]:
        return "background-color: #fff4cc; color: #7a5c00; font-weight: bold;"
    return ""


def show_alert_table(df: pd.DataFrame):
    if df.empty:
        st.info("해당 알림이 없습니다.")
        return

    styled = df.style
    if "상태표시" in df.columns:
        styled = styled.map(style_alert_value, subset=["상태표시"])
    st.dataframe(styled, use_container_width=True, hide_index=True)


# =========================================================
# 엑셀 다운로드
# =========================================================
def to_excel_bytes(data_dict):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        written_count = 0

        for sheet_name, df in data_dict.items():
            try:
                if df is None:
                    continue
                if not isinstance(df, pd.DataFrame):
                    df = pd.DataFrame(df)

                df2 = df.copy()
                if len(df2.columns) == 0 and df2.empty:
                    df2 = pd.DataFrame({"안내": ["데이터 없음"]})
                else:
                    df2.columns = [clean_text(str(col)) for col in df2.columns]
                    df2 = df2.apply(lambda col: col.map(clean_text))

                safe_sheet_name = str(sheet_name).strip() if sheet_name else f"Sheet{written_count + 1}"
                safe_sheet_name = clean_text(safe_sheet_name)[:31] or f"Sheet{written_count + 1}"
                df2.to_excel(writer, index=False, sheet_name=safe_sheet_name)
                written_count += 1
            except Exception as e:
                error_df = pd.DataFrame({"오류": [f"{sheet_name} 시트 저장 실패"], "상세": [str(e)]})
                error_df.to_excel(writer, index=False, sheet_name=f"오류{written_count + 1}"[:31])
                written_count += 1

        if written_count == 0:
            pd.DataFrame({"안내": ["저장할 데이터가 없습니다."]}).to_excel(writer, index=False, sheet_name="안내")

    output.seek(0)
    return output.getvalue()


def download_section(title, df, file_name):
    try:
        if df is None:
            st.warning(f"{title} 데이터 없음")
            return
        if not isinstance(df, pd.DataFrame):
            df = pd.DataFrame(df)
        if df.empty:
            st.warning(f"{title} 데이터 없음")
            return

        excel_data = to_excel_bytes({title: df})
        st.download_button(
            label=f"📥 {title} 다운로드",
            data=excel_data,
            file_name=f"{file_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_{file_name}_{title}",
        )
    except Exception as e:
        st.error(f"{title} 다운로드 오류: {e}")


# =========================================================
# 로고/헤더
# =========================================================
def render_header():
    col1, col2 = st.columns([1, 5])

    with col1:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=100)

    with col2:
        st.markdown(
            f"""
            <div style="padding-top:8px;">
                <h2 style="margin-bottom:0;">윤우 영업 통합 시스템</h2>
                <p style="margin-top:4px; color:gray;">현재 사업: <b>{st.session_state.business}</b></p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.divider()


# =========================================================
# 페이지들
# =========================================================
def page_dashboard():
    st.title("📊 대시보드")

    if st.session_state.business == "아이센서":
        sales_df = apply_role_filter(load_df("영업현황"))
        possible_df = apply_role_filter(load_df("가능단지"))
        bid_df = apply_role_filter(load_df("입찰공고"))
        contract_df = apply_role_filter(load_df("계약단지"))

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("영업현황", len(sales_df))
        c2.metric("가능단지", len(possible_df))
        c3.metric("입찰공고", len(bid_df))
        c4.metric("계약단지", len(contract_df))

        st.divider()

        # 라우터 KPI 추가
        router_base_df = build_router_base_df()
        if not router_base_df.empty:
            current_year = datetime.today().year
            current_month = datetime.today().month
            current_ym = f"{current_year:04d}-{current_month:02d}"

            router_df = router_base_df[router_base_df["라우터사용"] == "예"].copy()
            waiting_df = router_df[router_df["라우터명의이전상태"] == "이전대기"].copy()
            reject_df = router_df[router_df["라우터명의이전상태"] == "이전거부"].copy()
            charge_df = get_router_charge_target_df(router_base_df, current_year, current_month)

            r1, r2, r3, r4 = st.columns(4)
            r1.metric("라우터 사용 단지", len(router_df))
            r2.metric("이전대기", len(waiting_df))
            r3.metric("이전거부", len(reject_df))
            r4.metric(f"{current_ym} 청구대상", len(charge_df))

            if not charge_df.empty:
                st.subheader(f"📡 {current_ym} 라우터 청구대상 미리보기")
                preview_df = charge_df[[
                    "단지명_표시", "담당자_표시", "라우터명의이전상태",
                    "라우터청구대상", "라우터월비용", "라우터비고"
                ]].copy()
                preview_df = preview_df.rename(columns={
                    "단지명_표시": "단지명",
                    "담당자_표시": "담당자",
                })
                st.dataframe(preview_df.head(20), use_container_width=True, hide_index=True)

        st.divider()

        if is_admin() and not sales_df.empty:
            manager_col = get_manager_column(sales_df)
            if manager_col:
                st.subheader("영업현황 담당자별 건수")
                count_df = sales_df[manager_col].astype(str).replace("", "미지정").value_counts().reset_index()
                count_df.columns = ["담당자", "건수"]
                st.dataframe(count_df, use_container_width=True)

    else:
        contract_df = apply_role_filter(load_df("계약접수현황"))

        c1, c2, c3 = st.columns(3)
        c1.metric("충전기 계약접수현황", len(contract_df))

        manager_col = get_manager_column(contract_df)
        site_col = get_site_column(contract_df)

        if manager_col and not contract_df.empty:
            c2.metric("담당자 수", contract_df[manager_col].astype(str).replace("", pd.NA).dropna().nunique())
        else:
            c2.metric("담당자 수", 0)

        if site_col and not contract_df.empty:
            c3.metric("현장/단지 수", contract_df[site_col].astype(str).replace("", pd.NA).dropna().nunique())
        else:
            c3.metric("현장/단지 수", 0)

        st.divider()

        if not contract_df.empty:
            st.subheader("충전기 계약접수 최근 현황")
            st.dataframe(contract_df.head(20), use_container_width=True, hide_index=True)

    st.divider()

    tasks_df = apply_author_filter(load_tasks_df())
    schedule_df = apply_author_filter(load_schedule_df())
    tax_df = apply_author_filter(load_tax_alert_df())
    meeting_df = apply_author_filter(load_meeting_alert_df())

    c5, c6, c7 = st.columns(3)

    tax_temp = tax_df.copy()
    if not tax_temp.empty:
        tax_temp["예정일_dt"] = pd.to_datetime(tax_temp["예정일"], errors="coerce")
        tax_temp["상태표시"] = tax_temp.apply(lambda r: make_alert_status(r["예정일_dt"], r["상태"]), axis=1)
        tax_pending_count = len(tax_temp[tax_temp["상태표시"].isin(["지남", "오늘", "긴급", "임박"])])
    else:
        tax_pending_count = 0

    meeting_temp = meeting_df.copy()
    if not meeting_temp.empty:
        meeting_temp["입대의일자_dt"] = pd.to_datetime(meeting_temp["입대의일자"], errors="coerce")
        meeting_temp["상태표시"] = meeting_temp.apply(lambda r: make_alert_status(r["입대의일자_dt"], r["상태"]), axis=1)
        meeting_pending_count = len(meeting_temp[meeting_temp["상태표시"].isin(["지남", "오늘", "긴급", "임박"])])
    else:
        meeting_pending_count = 0

    schedule_temp = schedule_df.copy()
    if not schedule_temp.empty:
        schedule_temp["날짜_dt"] = pd.to_datetime(schedule_temp["날짜"], errors="coerce")
        schedule_temp["상태표시"] = schedule_temp["날짜_dt"].apply(lambda x: make_alert_status(x, ""))
        schedule_pending_count = len(schedule_temp[schedule_temp["상태표시"].isin(["지남", "오늘", "긴급", "임박"])])
    else:
        schedule_pending_count = 0

    c5.metric("세금계산서 알림", tax_pending_count)
    c6.metric("입대의 알림", meeting_pending_count)
    c7.metric("일정 알림", schedule_pending_count)

    st.divider()
    st.subheader("📝 오늘 할 일")
    if tasks_df.empty:
        st.info("등록된 할 일이 없습니다.")
    else:
        st.dataframe(tasks_df.tail(10).iloc[::-1], use_container_width=True, hide_index=True)


def page_import():
    st.title("📥 데이터 가져오기 / 구글시트 연결")
    st.info(f"현재 사업: {st.session_state.business}")

    sheet_urls = get_current_sheet_urls()

    with st.expander("현재 구글시트 링크 상태 보기", expanded=True):
        st.write(sheet_urls)

    st.divider()
    st.subheader("구글시트 연결 확인")
    for sheet_name, url in sheet_urls.items():
        try:
            df = load_google_sheet_data(st.session_state.business, sheet_name, url)
            if df.empty:
                st.warning(f"{sheet_name}: 링크 미입력 또는 데이터 없음")
            else:
                st.success(f"{sheet_name}: 연결 성공 ({len(df)}건)")
        except Exception as e:
            st.error(f"{sheet_name}: 연결 실패 - {e}")

    st.divider()
    st.subheader("현재 사업 전체 데이터 백업 다운로드")
    export_dict = {key: load_df(key) for key in sheet_urls.keys()}
    all_excel = to_excel_bytes(export_dict)
    st.download_button(
        label="전체 데이터 엑셀 백업 다운로드",
        data=all_excel,
        file_name=f"{st.session_state.business}_전체백업_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_all_backup_{st.session_state.business}",
    )


def generic_data_page(title, key, filters_config, search_key):
    st.title(title)
    df = load_df(key)
    df = apply_role_filter(df)

    if df.empty:
        st.warning(f"{title} 데이터가 없습니다.")
        return

    cols = st.columns(4)
    filters = {}

    for i, (label, candidates) in enumerate(filters_config):
        if i >= 4:
            break
        col_name = candidates if isinstance(candidates, str) else get_best_column(df, candidates)
        if col_name and col_name in df.columns:
            options = sorted([x for x in df[col_name].astype(str).unique() if x != ""])

            if not is_admin() and label in ["담당자", "영업담당"] and current_user_name() in options:
                cols[i].selectbox(label, [current_user_name()], index=0, key=f"{search_key}_{label}")
                filters[col_name] = current_user_name()
            else:
                value = cols[i].selectbox(label, ["전체"] + options, key=f"{search_key}_{label}")
                filters[col_name] = value

    keyword = st.text_input("검색", placeholder="키워드 검색", key=search_key)
    df2 = filtered_df(df, filters)
    df2 = keyword_filter(df2, keyword)

    st.write(f"조회 건수: {len(df2)}건")
    styled_dataframe(df2)
    download_section(f"{title}_필터결과", df2, title)


def page_tasks():
    st.title("📝 오늘 할 일")
    df = load_tasks_df()

    new_task = st.text_input("할 일 입력")
    if st.button("할 일 추가"):
        if new_task.strip():
            new_row = {
                "등록일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "작성자": current_user_name(),
                "사업": st.session_state.business,
                "할일": new_task.strip(),
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            save_tasks_df(df)

            st.success("추가 완료")
            st.rerun()

    st.write("---")
    view_df = apply_author_filter(df)

    if view_df.empty:
        st.info("등록된 할 일이 없습니다.")
    else:
        st.dataframe(view_df.iloc[::-1].reset_index(drop=True), use_container_width=True, hide_index=True)


def page_schedule():
    st.title("📅 일정 관리")
    df = load_schedule_df()

    title = st.text_input("일정 제목")
    selected_date = st.date_input("날짜")

    if st.button("일정 추가"):
        if title.strip():
            new_row = {
                "등록일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "작성자": current_user_name(),
                "사업": st.session_state.business,
                "일정명": title.strip(),
                "날짜": str(selected_date),
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            save_schedule_df(df)
            st.success("일정 추가 완료")
            st.rerun()

    st.write("---")
    view_df = apply_author_filter(df)

    if view_df.empty:
        st.info("등록된 일정이 없습니다.")
    else:
        temp_df = view_df.copy()
        temp_df["날짜정렬"] = pd.to_datetime(temp_df["날짜"], errors="coerce")
        temp_df = temp_df.sort_values(["날짜정렬", "등록일시"], ascending=[True, False]).drop(columns=["날짜정렬"])
        st.dataframe(temp_df, use_container_width=True, hide_index=True)


def page_alerts():
    st.title("🚨 영업 알림")

    tax_df = load_tax_alert_df()
    meeting_df = load_meeting_alert_df()
    schedule_df = load_schedule_df()

    view_tax_df = apply_author_filter(tax_df)
    view_meeting_df = apply_author_filter(meeting_df)
    view_schedule_df = apply_author_filter(schedule_df)

    st.subheader("1. 알림 요약")
    c1, c2, c3 = st.columns(3)

    tax_temp = view_tax_df.copy()
    if not tax_temp.empty:
        tax_temp["예정일_dt"] = pd.to_datetime(tax_temp["예정일"], errors="coerce")
        tax_temp["상태표시"] = tax_temp.apply(lambda r: make_alert_status(r["예정일_dt"], r["상태"]), axis=1)
        tax_pending = tax_temp[tax_temp["상태표시"].isin(["지남", "오늘", "긴급", "임박"])]
    else:
        tax_pending = pd.DataFrame()

    meeting_temp = view_meeting_df.copy()
    if not meeting_temp.empty:
        meeting_temp["입대의일자_dt"] = pd.to_datetime(meeting_temp["입대의일자"], errors="coerce")
        meeting_temp["상태표시"] = meeting_temp.apply(lambda r: make_alert_status(r["입대의일자_dt"], r["상태"]), axis=1)
        meeting_pending = meeting_temp[meeting_temp["상태표시"].isin(["지남", "오늘", "긴급", "임박"])]
    else:
        meeting_pending = pd.DataFrame()

    schedule_temp = view_schedule_df.copy()
    if not schedule_temp.empty:
        schedule_temp["날짜_dt"] = pd.to_datetime(schedule_temp["날짜"], errors="coerce")
        schedule_temp["상태표시"] = schedule_temp["날짜_dt"].apply(lambda x: make_alert_status(x, ""))
        schedule_pending = schedule_temp[schedule_temp["상태표시"].isin(["지남", "오늘", "긴급", "임박"])]
    else:
        schedule_pending = pd.DataFrame()

    c1.metric("세금계산서 알림", len(tax_pending))
    c2.metric("입대의 알림", len(meeting_pending))
    c3.metric("일정 알림", len(schedule_pending))

    st.divider()
    st.subheader("2. 세금계산서 발행 알림 등록")
    col1, col2, col3 = st.columns(3)
    tax_site = col1.text_input("단지명", key="tax_site")
    tax_date = col2.date_input("발행 예정일", key="tax_date")
    tax_note = col3.text_input("비고", key="tax_note")

    if st.button("세금계산서 알림 추가"):
        if tax_site.strip():
            new_row = {
                "등록일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "작성자": current_user_name(),
                "사업": st.session_state.business,
                "단지명": tax_site.strip(),
                "예정일": str(tax_date),
                "상태": "예정",
                "비고": tax_note.strip(),
            }
            tax_df = pd.concat([tax_df, pd.DataFrame([new_row])], ignore_index=True)
            save_tax_alert_df(tax_df)
            st.success("세금계산서 알림이 등록되었습니다.")
            st.rerun()

    st.subheader("세금계산서 알림 목록")
    if view_tax_df.empty:
        st.info("등록된 세금계산서 알림이 없습니다.")
    else:
        view_tax = view_tax_df.copy()
        view_tax["예정일_dt"] = pd.to_datetime(view_tax["예정일"], errors="coerce")
        view_tax["D-Day"] = view_tax["예정일_dt"].apply(get_d_day_label)
        view_tax["상태표시"] = view_tax.apply(lambda r: make_alert_status(r["예정일_dt"], r["상태"]), axis=1)
        view_tax = view_tax.sort_values(["예정일_dt", "등록일시"], ascending=[True, False])
        show_alert_table(view_tax[["단지명", "예정일", "D-Day", "상태", "상태표시", "비고", "작성자"]])

    st.divider()
    st.subheader("3. 입대의 알림 등록")
    m1, m2, m3 = st.columns(3)
    meeting_site = m1.text_input("단지명", key="meeting_site")
    meeting_date = m2.date_input("입대의일자", key="meeting_date")
    meeting_note = m3.text_input("비고", key="meeting_note")

    if st.button("입대의 알림 추가"):
        if meeting_site.strip():
            new_row = {
                "등록일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "작성자": current_user_name(),
                "사업": st.session_state.business,
                "단지명": meeting_site.strip(),
                "입대의일자": str(meeting_date),
                "상태": "예정",
                "비고": meeting_note.strip(),
            }
            meeting_df = pd.concat([meeting_df, pd.DataFrame([new_row])], ignore_index=True)
            save_meeting_alert_df(meeting_df)
            st.success("입대의 알림이 등록되었습니다.")
            st.rerun()

    st.subheader("입대의 알림 목록")
    if view_meeting_df.empty:
        st.info("등록된 입대의 알림이 없습니다.")
    else:
        view_meeting = view_meeting_df.copy()
        view_meeting["입대의일자_dt"] = pd.to_datetime(view_meeting["입대의일자"], errors="coerce")
        view_meeting["D-Day"] = view_meeting["입대의일자_dt"].apply(get_d_day_label)
        view_meeting["상태표시"] = view_meeting.apply(lambda r: make_alert_status(r["입대의일자_dt"], r["상태"]), axis=1)
        view_meeting = view_meeting.sort_values(["입대의일자_dt", "등록일시"], ascending=[True, False])
        show_alert_table(view_meeting[["단지명", "입대의일자", "D-Day", "상태", "상태표시", "비고", "작성자"]])


def page_admin_tools():
    st.title("🛠 관리자 도구")
    if not is_admin():
        st.error("이 메뉴는 관리자만 사용할 수 있습니다.")
        return

    if st.session_state.business == "아이센서":
        menu = st.radio("관리 작업 선택", ["영업현황 담당자 일괄 확인", "계약단지 담당자 일괄 확인"])

        if menu == "영업현황 담당자 일괄 확인":
            df = load_df("영업현황")
            if df.empty:
                st.info("데이터가 없습니다.")
                return
            site_col = get_site_column(df)
            manager_col = get_manager_column(df)
            if not site_col or not manager_col:
                st.warning("필수 컬럼을 찾지 못했습니다.")
                return
            st.dataframe(df[[site_col, manager_col]], use_container_width=True, height=500)

        elif menu == "계약단지 담당자 일괄 확인":
            df = load_df("계약단지")
            if df.empty:
                st.info("데이터가 없습니다.")
                return
            site_col = get_best_column(df, ["아파트명", "아파트 명", "단지명", "주소"])
            manager_col = get_manager_column(df)
            if not site_col or not manager_col:
                st.warning("필수 컬럼을 찾지 못했습니다.")
                return
            st.dataframe(df[[site_col, manager_col]], use_container_width=True, height=500)

    else:
        st.info("충전기 1차 버전에서는 계약접수현황만 연결되어 있습니다.")
        df = load_df("계약접수현황")
        if not df.empty:
            st.dataframe(df.head(50), use_container_width=True, height=500)

# =========================================================
# 라우터 관리
# =========================================================
ROUTER_COLUMNS = [
    "라우터사용",
    "라우터개통일",
    "라우터명의이전상태",
    "라우터명의이전일",
    "라우터청구대상",
    "라우터월비용",
    "라우터청구시작월",
    "라우터청구종료월",
    "라우터비고",
]

ROUTER_STATUS_OPTIONS = ["", "이전대기", "이전완료", "이전거부", "해지"]
ROUTER_BILLING_OPTIONS = ["", "청구", "미청구"]
ROUTER_USE_OPTIONS = ["", "예", "아니오"]


def clean_router_text(value):
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def normalize_year_month(value):
    s = clean_router_text(value)
    if s == "":
        return ""

    s = s.replace(".", "-").replace("/", "-").replace(" ", "")
    dt = pd.to_datetime(s, errors="coerce")

    if pd.notna(dt):
        return dt.strftime("%Y-%m")

    # YYYY-MM 형태 직접 허용
    if re.match(r"^\d{4}-\d{1,2}$", s):
        year, month = s.split("-")
        return f"{int(year):04d}-{int(month):02d}"

    return s


def normalize_date_string(value):
    s = clean_router_text(value)
    if s == "":
        return ""
    s = s.replace(".", "-").replace("/", "-").replace(" ", "")
    dt = pd.to_datetime(s, errors="coerce")
    if pd.notna(dt):
        return dt.strftime("%Y-%m-%d")
    return s


def router_safe_amount(value):
    s = clean_router_text(value)
    if s == "":
        return 0
    s = s.replace(",", "").replace("원", "").replace(" ", "")
    num = pd.to_numeric(s, errors="coerce")
    if pd.isna(num):
        return 0
    return int(round(float(num)))


def router_yes(value):
    return clean_router_text(value) == "예"


def build_router_base_df():
    if st.session_state.business != "아이센서":
        return pd.DataFrame()

    df = load_df("계약단지")
    df = apply_role_filter(df)

    if df.empty:
        return pd.DataFrame()

    df = df.copy()

    # 라우터 컬럼 없으면 생성
    for col in ROUTER_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    # 주요 컬럼 찾기
    site_col = get_best_column(df, ["아파트명", "아파트 명", "단지명", "현장명", "주소"])
    manager_col = get_manager_column(df)
    code_col = get_code_column(df)
    region_col = get_best_column(df, ["지역"])
    qty_col = get_best_column(df, ["수량"])

    if site_col is None:
        site_col = "단지명_표시"
        df[site_col] = ""

    if manager_col is None:
        manager_col = "담당자_표시"
        df[manager_col] = ""

    if code_col is None:
        code_col = "관리코드_표시"
        df[code_col] = ""

    if region_col is None:
        region_col = "지역_표시"
        df[region_col] = ""

    if qty_col is None:
        qty_col = "수량_표시"
        df[qty_col] = ""

    df["단지명_표시"] = df[site_col].astype(str).str.strip()
    df["담당자_표시"] = df[manager_col].astype(str).str.strip()
    df["관리코드_표시"] = df[code_col].astype(str).str.strip()
    df["지역_표시"] = df[region_col].astype(str).str.strip()
    df["수량_표시"] = df[qty_col].astype(str).str.strip()

    # 라우터 컬럼 정규화
    df["라우터사용"] = df["라우터사용"].apply(clean_router_text)
    df["라우터개통일"] = df["라우터개통일"].apply(normalize_date_string)
    df["라우터명의이전상태"] = df["라우터명의이전상태"].apply(clean_router_text)
    df["라우터명의이전일"] = df["라우터명의이전일"].apply(normalize_date_string)
    df["라우터청구대상"] = df["라우터청구대상"].apply(clean_router_text)
    df["라우터월비용"] = df["라우터월비용"].apply(router_safe_amount)
    df["라우터청구시작월"] = df["라우터청구시작월"].apply(normalize_year_month)
    df["라우터청구종료월"] = df["라우터청구종료월"].apply(normalize_year_month)
    df["라우터비고"] = df["라우터비고"].apply(clean_router_text)

    # 날짜형
    df["라우터개통일_dt"] = pd.to_datetime(df["라우터개통일"], errors="coerce")
    df["라우터명의이전일_dt"] = pd.to_datetime(df["라우터명의이전일"], errors="coerce")

    # 표시용 상태 보정
    df["라우터명의이전상태"] = df["라우터명의이전상태"].replace("", "미입력")
    df["라우터청구대상"] = df["라우터청구대상"].replace("", "미입력")
    df["라우터사용"] = df["라우터사용"].replace("", "미입력")

    return df


def get_router_charge_target_df(base_df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    if base_df.empty:
        return pd.DataFrame()

    target_ym = f"{int(year):04d}-{int(month):02d}"
    df = base_df.copy()

    # 청구 대상 기본 조건
    df = df[
        (df["라우터사용"] == "예") &
        (df["라우터청구대상"].isin(["청구", "청구대상"])) &
        (df["라우터월비용"] > 0)
    ].copy()

    if df.empty:
        return df

    # 시작월 조건
    start_ok = df["라우터청구시작월"].astype(str).str.strip().eq("") | (df["라우터청구시작월"] <= target_ym)
    df = df[start_ok].copy()

    if df.empty:
        return df

    # 종료월 조건
    end_ok = df["라우터청구종료월"].astype(str).str.strip().eq("") | (df["라우터청구종료월"] >= target_ym)
    df = df[end_ok].copy()

    if df.empty:
        return df

    # 이전완료 제외
    df = df[df["라우터명의이전상태"] != "이전완료"].copy()

    return df


def get_router_issue_df(base_df: pd.DataFrame) -> pd.DataFrame:
    if base_df.empty:
        return pd.DataFrame()

    df = base_df.copy()

    cond = (
        (df["라우터사용"] == "예") &
        (
            ((df["라우터청구대상"].isin(["청구", "청구대상"])) & (df["라우터월비용"] <= 0)) |
            ((df["라우터명의이전상태"] == "이전완료") & (df["라우터청구대상"] == "청구")) |
            ((df["라우터명의이전상태"].isin(["이전대기", "이전거부"])) & (df["라우터청구시작월"] == "")) |
            ((df["라우터사용"] == "예") & (df["라우터청구대상"] == "미입력")) |
            ((df["라우터사용"] == "예") & (df["라우터명의이전상태"] == "미입력"))
        )
    )

    return df[cond].copy()

def get_router_warning_df(base_df: pd.DataFrame, overdue_days: int = 30) -> pd.DataFrame:
    if base_df.empty:
        return pd.DataFrame()

    df = base_df.copy()
    today = pd.Timestamp(datetime.today().date())

    df["개통후경과일"] = None
    if "라우터개통일_dt" in df.columns:
        df["개통후경과일"] = df["라우터개통일_dt"].apply(
            lambda x: (today - x).days if pd.notna(x) else None
        )

    cond_waiting_long = (
        (df["라우터사용"] == "예") &
        (df["라우터명의이전상태"].isin(["이전대기", "이전거부"])) &
        (df["개통후경과일"].notna()) &
        (df["개통후경과일"] >= overdue_days)
    )

    cond_claim_no_amount = (
        (df["라우터사용"] == "예") &
        (df["라우터청구대상"].isin(["청구", "청구대상"])) &
        (pd.to_numeric(df["라우터월비용"], errors="coerce").fillna(0) <= 0)
    )

    cond_done_but_claim = (
        (df["라우터사용"] == "예") &
        (df["라우터명의이전상태"] == "이전완료") &
        (df["라우터청구대상"].isin(["청구", "청구대상"]))
    )

    cond_missing_start = (
        (df["라우터사용"] == "예") &
        (df["라우터청구대상"].isin(["청구", "청구대상"])) &
        (df["라우터청구시작월"].astype(str).str.strip() == "")
    )

    warn_df = df[
        cond_waiting_long | cond_claim_no_amount | cond_done_but_claim | cond_missing_start
    ].copy()

    if warn_df.empty:
        return warn_df

    def make_warning_reason(row):
        reasons = []

        days = row.get("개통후경과일", None)
        status = str(row.get("라우터명의이전상태", "")).strip()
        target = str(row.get("라우터청구대상", "")).strip()
        amount = pd.to_numeric(row.get("라우터월비용", 0), errors="coerce")
        start_ym = str(row.get("라우터청구시작월", "")).strip()

        if status in ["이전대기", "이전거부"] and pd.notna(days) and days >= overdue_days:
            reasons.append(f"개통 후 {int(days)}일 경과")

        if target == "청구" and (pd.isna(amount) or float(amount) <= 0):
            reasons.append("청구대상인데 월비용 0")

        if status == "이전완료" and target == "청구":
            reasons.append("이전완료인데 청구중")

        if target == "청구" and start_ym == "":
            reasons.append("청구시작월 미입력")

        return " / ".join(reasons)

    warn_df["경고사유"] = warn_df.apply(make_warning_reason, axis=1)
    return warn_df


def build_router_claim_export_df(base_df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    target_df = get_router_charge_target_df(base_df, year, month)

    if target_df.empty:
        return pd.DataFrame(columns=[
            "청구년월", "관리코드", "단지명", "담당자", "지역", "수량",
            "라우터명의이전상태", "라우터청구대상", "라우터월비용", "비고"
        ])

    export_df = target_df.copy()
    export_df["청구년월"] = f"{int(year):04d}-{int(month):02d}"

    export_df = export_df[[
        "청구년월",
        "관리코드_표시",
        "단지명_표시",
        "담당자_표시",
        "지역_표시",
        "수량_표시",
        "라우터명의이전상태",
        "라우터청구대상",
        "라우터월비용",
        "라우터비고",
    ]].copy()

    export_df = export_df.rename(columns={
        "관리코드_표시": "관리코드",
        "단지명_표시": "단지명",
        "담당자_표시": "담당자",
        "지역_표시": "지역",
        "수량_표시": "수량",
        "라우터비고": "비고",
    })

    return export_df

def page_router_management():
    st.title("📡 아이센서 라우터 관리")

    if st.session_state.business != "아이센서":
        st.info("라우터 관리는 아이센서 사업에서만 사용합니다.")
        return

    base_df = build_router_base_df()

    if base_df.empty:
        st.warning("계약단지 데이터가 없습니다.")
        return

    current_year = datetime.today().year
    current_month = datetime.today().month
    current_ym = f"{current_year:04d}-{current_month:02d}"

    router_df = base_df[base_df["라우터사용"] == "예"].copy()
    waiting_df = router_df[router_df["라우터명의이전상태"] == "이전대기"].copy()
    rejected_df = router_df[router_df["라우터명의이전상태"] == "이전거부"].copy()
    finished_df = router_df[router_df["라우터명의이전상태"] == "이전완료"].copy()
    charge_target_df = get_router_charge_target_df(base_df, current_year, current_month)
    issue_df = get_router_issue_df(base_df)
    warning_df = get_router_warning_df(base_df, overdue_days=30)

    total_charge_cost = int(charge_target_df["라우터월비용"].sum()) if not charge_target_df.empty else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("라우터 사용 단지", len(router_df))
    c2.metric("이전대기", len(waiting_df))
    c3.metric("이전거부", len(rejected_df))
    c4.metric("이전완료", len(finished_df))
    c5.metric(f"{current_ym} 청구대상건수", len(charge_target_df))
    c6.metric(f"{current_ym} 청구총액", f"{total_charge_cost:,}")
        # 수금 기준 KPI
    billing_df, unpaid_df, manager_df = load_billing_dashboard_data()

    month_billing_df = billing_df[billing_df["기준월"].astype(str).str.strip() == current_ym].copy() if not billing_df.empty else pd.DataFrame()
    month_unpaid_df = unpaid_df[unpaid_df["기준월"].astype(str).str.strip() == current_ym].copy() if not unpaid_df.empty else pd.DataFrame()

    month_unpaid_amount = 0
    month_unpaid_count = 0

    if not month_unpaid_df.empty:
        month_unpaid_amount = int(pd.to_numeric(month_unpaid_df["미수금"], errors="coerce").fillna(0).sum())
        month_unpaid_count = len(month_unpaid_df)

    k1, k2 = st.columns(2)
    k1.metric(f"{current_ym} 미수금", f"{month_unpaid_amount:,}")
    k2.metric(f"{current_ym} 미입금건수", month_unpaid_count)

    if not warning_df.empty:
        st.error(f"경고 항목 {len(warning_df)}건이 있습니다. 아래 '경고/누락 확인'에서 확인하세요.")
    elif not issue_df.empty:
        st.warning(f"입력 보완 필요 항목 {len(issue_df)}건이 있습니다.")
    else:
        st.success("라우터 데이터 상태가 양호합니다.")

    st.divider()

    f1, f2, f3, f4 = st.columns(4)

    manager_options = ["전체"] + sorted([x for x in router_df["담당자_표시"].astype(str).unique().tolist() if x.strip() != ""])
    status_options = ["전체"] + sorted([x for x in router_df["라우터명의이전상태"].astype(str).unique().tolist() if x.strip() != ""])
    billing_options = ["전체"] + sorted([x for x in router_df["라우터청구대상"].astype(str).unique().tolist() if x.strip() != ""])
    region_options = ["전체"] + sorted([x for x in router_df["지역_표시"].astype(str).unique().tolist() if x.strip() != ""])

    sel_manager = f1.selectbox("담당자", manager_options, key="router_manager_filter")
    sel_status = f2.selectbox("명의이전상태", status_options, key="router_status_filter")
    sel_billing = f3.selectbox("청구대상", billing_options, key="router_billing_filter")
    sel_region = f4.selectbox("지역", region_options, key="router_region_filter")

    keyword = st.text_input("단지명 / 관리코드 / 비고 검색", key="router_keyword")

    view_df = router_df.copy()

    if sel_manager != "전체":
        view_df = view_df[view_df["담당자_표시"] == sel_manager]
    if sel_status != "전체":
        view_df = view_df[view_df["라우터명의이전상태"] == sel_status]
    if sel_billing != "전체":
        view_df = view_df[view_df["라우터청구대상"] == sel_billing]
    if sel_region != "전체":
        view_df = view_df[view_df["지역_표시"] == sel_region]

    if keyword.strip():
        kw = keyword.strip()
        view_df = view_df[
            view_df["단지명_표시"].astype(str).str.contains(kw, case=False, na=False) |
            view_df["관리코드_표시"].astype(str).str.contains(kw, case=False, na=False) |
            view_df["라우터비고"].astype(str).str.contains(kw, case=False, na=False)
        ].copy()

    st.subheader("1. 라우터 전체 현황")
    show_cols = [
        "관리코드_표시", "단지명_표시", "담당자_표시", "지역_표시", "수량_표시",
        "라우터사용", "라우터개통일", "라우터명의이전상태", "라우터명의이전일",
        "라우터청구대상", "라우터월비용", "라우터청구시작월", "라우터청구종료월", "라우터비고"
    ]
    show_df = view_df[show_cols].copy()
    show_df = show_df.rename(columns={
        "관리코드_표시": "관리코드",
        "단지명_표시": "단지명",
        "담당자_표시": "담당자",
        "지역_표시": "지역",
        "수량_표시": "수량",
    })
    styled_dataframe(show_df)

    st.divider()

    st.subheader(f"2. {current_ym} 청구 대상")
    if charge_target_df.empty:
        st.info("이번달 청구 대상이 없습니다.")
    else:
        target_show = charge_target_df[[
            "관리코드_표시", "단지명_표시", "담당자_표시", "지역_표시",
            "라우터명의이전상태", "라우터청구대상", "라우터월비용",
            "라우터청구시작월", "라우터청구종료월", "라우터비고"
        ]].copy()
        target_show = target_show.rename(columns={
            "관리코드_표시": "관리코드",
            "단지명_표시": "단지명",
            "담당자_표시": "담당자",
            "지역_표시": "지역",
        })
        st.dataframe(target_show, use_container_width=True, hide_index=True)
        download_section(f"{current_ym}_라우터청구대상", target_show, f"라우터청구대상_{current_ym}")

    st.divider()

    st.subheader("3. 청구 엑셀 자동 생성")
    g1, g2 = st.columns(2)
    claim_year = g1.number_input("청구 연도", min_value=2024, max_value=2100, value=current_year, step=1, key="router_claim_year")
    claim_month = g2.selectbox("청구 월", list(range(1, 13)), index=current_month - 1, key="router_claim_month")

    claim_export_df = build_router_claim_export_df(base_df, int(claim_year), int(claim_month))
    claim_ym = f"{int(claim_year):04d}-{int(claim_month):02d}"

    if claim_export_df.empty:
        st.info(f"{claim_ym} 청구 생성 대상이 없습니다.")
    else:
        st.success(f"{claim_ym} 청구 생성 대상 {len(claim_export_df)}건 / 합계 {int(claim_export_df['라우터월비용'].sum()):,}원")
        st.dataframe(claim_export_df, use_container_width=True, hide_index=True)

        b1, b2 = st.columns([1, 1])

        with b1:
            download_section(
                f"{claim_ym}_라우터청구리스트",
                claim_export_df,
                f"라우터청구리스트_{claim_ym}"
            )

        with b2:
            if st.button(f"{claim_ym} 이번달 청구 생성", key=f"create_billing_{claim_ym}"):
                result = add_monthly_billing_data(claim_export_df)

                if result["added_count"] > 0:
                    st.success(
                        f"수금관리 생성 완료: 신규 {result['added_count']}건 / "
                        f"중복 제외 {result['duplicate_count']}건"
                    )
                else:
                    st.warning(
                        f"신규 생성 없음. 이미 모두 생성된 월입니다. "
                        f"(중복 제외 {result['duplicate_count']}건)"
                    )
                if st.session_state.get("google_update_msg"):
                                    st.info(st.session_state["google_update_msg"])
                                    
                st.rerun()


        st.markdown("### 3-1. 수금관리")
        current_msg = st.session_state.get("google_update_msg", "")
        if current_msg:
            st.error(current_msg)
            st.session_state["google_update_msg"] = ""

        billing_df, unpaid_df, manager_df = load_billing_dashboard_data()

        # 기존 카드/요약
        if billing_df.empty:
            st.info("수금관리 데이터가 없습니다.")
        else:
            total_unpaid = int(pd.to_numeric(billing_df["미수금"], errors="coerce").fillna(0).sum())
            unpaid_count = int((pd.to_numeric(billing_df["미수금"], errors="coerce").fillna(0) > 0).sum())

            s1, s2 = st.columns(2)
            s1.metric("총 미수금", f"{total_unpaid:,}")
            s2.metric("미입금건수", unpaid_count)

            month_billing_df = billing_df[
                billing_df["기준월"].astype(str).str.strip() == claim_ym
            ].copy()

            month_billing_df["입금여부"] = month_billing_df["입금여부"].replace("", "미입금")

            if month_billing_df.empty:
                st.info(f"{claim_ym} 수금관리 데이터가 없습니다.")
            else:
                st.dataframe(month_billing_df, use_container_width=True, hide_index=True)

                st.markdown("### 입금 처리")

                unpaid_candidates = month_billing_df[
                    pd.to_numeric(month_billing_df["미수금"], errors="coerce").fillna(0) > 0
                ].copy()

                if unpaid_candidates.empty:
                    st.success("이 달의 청구 건은 모두 입금 처리되었습니다.")
                else:
                    unpaid_candidates = month_unpaid_df.copy()

                    unpaid_candidates["청구금액"] = pd.to_numeric(
                        unpaid_candidates["청구금액"], errors="coerce"
                    ).fillna(0)

                    unpaid_candidates["미수금"] = pd.to_numeric(
                        unpaid_candidates["미수금"], errors="coerce"
                    ).fillna(unpaid_candidates["청구금액"]).fillna(0)

                    unpaid_candidates["표시금액"] = unpaid_candidates["미수금"]
                    unpaid_candidates.loc[unpaid_candidates["표시금액"] <= 0, "표시금액"] = unpaid_candidates["청구금액"]

                    unpaid_candidates["표시금액"] = unpaid_candidates["표시금액"].astype(int)

                    unpaid_candidates["선택표시"] = (
                        unpaid_candidates["단지명"].astype(str).str.strip()
                        + " / "
                        + unpaid_candidates["담당자"].astype(str).str.strip()
                        + " / "
                        + unpaid_candidates["표시금액"].astype(str).str.strip()
                        + "원"
                    )

                    selected_label = st.selectbox(
                        "입금 처리할 항목 선택",
                        unpaid_candidates["선택표시"].tolist(),
                        key=f"paid_select_{claim_ym}"
                    )

                    selected_row = unpaid_candidates[unpaid_candidates["선택표시"] == selected_label].iloc[0]

                    if st.button("선택 항목 입금 처리", key=f"mark_paid_{claim_ym}"):
                        st.session_state["google_update_msg"] = ""

                        amount_raw = selected_row.get("표시금액", "")
                        amount_num = pd.to_numeric(amount_raw, errors="coerce")

                        if pd.isna(amount_num):
                            st.error(f"청구금액이 비어 있거나 숫자가 아닙니다: {amount_raw}")
                        else:
                            ok = mark_billing_paid(
                                기준월=str(selected_row.get("기준월", "")).strip(),
                                단지명=str(selected_row.get("단지명", "")).strip(),
                                담당자=str(selected_row.get("담당자", "")).strip(),
                                청구금액=int(amount_num)
                            )

                            if ok:
                                st.success("입금 처리 완료")
                                st.rerun()
                            else:
                                st.error("입금 처리 실패")  

        st.markdown("### 3-2. 미입금관리")
        if unpaid_df.empty:
            st.success("현재 미입금 항목이 없습니다.")
        else:
            month_unpaid_df = unpaid_df[unpaid_df["기준월"].astype(str).str.strip() == claim_ym].copy()
            st.dataframe(month_unpaid_df, use_container_width=True, hide_index=True)

        st.markdown("### 3-3. 담당자별현황")
        if manager_df.empty:
            st.info("담당자별현황 데이터가 없습니다.")
        else:
            st.dataframe(manager_df, use_container_width=True, hide_index=True)

    st.divider()

    st.subheader("6. 경고 / 누락 확인")
    if warning_df.empty:
        st.success("긴급 경고 항목이 없습니다.")
    else:
        warn_show = warning_df[[
            "관리코드_표시", "단지명_표시", "담당자_표시", "지역_표시",
            "라우터명의이전상태", "라우터청구대상", "라우터월비용",
            "라우터청구시작월", "라우터청구종료월", "경고사유"
        ]].copy()
        warn_show = warn_show.rename(columns={
            "관리코드_표시": "관리코드",
            "단지명_표시": "단지명",
            "담당자_표시": "담당자",
            "지역_표시": "지역",
        })
        st.dataframe(warn_show, use_container_width=True, hide_index=True)
        download_section("라우터_경고항목", warn_show, "라우터_경고항목")

    st.divider()
    st.info("※ 현재 화면은 계약단지 구글시트의 라우터 컬럼을 읽어서 보여주는 관리 화면입니다. 값 수정은 원본 구글시트에서 진행해 주세요.")

# =========================================================
# 메인
# =========================================================
def main():
    init_files()

    st.sidebar.title("메뉴")
    st.sidebar.write(f"로그인: {st.session_state.username}")
    st.sidebar.write(f"이름: {current_user_name()}")
    st.sidebar.write(f"권한: {'관리자' if is_admin() else '담당자'}")

    selected_business = st.sidebar.selectbox("사업 선택", list(BUSINESS_CONFIG.keys()))
    st.session_state.business = selected_business

    if st.sidebar.button("로그아웃"):
        logout()

    render_header()

    menus = BUSINESS_CONFIG[st.session_state.business]["menus"].copy()
    if is_admin():
        menus.append("관리자 도구")

    menu = st.sidebar.radio("선택", menus)

    if menu == "대시보드":
        page_dashboard()

    elif menu == "데이터 가져오기":
        page_import()

    elif menu == "영업현황":
        generic_data_page(
            "📞 영업현황",
            "영업현황",
            [("담당자", ["담당자", "영업담당"]), ("지역", "지역"), ("상품", "상품"), ("진행여부", "진행여부")],
            "sales_search",
        )

    elif menu == "가능단지":
        generic_data_page(
            "✅ 가능단지",
            "가능단지",
            [("영업담당", ["담당자", "영업담당"]), ("지역", "지역"), ("상품", "상품"), ("결과", "결과")],
            "possible_search",
        )

    elif menu == "입찰공고단지":
        generic_data_page(
            "📝 입찰공고단지",
            "입찰공고",
            [("영업담당", ["담당자", "영업담당"]), ("지역", "지역"), ("상품", "상품"), ("판매형태", "판매형태")],
            "bid_search",
        )

    elif menu == "계약단지":
        generic_data_page(
            "📦 계약단지",
            "계약단지",
            [("영업담당", ["담당자", "영업담당"]), ("지역", "지역"), ("상품", "상품"), ("시공여부", "시공여부")],
            "contract_search",
        )

    elif menu == "라우터 관리":
        page_router_management()    

    elif menu == "계약접수현황":
        generic_data_page(
            "🔋 충전기 계약접수현황",
            "계약접수현황",
            [
                ("담당자", ["담당자", "영업담당", "실사담당"]),
                ("지역", "지역"),
                ("구분", "구분"),
                ("운영사", "운영사"),
            ],
            "ev_contract_search",
        )

    elif menu == "오늘 할 일":
        page_tasks()

    elif menu == "일정 관리":
        page_schedule()

    elif menu == "영업 알림":
        page_alerts()

    elif menu == "관리자 도구":
        page_admin_tools()


if not st.session_state.logged_in:
    login()
else:
    main()
