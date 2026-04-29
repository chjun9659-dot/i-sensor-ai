import os
import re
import io
from datetime import datetime, date, timedelta

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
# 👇 여기 추가 (이 위치가 핵심)
def render_common_style():
    st.markdown("""               
    <style>
    .main {
        background-color: #f1f5f9 !important;
    }
    .yw-card {
        background: #ffffff;
        border-radius: 14px;
        padding: 16px;
        border: 1px solid #e2e8f0;

        /* 🔥 여기 핵심 */
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);

        border-left: 5px solid transparent;
        min-height: 80px;
        margin-bottom: 8px;
        transition: all 0.2s ease;
    }

    /* 카드 hover */
    .yw-card:hover {
        transform: translateY(-4px) scale(1.01);
        box-shadow: 0 10px 20px rgba(0,0,0,0.12);
    }

    /* expander hover */
    div[data-testid="stExpander"] {
        border-radius: 12px;
        transition: all 0.2s ease;
    }

    div[data-testid="stExpander"] > div {
        background: #ffffff;
        border-radius: 12px;
    }

    /* 클릭 느낌 */
    div[data-testid="stExpanderHeader"] {
        cursor: pointer;
    }
                
    /* Streamlit metric 카드 hover */
    div[data-testid="stMetric"] {
        background: #ffffff !important;
        border-radius: 14px !important;
        border: 1px solid #e2e8f0 !important;
        padding: 14px !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.04) !important;
    }

    div[data-testid="stMetric"]:hover {
        transform: translateY(-4px) scale(1.01);
        box-shadow: 0 10px 20px rgba(0,0,0,0.12) !important;
    }

    /* 컬러 라인 */
    .yw-card.success { border-left-color: #16a34a; }
    .yw-card.warning { border-left-color: #f59e0b; }
    .yw-card.danger  { border-left-color: #ef4444; }
    .yw-card.info    { border-left-color: #2563eb; }

    /* 제목 */
    .card-title {
        font-size: 13px;
        font-weight: 700;
        color: #64748b;
        margin-bottom: 3px;
    }

    /* 숫자 */
    .card-value {
        font-size: 24px;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.2;
    }

    /* 설명 */
    .card-sub {
        font-size: 11px;
        color: #94a3b8;
        margin-top: 3px;
    }
    </style>
    """, unsafe_allow_html=True)
    
def ui_card(title, value, sub="", status=""):
    st.markdown(f"""
    <div class="yw-card {status}">
        <div class="card-title">{title}</div>
        <div class="card-value">{value}</div>
        <div class="card-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)

st.set_page_config(page_title="윤우 영업 통합 시스템", layout="wide")

# =========================================================
# 기본 설정
# =========================================================
DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

# =========================================================
# 사용자관리 구글시트
# =========================================================
USER_SHEET_URL = "https://docs.google.com/spreadsheets/d/1uUjrdRwTjdvKoED1dWKsikAOGWNxMR59Eht-SYdRX1w/edit?gid=0#gid=0"


def load_users_from_gsheet():
    """
    사용자관리 시트 컬럼:
    아이디 / 비밀번호 / 권한 / 사용여부 / 이름 / 부서 / 직급 / 코드
    """
    try:
        if not USER_SHEET_URL or "여기에_" in USER_SHEET_URL:
            return {}

        client = get_gsheet_client()
        sheet_id = re.search(r"/d/([a-zA-Z0-9-_]+)", USER_SHEET_URL).group(1)
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.get_worksheet(0)

        values = worksheet.get_all_values()

        if not values or len(values) < 2:
            st.error("사용자관리 시트에 데이터가 없습니다.")
            return {}

        headers = [str(x).strip() for x in values[0]]
        rows = values[1:]

        df = pd.DataFrame(rows, columns=headers).fillna("")
        df.columns = [str(c).strip() for c in df.columns]

        required_cols = ["아이디", "비밀번호", "권한", "사용여부", "이름"]
        for col in required_cols:
            if col not in df.columns:
                st.error(f"사용자관리 시트에 '{col}' 컬럼이 없습니다.")
                return {}

        users = {}

        for _, row in df.iterrows():
            user_id = str(row.get("아이디", "")).strip()
            password = str(row.get("비밀번호", "")).strip()
            role = str(row.get("권한", "")).strip()   # 🔥 핵심 수정
            use_yn = str(row.get("사용여부", "")).strip().upper()
            name = str(row.get("이름", "")).strip()

            if not user_id:
                continue

            users[user_id] = {
                "pw": password,
                "role": role,   # 🔥 그대로 넣기
                "name": name,
                "use_yn": use_yn,
                "department": str(row.get("부서", "")).strip(),
                "position": str(row.get("직급", "")).strip(),
                "code": str(row.get("코드", "")).strip(),
            }

        return users

    except Exception as e:
        st.error(f"사용자관리 시트 불러오기 오류: {e}")
        return {}

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
            "연차관리": "https://docs.google.com/spreadsheets/d/1n7AXfaCIljI8cBMSX7DrLjPVEQPTTCXL2-rxqcLzepE/edit?gid=0#gid=0",
            "시공일정": "https://docs.google.com/spreadsheets/d/1-6P8Orzas1U6W-Rmv7pcx-N0-fiPnH-Ah20jDEsRXgs/edit?gid=0#gid=0",
            "실사관리": "https://docs.google.com/spreadsheets/d/1rV9nZWGQDgUBgxldvojixj8Ys5DsefH9wa5-IWg_t34/edit?gid=859568227#gid=859568227",
        },
        "menus": [
            "대시보드",
            "데이터 가져오기",

            "영업현황",
            "가능단지",
            "입찰공고단지",
            "계약단지",
            "라우터 관리",
            "아이센서 유지보수관리",

            "연차 관리",
            "시공 일정",
            "실사 관리",

            "오늘 할 일",
            "일정 관리",
            "영업 알림",
        ],
    },
        "전기차 충전기": {
        "sheets": {
            "계약접수현황": "https://docs.google.com/spreadsheets/d/1Efuld8DUSHs8jR42R5XKIT3y-3P_pf--aP9ED_SYWhc/edit?gid=784400128#gid=784400128",
            "시공일정": "https://docs.google.com/spreadsheets/d/1-6P8Orzas1U6W-Rmv7pcx-N0-fiPnH-Ah20jDEsRXgs/edit?gid=0#gid=0",
            "실사관리": "https://docs.google.com/spreadsheets/d/1rV9nZWGQDgUBgxldvojixj8Ys5DsefH9wa5-IWg_t34/edit?gid=859568227#gid=859568227",
        },
        "menus": [
            "대시보드",
            "데이터 가져오기",
            "계약접수현황",
            "시공 일정",
            "실사 관리",
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

if "department" not in st.session_state:
    st.session_state.department = ""

if "position" not in st.session_state:
    st.session_state.position = ""

if "user_code" not in st.session_state:
    st.session_state.user_code = ""

if "inspection_form_version" not in st.session_state:
    st.session_state.inspection_form_version = 0

if "inspection_edit_mode" not in st.session_state:
    st.session_state.inspection_edit_mode = False

if "inspection_edit_target" not in st.session_state:
    st.session_state.inspection_edit_target = None    


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
    import os
    import gspread
    from google.oauth2.service_account import Credentials
    import streamlit as st

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    base_dir = os.path.dirname(os.path.abspath(__file__))
    key_path = os.path.join(base_dir, "service_account.json")

    # 1) 배포 우선: Streamlit secrets 시도
    try:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(dict(creds_dict), scopes=scopes)
        client = gspread.authorize(creds)

        st.session_state["google_auth_debug"] = {
            "mode": "streamlit_secrets",
            "client_email": creds.service_account_email,
        }
        return client

    except Exception as secrets_error:
        # 2) 로컬 fallback: service_account.json 이 실제 있을 때만 사용
        if os.path.exists(key_path):
            creds = Credentials.from_service_account_file(key_path, scopes=scopes)
            client = gspread.authorize(creds)

            st.session_state["google_auth_debug"] = {
                "mode": "local_json",
                "key_path": key_path,
                "client_email": creds.service_account_email,
                "secrets_error": str(secrets_error),
            }
            return client

        # 3) 둘 다 없으면 실제 원인을 보여줌
        raise RuntimeError(
            f"구글 인증 실패 / secrets 오류: {secrets_error} / "
            f"로컬 JSON 없음: {key_path}"
        )

    raise FileNotFoundError(
        "구글 인증 정보를 찾지 못했습니다. "
        "배포환경은 st.secrets['gcp_service_account'], "
        "로컬환경은 service_account.json 이 필요합니다."
    )

    
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

    if sheet_name == "연차관리":
        client = get_gsheet_client()
        sheet_id = re.search(r"/d/([a-zA-Z0-9-_]+)", url).group(1)
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.get_worksheet(0)

        values = worksheet.get_all_values()

        if not values or len(values) < 2:
            return pd.DataFrame()

        headers = [str(x).strip() for x in values[0]]
        rows = values[1:]

        df = pd.DataFrame(rows, columns=headers).fillna("")
        df.columns = [str(c).strip() for c in df.columns]

        return df

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

    return df   # 🔥 이거 추가

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
        existing_row_map = {}

        for idx, row in enumerate(rows, start=2):
            row_기준월 = str(row[0]).strip() if len(row) > 0 else ""
            row_단지명 = str(row[1]).strip() if len(row) > 1 else ""
            row_담당자 = str(row[2]).strip() if len(row) > 2 else ""
            row_청구금액 = str(row[3]).strip() if len(row) > 3 else ""

            key = f"{row_기준월}||{row_단지명}||{row_담당자}"

            existing_row_map[key] = {
                "row_index": idx,
                "amount": row_청구금액
            }

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

            if key in existing_row_map:
                row_info = existing_row_map[key]
                old_amount = row_info["amount"]
                row_index = row_info["row_index"]

                # 🔥 핵심: 금액 없으면 업데이트
                if old_amount in ["", "0"]:
                    worksheet.update_cell(row_index, 4, 청구금액)  # D열
                    worksheet.update_cell(row_index, 5, "")        # E열 초기화
                else:
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
            existing_row_map[key] = {
                "row_index": None,
                "amount": str(청구금액)
            }

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
        users = load_users_from_gsheet()

        if not users:
            st.error("사용자 정보를 불러오지 못했습니다. 사용자관리 시트를 확인하세요.")
            return

        if user_id not in users:
            st.error("아이디 또는 비밀번호가 맞지 않습니다.")
            return

        user = users[user_id]

        if str(user.get("use_yn", "")).upper() != "Y":
            st.error("사용 중지된 계정입니다. 관리자에게 문의하세요.")
            return

        if user.get("pw") != password:
            st.error("아이디 또는 비밀번호가 맞지 않습니다.")
            return

        st.session_state.logged_in = True
        st.session_state.username = user_id
        st.session_state.role = user.get("role", "user")
        st.session_state.display_name = user.get("name", user_id)
        st.session_state.department = user.get("department", "")
        st.session_state.position = user.get("position", "")
        st.session_state.user_code = user.get("code", "")

        st.rerun()


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
        st.markdown("""
        <style>
        div[data-testid="stMetricValue"] {
            font-size: 22px !important;
        }

        div[data-testid="stMetricLabel"] {
            font-size: 14px !important;
        }

        h1, h2, h3 {
            font-size: 22px !important;
        }
        </style>
        """, unsafe_allow_html=True)

    st.divider()

def save_vacation_log(action, target_name="", use_date="", used_days="", reason="", note=""):
    """
    연차 사용/취소/수정/재계산 로그 저장
    """
    try:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        actor = st.session_state.get("username", "")
        if not actor:
            actor = st.session_state.get("name", "")
        if not actor:
            actor = "알수없음"

        sh = get_google_sheet()
        try:
            ws = sh.worksheet(VACATION_LOG_SHEET_NAME)
        except:
            ws = sh.add_worksheet(title=VACATION_LOG_SHEET_NAME, rows=1000, cols=8)
            ws.append_row(["기록일시", "작업자", "작업구분", "대상직원", "사용일자", "사용일수", "사유", "비고"])

        ws.append_row([
            now,
            actor,
            action,
            str(target_name),
            str(use_date),
            str(used_days),
            str(reason),
            str(note),
        ])

    except Exception as e:
        st.warning(f"연차 로그 저장 중 오류가 발생했습니다: {e}")

def save_vacation_data_to_excel(df: pd.DataFrame):
    df = df.copy()

    # ✅ 숫자 컬럼 강제 float 처리 (반차 0.5 대응)
    for col in ["발생 연차", "사용 연차", "잔여 연차"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(float)

    wb = load_workbook(VACATION_FILE_PATH)
    ws = wb[VACATION_SHEET_NAME]

    header_row = 2     # 실제 헤더 행
    start_row = 3      # 실제 데이터 시작 행

    # 엑셀 헤더 읽기
    excel_headers = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        excel_headers.append(str(v).strip() if v is not None else "")

    # 헤더명 -> 엑셀 컬럼번호
    col_map = {name: idx for idx, name in enumerate(excel_headers, start=1) if name}

    # df에 있는 컬럼만 엑셀에서 지우기
    last_row = max(ws.max_row, start_row + len(df) + 50)
    for r in range(start_row, last_row + 1):
        for col_name in df.columns:
            if col_name in col_map:
                ws.cell(row=r, column=col_map[col_name]).value = None

    # 헤더명 기준으로 정확히 저장
    for row_idx, (_, row) in enumerate(df.iterrows(), start=start_row):
        for col_name in df.columns:
            if col_name in col_map:
                value = row[col_name]

                if pd.isna(value):
                    value = None

                ws.cell(row=row_idx, column=col_map[col_name]).value = value

    wb.save(VACATION_FILE_PATH)


# =========================================================
# 연차 관리 설정
# =========================================================
VACATION_BACKUP_DIR = "backup"
VACATION_FILE_PATH = "data/vacation.csv"
USE_COLS = [f"사용일{i}" for i in range(1, 70)]
VACATION_LOG_SHEET_NAME = "연차사용로그"

def to_number(value, default=0):
    num = pd.to_numeric(value, errors="coerce")
    return default if pd.isna(num) else float(num)

def format_leave_number(value):
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return ""
    if float(num).is_integer():
        return str(int(num))
    return str(num)

def format_display_date(value):
    try:
        dt = pd.to_datetime(value, errors="coerce")
        if pd.isna(dt):
            return ""
        return dt.strftime("%Y-%m-%d")
    except:
        return ""
    
def format_leave_date(use_date, leave_type):
    date_str = pd.to_datetime(use_date).strftime("%Y-%m-%d")
    if leave_type == "반차":
        return f"{date_str} (반차)"
    return date_str   
    
def get_target_year():
    return datetime.today().year    

def calculate_service_years(hire_date, base_date):
    years = base_date.year - hire_date.year
    if (base_date.month, base_date.day) < (hire_date.month, hire_date.day):
        years -= 1
    return max(0, years)


def calculate_anniversary_period(hire_date, target_year):
    try:
        start = date(target_year, hire_date.month, hire_date.day)
    except ValueError:
        if hire_date.month == 2 and hire_date.day == 29:
            start = date(target_year, 2, 28)
        else:
            raise

    try:
        end = date(target_year + 1, hire_date.month, hire_date.day) - timedelta(days=1)
    except ValueError:
        if hire_date.month == 2 and hire_date.day == 29:
            end = date(target_year + 1, 2, 28) - timedelta(days=1)
        else:
            raise

    return start, end


def calculate_auto_leave_days(hire_date, target_year=None):
    today = date.today()

    hire_date = pd.to_datetime(hire_date, errors="coerce")
    if pd.isna(hire_date):
        return None, None, 0, 0

    hire_date = hire_date.date()

    def safe_replace_year(d, year):
        try:
            return d.replace(year=year)
        except ValueError:
            if d.month == 2 and d.day == 29:
                return date(year, 2, 28)
            raise

    one_year_anniversary = safe_replace_year(hire_date, hire_date.year + 1)

    if today < one_year_anniversary:
        start_date = hire_date
        end_date = one_year_anniversary - timedelta(days=1)
        service_years = 0

        months_worked = (today.year - hire_date.year) * 12 + (today.month - hire_date.month)
        if today.day < hire_date.day:
            months_worked -= 1

        months_worked = max(0, min(11, months_worked))
        leave_days = float(months_worked)

    else:
        this_year_anniversary = safe_replace_year(hire_date, today.year)

        if today >= this_year_anniversary:
            start_date = this_year_anniversary
        else:
            start_date = safe_replace_year(hire_date, today.year - 1)

        end_date = safe_replace_year(start_date, start_date.year + 1) - timedelta(days=1)

        service_years = today.year - hire_date.year
        if (today.month, today.day) < (hire_date.month, hire_date.day):
            service_years -= 1

        extra_days = max(0, (service_years - 1) // 2)
        leave_days = float(min(25, 15 + extra_days))

    return start_date, end_date, service_years, leave_days


def parse_use_entry(value):
    if pd.isna(value):
        return None, None

    text = str(value).strip()
    if text == "" or text.lower() == "none":
        return None, None

    amount = 0.5 if "반차" in text else 1.0

    clean = text.replace("(반차)", "").strip()
    clean = clean.replace(".", "-").replace(" ", "")

    if len(clean) == 8 and clean[2] == "-":
        clean = "20" + clean

    parsed_date = pd.to_datetime(clean, errors="coerce")

    if pd.isna(parsed_date):
        return None, None

    return parsed_date, amount


def parse_cancel_amount(value):
    text = str(value)
    return 0.5 if "반차" in text else 1.0


def recalculate_vacation_summary(df: pd.DataFrame):
    for idx in df.index:
        total_leave = to_number(df.loc[idx, "발생 연차"])
        used_leave = 0.0

        start_date = pd.to_datetime(df.loc[idx, "기산시작일"], errors="coerce")
        end_date = pd.to_datetime(df.loc[idx, "기산종료일"], errors="coerce")

        for col in USE_COLS:
            if col not in df.columns:
                continue

            value = df.loc[idx, col]

            if pd.isna(value):
                continue

            text = str(value).strip()
            if text == "" or text.lower() == "none":
                continue

            parsed_date, amount = parse_use_entry(value)
            if parsed_date is None:
                continue

            if pd.notna(start_date) and pd.notna(end_date):
                if start_date.date() <= parsed_date.date() <= end_date.date():
                    used_leave += amount
            else:
                used_leave += amount

        remain_leave = total_leave - used_leave

        df.loc[idx, "사용 연차"] = format_leave_number(used_leave)
        df.loc[idx, "잔여 연차"] = format_leave_number(remain_leave)

    return df


def recalculate_all_vacation_data(df: pd.DataFrame):
    df = df.copy()

    for idx in df.index:
        hire_date = pd.to_datetime(df.loc[idx, "입사일"], errors="coerce")

        if pd.isna(hire_date):
            continue

        hire_date = hire_date.date()

        start_date, end_date, service_years, leave_days = calculate_auto_leave_days(hire_date)

        df.loc[idx, "기산시작일"] = str(start_date)
        df.loc[idx, "기산종료일"] = str(end_date)
        df.loc[idx, "근속년수"] = int(service_years)
        df.loc[idx, "발생 연차"] = float(leave_days)

    df = recalculate_vacation_summary(df)
    return df

def refresh_expired_vacation_rows(df: pd.DataFrame):
    """
    기산종료일이 지난 직원만 자동 갱신
    전체 재정리가 아니라 만료된 직원 행만 처리
    """
    df = df.copy()
    today = date.today()
    changed = False
    changed_names = []

    for idx in df.index:
        end_date = pd.to_datetime(df.loc[idx, "기산종료일"], errors="coerce")

        if pd.isna(end_date):
            continue

        if today > end_date.date():
            hire_date = pd.to_datetime(df.loc[idx, "입사일"], errors="coerce")

            if pd.isna(hire_date):
                continue

            start_date, new_end_date, service_years, leave_days = calculate_auto_leave_days(hire_date.date())

            df.loc[idx, "기산시작일"] = str(start_date)
            df.loc[idx, "기산종료일"] = str(new_end_date)
            df.loc[idx, "근속년수"] = int(service_years)
            df.loc[idx, "발생 연차"] = float(leave_days)

            changed = True
            changed_names.append(str(df.loc[idx, "이름"]))

    if changed:
        df = recalculate_vacation_summary(df)

    return df, changed, changed_names

def build_monthly_stats(df, target_year, target_month):
    rows = []
    total_count = 0
    total_amount = 0.0

    for _, row in df.iterrows():
        emp_name = str(row.get("이름", "")).strip()
        emp_count = 0
        emp_amount = 0.0

        for col in USE_COLS:
            value = row.get(col, None)
            parsed_date, amount = parse_use_entry(value)
            if parsed_date is None:
                continue

            if parsed_date.year == int(target_year) and parsed_date.month == int(target_month):
                emp_count += 1
                emp_amount += amount

        if emp_count > 0:
            rows.append({
                "이름": emp_name,
                "사용 건수": emp_count,
                "사용 일수": format_leave_number(emp_amount)
            })
            total_count += emp_count
            total_amount += emp_amount

    return pd.DataFrame(rows), total_count, total_amount
    

def render_employee_vacation_cards(df: pd.DataFrame):
    st.subheader("직원별 연차 요약 카드")
    st.caption("전체 직원 기준")

    if df.empty:
        st.info("표시할 직원 데이터가 없습니다.")
        return

    card_df = df.copy()

    for col in ["발생 연차", "사용 연차", "잔여 연차"]:
        if col not in card_df.columns:
            card_df[col] = 0
        card_df[col] = pd.to_numeric(card_df[col], errors="coerce").fillna(0)

    card_df["사용률"] = card_df.apply(
        lambda row: 0 if float(row["발생 연차"]) <= 0 else round((float(row["사용 연차"]) / float(row["발생 연차"])) * 100, 1),
        axis=1
    )

    card_df = card_df.sort_values(by=["잔여 연차", "사용률"], ascending=[True, False]).reset_index(drop=True)

    cols_per_row = 4

    for start in range(0, len(card_df), cols_per_row):
        row_cols = st.columns(cols_per_row)

        for i in range(cols_per_row):
            idx = start + i
            if idx >= len(card_df):
                row_cols[i].empty()
                continue

            row = card_df.iloc[idx]

            name = str(row["이름"]).strip()
            total = float(row["발생 연차"])
            used = float(row["사용 연차"])
            remain = float(row["잔여 연차"])
            rate = float(row["사용률"])

            if remain <= 0:
                status_text = "🔴 위험"
                border_color = "#ef4444"
                bg_color = "#fef2f2"
            elif remain <= 5:
                status_text = "🟡 주의"
                border_color = "#f59e0b"
                bg_color = "#fffbeb"
            else:
                status_text = "🟢 정상"
                border_color = "#22c55e"
                bg_color = "#f0fdf4"

            row_cols[i].markdown(
                f"""
                <div style="
                    border: 2px solid {border_color};
                    background: {bg_color};
                    border-radius: 14px;
                    padding: 16px;
                    margin-bottom: 12px;
                    min-height: 170px;
                ">
                    <div style="font-size:20px; font-weight:800; margin-bottom:10px;">
                        {name}
                    </div>
                    <div style="font-size:15px; line-height:1.9;">
                        • 발생 연차: <b>{format_leave_number(total)}일</b><br>
                        • 사용 연차: <b>{format_leave_number(used)}일</b><br>
                        • 잔여 연차: <b>{format_leave_number(remain)}일</b><br>
                        • 사용률: <b>{rate}%</b><br>
                        • 상태: <b>{status_text}</b>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

def style_remaining_leave(val):
    num = pd.to_numeric(val, errors="coerce")
    if pd.isna(num):
        return ""
    if num <= 0:
        return "background-color: #f8d7da; color: #842029; font-weight: bold;"
    elif num <= 5:
        return "background-color: #fff3cd; color: #664d03; font-weight: bold;"
    return ""

def find_first_empty_use_col(row, df_columns):
    for col in USE_COLS:
        matching_indexes = [i for i, c in enumerate(df_columns) if str(c).strip() == col]

        for col_idx in matching_indexes:
            value = row.iloc[col_idx]

            if pd.isna(value) or clean_text(value) == "" or clean_text(value).lower() == "none":
                return col_idx

    return None    

def create_backup():
    os.makedirs(VACATION_BACKUP_DIR, exist_ok=True)

    df = load_df("연차관리")
    backup_path = os.path.join(
        VACATION_BACKUP_DIR,
        f"연차관리_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    )

    df.to_csv(backup_path, index=False, encoding="utf-8-sig")
    return backup_path

def save_vacation_data(df):
    sheet_urls = get_current_sheet_urls()
    url = sheet_urls.get("연차관리", "")

    if not url:
        raise Exception("연차관리 구글시트 URL이 없습니다.")

    client = get_gsheet_client()
    sheet_id = re.search(r"/d/([a-zA-Z0-9-_]+)", url).group(1)
    spreadsheet = client.open_by_key(sheet_id)
    worksheet = spreadsheet.get_worksheet(0)

    save_df = df.copy()

    if "row_id" in save_df.columns:
        save_df = save_df.drop(columns=["row_id"])

    values = worksheet.get_all_values()
    if not values:
        raise Exception("연차관리 시트에 헤더가 없습니다.")

    headers = [str(x).strip() for x in values[0]]
    headers = [h for h in headers if h != ""]

    for col in headers:
        if col not in save_df.columns:
            save_df[col] = ""

    save_df = save_df[headers].copy()
    save_df = save_df.where(pd.notnull(save_df), "")

    for col in save_df.columns:
        save_df[col] = save_df[col].apply(
            lambda x: x.strftime("%Y-%m-%d") if isinstance(x, (datetime, date, pd.Timestamp)) else str(x).strip()
        )

    rows = save_df.values.tolist()

    worksheet.update("A2", rows, value_input_option="USER_ENTERED")

    old_rows = max(0, len(values) - 1)
    new_rows = len(rows)

    if old_rows > new_rows:
        blank_rows = old_rows - new_rows
        clear_start = new_rows + 2
        empty_rows = [[""] * len(headers) for _ in range(blank_rows)]
        worksheet.update(f"A{clear_start}", empty_rows)

    st.cache_data.clear()

# =========================================================
# 시공 일정 시스템
# =========================================================
SCHEDULE_SHEET_NAME = "시공일정"
EXPECTED_COLUMNS = ["날짜", "상품구분", "설치현장", "시공담당", "수량", "비고", "상태", "완료일"]

def get_schedule_sheet():
    url = get_current_sheet_urls().get("시공일정")
    sheet_id = re.search(r"/d/([a-zA-Z0-9-_]+)", url).group(1)
    client = get_gsheet_client()
    return client.open_by_key(sheet_id).sheet1

def append_schedule_data(new_row_df):
    sheet = get_schedule_sheet()

    save_df = new_row_df.copy()

    for col in EXPECTED_COLUMNS:
        if col not in save_df.columns:
            save_df[col] = ""

    save_df = save_df[EXPECTED_COLUMNS].fillna("")

    rows = save_df.astype(str).values.tolist()

    sheet.append_rows(rows, value_input_option="USER_ENTERED")

    st.cache_data.clear()    

def ensure_schedule_sheet_header(sheet):
    values = sheet.get_all_values()

    if not values:
        sheet.update("A1:H1", [EXPECTED_COLUMNS])
        return

    header = values[0]

    if header != EXPECTED_COLUMNS:
        # 👉 기존 데이터 유지하면서 컬럼만 맞춤
        df = pd.DataFrame(values[1:], columns=header)

        for col in EXPECTED_COLUMNS:
            if col not in df.columns:
                df[col] = ""

        df = df[EXPECTED_COLUMNS]

        # 👉 헤더만 수정 (데이터 유지)
        sheet.update("A1:H1", [EXPECTED_COLUMNS])

@st.cache_data(ttl=60)
def load_schedule_data():
    sheet = get_schedule_sheet()
    ensure_schedule_sheet_header(sheet)

    data = sheet.get_all_records()
    df = pd.DataFrame(data)

    if df.empty:
        return pd.DataFrame(columns=EXPECTED_COLUMNS)

    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[EXPECTED_COLUMNS]

    df["수량"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0).astype(int)
    df["날짜"] = df["날짜"].astype(str)
    df["완료일"] = df["완료일"].astype(str)
    df["상태"] = df["상태"].astype(str).replace("", "진행중")

    return df


def save_schedule_data(df, sheet=None):
    if sheet is None:
        sheet = get_schedule_sheet()

    save_df = df.copy()

    for col in EXPECTED_COLUMNS:
        if col not in save_df.columns:
            save_df[col] = ""

    save_df = save_df[EXPECTED_COLUMNS].fillna("")
    if list(save_df.columns) != EXPECTED_COLUMNS:
        st.error("컬럼 구조 이상 - 저장 중단")
        return
    # ✅ 필수값 정리
    save_df["날짜"] = save_df["날짜"].astype(str).str.strip()
    save_df["상품구분"] = save_df["상품구분"].astype(str).str.strip()
    save_df["설치현장"] = save_df["설치현장"].astype(str).str.strip()
    save_df["시공담당"] = save_df["시공담당"].astype(str).str.strip()
    save_df["상태"] = save_df["상태"].astype(str).str.strip()

    # ✅ 빈 행 저장 방지
    save_df = save_df[save_df["날짜"] != ""].copy()
    save_df = save_df[save_df["설치현장"] != ""].copy()
    save_df = save_df[save_df["시공담당"] != ""].copy()

    # ✅ 상태값 보정
    save_df.loc[~save_df["상태"].isin(["진행중", "완료"]), "상태"] = "진행중"
    save_df["수량"] = pd.to_numeric(save_df["수량"], errors="coerce").fillna(0).astype(int)

    # 기존 시트 데이터 길이 확인
    old_values = sheet.get_all_values()
    old_data_rows = max(0, len(old_values) - 1)  # 헤더 제외

    # 새로 저장할 데이터
    rows = [save_df.columns.tolist()] + save_df.astype(str).values.tolist()

    # 1) 새 데이터 저장
    sheet.update("A1", rows)

    # 2) 예전 데이터가 더 길었다면 남는 행 비우기
    new_data_rows = len(save_df)

    if old_data_rows > new_data_rows:
        blank_rows = old_data_rows - new_data_rows
        start_row = new_data_rows + 2   # 헤더가 1행이므로 실제 데이터 시작 보정
        end_row = old_data_rows + 1

        clear_range = f"A{start_row}:H{end_row}"
        empty_values = [[""] * len(EXPECTED_COLUMNS) for _ in range(blank_rows)]
        sheet.update(clear_range, empty_values)

    st.cache_data.clear()

SCHEDULE_LOG_COLUMNS = ["시간", "사용자", "사업", "작업", "설치현장", "시공담당", "비고"]

def find_original_schedule_index(full_df, target_row):
    full_df = full_df[EXPECTED_COLUMNS].copy().fillna("")

    for col in EXPECTED_COLUMNS:
        full_df[col] = full_df[col].astype(str).str.strip()

    mask = (
        (full_df["날짜"] == str(target_row["날짜"]).strip()) &
        (full_df["상품구분"] == str(target_row["상품구분"]).strip()) &
        (full_df["설치현장"] == str(target_row["설치현장"]).strip()) &
        (full_df["시공담당"] == str(target_row["시공담당"]).strip()) &
        (full_df["수량"].astype(str) == str(target_row["수량"]).strip()) &
        (full_df["비고"] == str(target_row["비고"]).strip()) &
        (full_df["상태"] == str(target_row["상태"]).strip())
    )

    matched = full_df[mask]

    if matched.empty:
        return None

    return matched.index[0]

def save_schedule_log(action, site="", manager="", note=""):
    try:
        log_path = os.path.join(DATA_DIR, "schedule_work_log.csv")

        user = str(st.session_state.get("display_name", "")).strip()
        business = str(st.session_state.get("business", "")).strip()

        new_log = pd.DataFrame([{
            "시간": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "사용자": user,
            "사업": business,
            "작업": action,
            "설치현장": site,
            "시공담당": manager,
            "비고": note,
        }])

        if os.path.exists(log_path):
            old_log = pd.read_csv(log_path, encoding="utf-8-sig").fillna("")
            log_df = pd.concat([old_log, new_log], ignore_index=True)
        else:
            log_df = new_log

        log_df = log_df[SCHEDULE_LOG_COLUMNS]
        log_df.to_csv(log_path, index=False, encoding="utf-8-sig")

    except Exception as e:
        st.warning(f"작업 로그 저장 실패: {e}")    

def schedule_page():
    render_inspection_common_style()
    render_common_style()

    st.markdown('<div class="erp-page-title">시공 일정 관리 프로그램</div>', unsafe_allow_html=True)
    st.markdown('<div class="erp-page-desc">시공 일정 등록, 수정, 진행 현황 관리</div>', unsafe_allow_html=True)


    def style_schedule_status(val):
        s = str(val).strip()

        if s == "완료":
            return "background-color: #dcfce7; color: #166534; font-weight: 700;"

        if s == "진행중":
            return "background-color: #fef3c7; color: #92400e; font-weight: 700;"

        return ""

    try:
        df = load_schedule_data()

        if df is None:
            df = pd.DataFrame(columns=EXPECTED_COLUMNS)

        if df.empty:
            df = pd.DataFrame(columns=EXPECTED_COLUMNS)
        else:
            df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
            df["완료일"] = pd.to_datetime(df["완료일"], errors="coerce")

            df = df.dropna(subset=["날짜"])

            df["날짜"] = df["날짜"].dt.strftime("%Y-%m-%d")
            df["완료일"] = df["완료일"].dt.strftime("%Y-%m-%d")

            df["날짜"] = df["날짜"].fillna("")
            df["완료일"] = df["완료일"].fillna("")

            # ✅ 사업별 필터
            if "상품구분" in df.columns:
                product_series = df["상품구분"].astype(str).str.strip()

                if st.session_state.business == "아이센서":
                    df = df[product_series.str.contains("아이센서", na=False)].copy()

                elif st.session_state.business == "전기차 충전기":
                    df = df[product_series.str.contains("전기차", na=False)].copy()

            # ✅ 권한 필터: 관리자는 전체 / 직원은 본인 담당 일정만
            login_role = str(st.session_state.get("role", "")).strip()
            login_name = str(st.session_state.get("display_name", "")).strip()

            if login_role != "관리자" and "시공담당" in df.columns:
                df = df[df["시공담당"].astype(str).str.strip() == login_name].copy()

            # ✅ 날짜 정렬
            if "날짜" in df.columns:
                df = df.sort_values("날짜", ascending=True).reset_index(drop=True)      

    except Exception as e:
        st.error(f"시공일정 데이터를 불러오지 못했습니다: {e}")
        return

    df = df.reset_index(drop=True)
    df["row_id"] = df.index

    today_str = str(date.today())

    total_count = len(df)
    today_count = len(df[df["날짜"] == today_str])
    progress_count = len(df[df["상태"] == "진행중"])
    done_count = len(df[df["상태"] == "완료"])
    total_qty = int(df["수량"].sum()) if not df.empty else 0

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        ui_card("전체 일정", total_count)

    with c2:
        ui_card("오늘 일정", today_count)

    with c3:
        ui_card("진행중", progress_count)

    with c4:
        ui_card("완료", done_count)

    with c5:
        ui_card("총 수량", total_qty)

    st.divider()

    with st.expander("📅 1. 시공 일정 등록", expanded=False):
        with st.form("add_schedule_form_unique"):

            a1, a2, a3 = st.columns(3)
            work_date = a1.date_input("시공 날짜", value=date.today(), key="sch_work_date_unique")
            site_name = a2.text_input("설치현장", key="sch_site_name_unique")
            manager_name = a3.text_input("시공담당", key="sch_manager_name_unique")
            product = st.selectbox(
                "상품구분",
                ["아이센서", "전기차충전기"],
                key="sch_product_unique"
            )

            a4, a5 = st.columns(2)
            quantity = a4.number_input("수량", min_value=0, step=1, value=0, key="sch_quantity_unique")
            note = a5.text_input("비고", key="sch_note_unique")

            submitted = st.form_submit_button("등록하기")

            if submitted:
                if not site_name.strip():
                    st.warning("설치현장을 입력해주세요.")
                elif not manager_name.strip():
                    st.warning("시공담당을 입력해주세요.")
                else:
                    new_row = pd.DataFrame([{
                        "날짜": str(work_date),
                        "상품구분": product,
                        "설치현장": site_name.strip(),
                        "시공담당": manager_name.strip(),
                        "수량": int(quantity),
                        "비고": note.strip(),
                        "상태": "진행중",
                        "완료일": ""
                    }])

                    append_schedule_data(new_row)

                    save_schedule_log(
                        "등록",
                        site=site_name.strip(),
                        manager=manager_name.strip(),
                        note=f"수량 {int(quantity)}"
                    )

                    st.success("등록 완료!")
                    st.rerun()

    st.divider()
    with st.expander("📅 2. 오늘 일정", expanded=False):
        today_df = df[df["날짜"] == today_str].copy()

        if today_df.empty:
            st.info("오늘 일정이 없습니다.")
        else:
            show_today = today_df[["날짜", "설치현장", "시공담당", "수량", "비고", "상태", "완료일"]]
            styled_today_df = show_today.style.map(style_schedule_status, subset=["상태"])

            st.dataframe(
                styled_today_df,
                use_container_width=True,
                hide_index=True
            )

    st.divider()
    with st.expander("📋 3. 시공 일정 보기", expanded=False):
        managers = ["전체"] + sorted([m for m in df["시공담당"].dropna().unique().tolist() if str(m).strip() != ""])

        f1, f2, f3, f4 = st.columns(4)
        status_filter = f1.selectbox("상태 선택", ["전체", "진행중", "완료"], key="sch_status_filter_unique")
        manager_filter = f2.selectbox("담당자 선택", managers, key="sch_manager_filter_unique")
        date_filter = f3.selectbox("날짜 기준", ["전체", "오늘", "미래", "지난 일정"], key="sch_date_filter_unique")
        keyword = f4.text_input("검색", placeholder="설치현장 / 비고 검색", key="sch_keyword_unique")

        filtered_df = df.copy()

        if status_filter != "전체":
            filtered_df = filtered_df[filtered_df["상태"] == status_filter]

        if manager_filter != "전체":
            filtered_df = filtered_df[filtered_df["시공담당"] == manager_filter]

        if date_filter == "오늘":
            filtered_df = filtered_df[filtered_df["날짜"] == today_str]
        elif date_filter == "미래":
            filtered_df = filtered_df[filtered_df["날짜"] > today_str]
        elif date_filter == "지난 일정":
            filtered_df = filtered_df[filtered_df["날짜"] < today_str]

        if keyword.strip():
            kw = keyword.strip()
            filtered_df = filtered_df[
                filtered_df["설치현장"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["비고"].astype(str).str.contains(kw, case=False, na=False)
            ]

        show_df = filtered_df[["날짜", "설치현장", "시공담당", "수량", "비고", "상태", "완료일"]].copy()

        if show_df.empty:
            st.info("조건에 맞는 일정이 없습니다.")
        else:
            styled_show_df = show_df.style.map(style_schedule_status, subset=["상태"])

            st.dataframe(
                styled_show_df,
                use_container_width=True,
                hide_index=True
            )

    st.divider()

    with st.expander("✏️ 4. 일정 수정", expanded=False):
        if df.empty:
            st.info("수정할 일정이 없습니다.")
        else:
            edit_options = [
                f"{row['row_id']} | {row['날짜']} | {row['설치현장']} | {row['시공담당']}"
                for _, row in df.iterrows()
            ]

            selected_edit = st.selectbox("수정할 일정 선택", edit_options, key="sch_edit_select_unique")
            edit_idx = int(selected_edit.split("|")[0].strip())
            edit_row = df.loc[df["row_id"] == edit_idx].iloc[0]

            default_edit_date = (
                pd.to_datetime(edit_row["날짜"]).date()
                if str(edit_row["날짜"]).strip()
                else date.today()
            )

            with st.form(f"edit_schedule_form_{edit_idx}"):
                e1, e2, e3 = st.columns(3)
                edit_date = e1.date_input("시공 날짜 수정", value=default_edit_date)
                edit_site = e2.text_input("설치현장 수정", value=str(edit_row["설치현장"]))
                edit_manager = e3.text_input("시공담당 수정", value=str(edit_row["시공담당"]))

                e4, e5, e6 = st.columns(3)
                edit_qty = e4.number_input("수량 수정", min_value=0, step=1, value=int(edit_row["수량"]))
                edit_note = e5.text_input("비고 수정", value=str(edit_row["비고"]))
                edit_status = e6.selectbox(
                    "상태 수정",
                    ["진행중", "완료"],
                    index=0 if str(edit_row["상태"]) == "진행중" else 1
                )

                edit_submit = st.form_submit_button("수정 저장")

                if edit_submit:
                    full_df = load_schedule_data()
                    full_df = full_df[EXPECTED_COLUMNS].copy()

                    original_idx = find_original_schedule_index(full_df, edit_row)

                    if original_idx is None:
                        st.error("원본 구글시트에서 해당 일정을 찾지 못했습니다. 저장을 중단합니다.")
                        st.stop()

                    full_df.loc[original_idx, "날짜"] = str(edit_date)
                    full_df.loc[original_idx, "설치현장"] = edit_site.strip()
                    full_df.loc[original_idx, "시공담당"] = edit_manager.strip()
                    full_df.loc[original_idx, "수량"] = int(edit_qty)
                    full_df.loc[original_idx, "비고"] = edit_note.strip()
                    full_df.loc[original_idx, "상태"] = edit_status

                    if edit_status == "완료" and not str(full_df.loc[original_idx, "완료일"]).strip():
                        full_df.loc[original_idx, "완료일"] = today_str
                    elif edit_status == "진행중":
                        full_df.loc[original_idx, "완료일"] = ""

                    save_schedule_data(full_df)

                    save_schedule_log(
                        "수정",
                        site=edit_site.strip(),
                        manager=edit_manager.strip(),
                        note=f"상태 {edit_status} / 수량 {int(edit_qty)}"
                    )

                    st.success("수정 완료!")
                    st.rerun()

    st.divider()

    with st.expander("✅ 5. 완료 처리", expanded=False):
        progress_df = df[df["상태"] == "진행중"].copy()

        if progress_df.empty:
            st.info("완료 처리할 일정이 없습니다.")
        else:
            complete_options = [
                f"{row['row_id']} | {row['날짜']} | {row['설치현장']} | {row['시공담당']} | 수량 {row['수량']}"
                for _, row in progress_df.iterrows()
            ]

            selected_complete = st.selectbox("완료 처리 일정 선택", complete_options, key="sch_complete_select_unique")

            if st.button("완료로 변경", key="sch_complete_btn_unique"):
                complete_idx = int(selected_complete.split("|")[0].strip())
                target_row = df.loc[df["row_id"] == complete_idx].iloc[0]

                full_df = load_schedule_data()
                full_df = full_df[EXPECTED_COLUMNS].copy()

                original_idx = find_original_schedule_index(full_df, target_row)

                if original_idx is None:
                    st.error("원본 구글시트에서 해당 일정을 찾지 못했습니다. 저장을 중단합니다.")
                    st.stop()

                full_df.loc[original_idx, "상태"] = "완료"
                full_df.loc[original_idx, "완료일"] = today_str

                save_schedule_data(full_df)

                save_schedule_log(
                    "완료",
                    site=str(target_row["설치현장"]),
                    manager=str(target_row["시공담당"]),
                    note=f"완료일 {today_str}"
                )

                st.success("완료 처리되었습니다.")
                st.rerun()

    st.divider()

    with st.expander("↩️ 6. 완료 취소", expanded=False):
        done_df = df[df["상태"] == "완료"].copy()

        if done_df.empty:
            st.info("완료 취소할 일정이 없습니다.")
        else:
            cancel_options = [
                f"{row['row_id']} | {row['날짜']} | {row['설치현장']} | {row['시공담당']} | 수량 {row['수량']}"
                for _, row in done_df.iterrows()
            ]

            selected_cancel = st.selectbox("완료 취소 일정 선택", cancel_options, key="sch_cancel_select_unique")

            if st.button("진행중으로 변경", key="sch_cancel_btn_unique"):
                cancel_idx = int(selected_cancel.split("|")[0].strip())
                target_row = df.loc[df["row_id"] == cancel_idx].iloc[0]

                full_df = load_schedule_data()
                full_df = full_df[EXPECTED_COLUMNS].copy()

                original_idx = find_original_schedule_index(full_df, target_row)

                if original_idx is None:
                    st.error("원본 구글시트에서 해당 일정을 찾지 못했습니다. 저장을 중단합니다.")
                    st.stop()

                full_df.loc[original_idx, "상태"] = "진행중"
                full_df.loc[original_idx, "완료일"] = ""

                save_schedule_data(full_df)

                st.success("완료 취소되었습니다.")
                st.rerun()

    st.divider()

    with st.expander("🗑️ 7. 일정 삭제", expanded=False):
        if df.empty:
            st.info("삭제할 일정이 없습니다.")
        else:
            delete_options = [
                f"{row['row_id']} | {row['날짜']} | {row['설치현장']} | {row['시공담당']} | 수량 {row['수량']}"
                for _, row in df.iterrows()
            ]

            selected_delete = st.selectbox("삭제할 일정 선택", delete_options, key="sch_delete_select_unique")

            if st.button("선택 일정 삭제", key="sch_delete_btn_unique"):
                delete_idx = int(selected_delete.split("|")[0].strip())
                target_row = df.loc[df["row_id"] == delete_idx].iloc[0]

                full_df = load_schedule_data()
                full_df = full_df[EXPECTED_COLUMNS].copy()

                original_idx = find_original_schedule_index(full_df, target_row)

                if original_idx is None:
                    st.error("원본 구글시트에서 해당 일정을 찾지 못했습니다. 삭제를 중단합니다.")
                    st.stop()

                full_df = full_df.drop(index=original_idx).reset_index(drop=True)

                save_schedule_data(full_df)

                save_schedule_log(
                    "삭제",
                    site=str(target_row["설치현장"]),
                    manager=str(target_row["시공담당"]),
                    note=f"날짜 {target_row['날짜']} / 수량 {target_row['수량']}"
                )

                st.success("삭제 완료!")
                st.rerun()

def vacation_page():

    st.markdown('<div class="erp-page-title">윤우테크 연차 관리 프로그램</div>', unsafe_allow_html=True)
    st.markdown('<div class="erp-page-desc">연차 관리 및 현황 확인 프로그램입니다.</div>', unsafe_allow_html=True)

    login_id = str(
        st.session_state.get("user_id")
        or st.session_state.get("username")
        or ""
    ).strip()

    login_role = str(
        st.session_state.get("role")
        or st.session_state.get("user_role")
        or st.session_state.get("권한")
        or ""
    ).strip()

    is_admin = login_role == "관리자"

    try:
        df = load_df("연차관리")

        # ✅ 기산종료일 지난 직원만 자동 갱신
        df, vacation_changed, changed_names = refresh_expired_vacation_rows(df)

        if vacation_changed:
            backup_file = create_backup()
            save_vacation_data(df)
            st.cache_data.clear()
            st.info(
                f"기산기간이 지난 직원 연차가 자동 갱신되었습니다: "
                f"{', '.join(changed_names)} / 백업: {backup_file}"
            )
            st.rerun()

        df = apply_role_filter(df)

        users = load_users_from_gsheet()

        if not users:
            st.error("사용자관리 정보를 불러오지 못했습니다.")
            return

        if login_id not in users:
            st.error(f"사용자관리 시트에서 로그인 ID '{login_id}'를 찾지 못했습니다.")
            return

        login_name = str(users[login_id].get("name", "")).strip()

        if is_admin:
            df_view = df.copy()
        else:
            df_view = df[df["이름"] == login_name].copy()

    except Exception as e:
        st.error(f"연차 파일을 불러오지 못했습니다: {e}")
        return

    # =====================================================
    # 관리 도구
    # =====================================================
    if is_admin:
        st.subheader("🛠️ 관리 도구")

        tool_col1, tool_col2, tool_col3 = st.columns(3)

        with tool_col1:
            if st.button("💾 지금 백업하기", use_container_width=True, key="vac_backup_btn_unique"):
                backup_file = create_backup()
                st.success(f"백업 완료: {backup_file}")

        with tool_col2:
            confirm_recalc = st.checkbox("정말 전체 연차를 재정리하고 구글시트에 저장합니다. 백업 후에만 체크하세요.")

            if st.button("📊 연차 수치 재정리", use_container_width=True, key="vac_recalc_btn_unique"):
                if not confirm_recalc:
                    st.warning("체크 후 실행하세요.")
                else:
                    df = recalculate_all_vacation_data(df)
                    save_vacation_data(df)
                    st.cache_data.clear()
                    st.success("전체 재계산 완료")
                    st.rerun()

        with tool_col3:
            if os.path.exists(VACATION_BACKUP_DIR):
                backup_files = sorted(os.listdir(VACATION_BACKUP_DIR), reverse=True)
                st.write(f"백업 파일 수: {len(backup_files)}")
            else:
                st.write("백업 파일 수: 0")

    # =====================================================
    # 직원 선택
    # =====================================================
    st.subheader("👤 직원 선택")

    names = sorted(df_view["이름"].dropna().astype(str).unique().tolist())

    if is_admin:
        search_name = st.text_input("직원 검색", placeholder="이름을 입력하세요", key="vac_search_name_unique")

        if search_name:
            filtered_names = [n for n in names if search_name.strip().lower() in n.lower()]
        else:
            filtered_names = names

        if not filtered_names:
            st.warning("검색 결과가 없습니다.")
            return

        selected_name = st.selectbox("직원 선택", filtered_names, key="vac_selected_name_unique")
    else:
        selected_name = login_name
        st.info(f"본인 연차만 조회됩니다: {selected_name}")

    employee_rows = df_view[df_view["이름"] == selected_name]

    if employee_rows.empty:
        st.error(f"연차관리 시트에서 '{selected_name}' 직원을 찾지 못했습니다.")
        return

    employee = employee_rows.iloc[0]

    # =====================================================
    # 현재 연차 현황
    # =====================================================
    st.subheader("📌 현재 연차 현황")

    col1, col2, col3 = st.columns(3)

    total = to_number(employee["발생 연차"])
    used = to_number(employee["사용 연차"])
    remain = to_number(employee["잔여 연차"])

    col1.metric("총 연차", format_leave_number(total))
    col2.metric("사용 연차", format_leave_number(used))
    col3.metric("잔여 연차", format_leave_number(remain))

    if remain <= 0:
        st.error("잔여 연차가 없습니다.")
    elif remain <= 5:
        st.warning("잔여 연차가 5일 이하입니다.")
    else:
        st.success("잔여 연차가 충분합니다.")

    if is_admin:
        if st.button("🔄 선택 직원 연차 다시 계산", use_container_width=True, key="vac_recalc_selected_btn"):

            # ✅ 선택 직원의 실제 행 위치를 숫자로 찾기
            match_positions = [
                i for i, name in enumerate(df["이름"].astype(str).str.strip().tolist())
                if name == str(selected_name).strip()
            ]

            if not match_positions:
                st.error("선택한 직원을 찾지 못했습니다.")
                st.stop()

            row_pos = match_positions[0]

            # ✅ 컬럼 위치를 숫자로 찾기
            used_col_pos = list(df.columns).index("사용 연차")
            remain_col_pos = list(df.columns).index("잔여 연차")
            total_col_pos = list(df.columns).index("발생 연차")
            start_col_pos = list(df.columns).index("기산시작일")
            end_col_pos = list(df.columns).index("기산종료일")

            total_leave = to_number(df.iloc[row_pos, total_col_pos])
            used_leave = 0.0

            start_date = pd.to_datetime(df.iloc[row_pos, start_col_pos], errors="coerce")
            end_date = pd.to_datetime(df.iloc[row_pos, end_col_pos], errors="coerce")

            for col in USE_COLS:
                if col not in df.columns:
                    continue

                col_pos = list(df.columns).index(col)
                value = df.iloc[row_pos, col_pos]

                if pd.isna(value):
                    continue

                text = str(value).strip()
                if text == "" or text.lower() == "none":
                    continue

                parsed_date, amount = parse_use_entry(value)

                if parsed_date is None:
                    continue

                if pd.notna(start_date) and pd.notna(end_date):
                    if start_date.date() <= parsed_date.date() <= end_date.date():
                        used_leave += amount
                else:
                    used_leave += amount

            remain_leave = total_leave - used_leave

            # ✅ 배포앱 안전 처리: df 전체를 object 타입으로 변경
            df = df.astype(object)

            # ✅ 선택 직원 값 저장
            df.iloc[row_pos, used_col_pos] = format_leave_number(used_leave)
            df.iloc[row_pos, remain_col_pos] = format_leave_number(remain_leave)

            save_vacation_log(
                action="재계산",
                target_name=str(selected_name),
                use_date="",
                used_days="",
                reason="",
                note="선택 직원 연차 다시 계산"
            )
            st.cache_data.clear()
            st.success(f"{selected_name} 연차가 다시 계산되었습니다.")
            st.rerun()

    # =====================================================
    # 관리자: 직원별 요약 카드
    # =====================================================
    if is_admin:
        with st.expander("직원별 연차 요약 카드 (전체 직원)", expanded=False):
            render_employee_vacation_cards(df)

    # =====================================================
    # 연차 사용 입력
    # =====================================================
    if is_admin:    
        with st.expander("📝 연차 사용 입력", expanded=False):
            use_date = st.date_input("사용 날짜 선택", datetime.today(), key="vac_use_date_unique")
            leave_type = st.radio("사용 종류 선택", ["연차", "반차"], horizontal=True, key="vac_leave_type_unique")
            leave_amount = 1.0 if leave_type == "연차" else 0.5

            st.write(f"선택된 사용값: **{format_leave_number(leave_amount)}일**")

            btn_col1, btn_col2 = st.columns(2)

            with btn_col1:
                register_btn = st.button("등록하기", type="primary", use_container_width=True, key="vac_register_btn_unique")

            with btn_col2:
                preview_btn = st.button("미리 확인", use_container_width=True, key="vac_preview_btn_unique")

            if preview_btn:
                expected_used = used + leave_amount
                expected_remain = total - expected_used
                st.info(
                    f"{selected_name} / {use_date.strftime('%Y-%m-%d')} / {leave_type} 등록 시 "
                    f"사용 연차 {format_leave_number(expected_used)}, 잔여 연차 {format_leave_number(expected_remain)}"
                )

            if register_btn:
                idx = df[df["이름"] == selected_name].index[0]

                current_total = float(to_number(df.loc[idx, "발생 연차"]))
                current_used = float(to_number(df.loc[idx, "사용 연차"]))
                current_remain = float(to_number(df.loc[idx, "잔여 연차"]))

                if current_remain < leave_amount:
                    st.error("잔여 연차가 부족합니다.")
                else:
                    row_pos = df.index.get_loc(idx)
                    empty_col_idx = find_first_empty_use_col(df.iloc[row_pos], df.columns)

                    if empty_col_idx is None:
                        st.error("사용일 칸이 모두 찼습니다. 사용일1~사용일30을 확인해주세요.")
                    else:

                        # ✅ 기산기간 밖 연차 입력 차단
                        start_date = pd.to_datetime(df.iloc[row_pos]["기산시작일"], errors="coerce")
                        end_date = pd.to_datetime(df.iloc[row_pos]["기산종료일"], errors="coerce")

                        if pd.notna(start_date) and pd.notna(end_date):
                            if use_date < start_date.date() or use_date > end_date.date():
                                st.error(
                                    f"⚠️ 기산기간 외 연차입니다. "
                                    f"이 직원의 기산기간은 {start_date.date()} ~ {end_date.date()} 입니다."
                                )
                                st.stop()

                        df.iat[row_pos, empty_col_idx] = format_leave_date(use_date, leave_type)

                        target_idx = df.index[row_pos]

                        target_one = df.loc[[target_idx]].copy()
                        target_one = recalculate_vacation_summary(target_one)

                        df.loc[target_idx, "사용 연차"] = target_one.loc[target_idx, "사용 연차"]
                        df.loc[target_idx, "잔여 연차"] = target_one.loc[target_idx, "잔여 연차"]

                        save_vacation_data(df)

                        save_vacation_log(
                            action="등록",
                            target_name=str(selected_name),
                            use_date=str(use_date),
                            used_days=str(leave_amount),
                            reason=str(leave_type),
                            note="연차 사용 등록"
                        )

                        st.cache_data.clear()
                        st.success("연차 등록 완료!")
                        st.rerun()

    # =====================================================
    # 선택 직원 사용일 내역
    # =====================================================
    with st.expander("🗂️ 선택 직원 사용일 내역", expanded=False):
        use_list = []

        for col in USE_COLS:
            value = employee.get(col, None)
            display_value = ""

            if pd.notna(value) and clean_text(value) != "" and clean_text(value).lower() != "none":
                display_value = format_display_date(value)

                if display_value == "":
                    display_value = clean_text(value)

            use_list.append({
                "구분": col,
                "사용내역": display_value
            })

        use_df = pd.DataFrame(use_list)
        use_df = use_df[use_df["사용내역"] != ""]

        if not use_df.empty:
            st.dataframe(use_df, use_container_width=True)
        else:
            st.info("등록된 사용일이 없습니다.")

    # =====================================================
    # 연차 취소
    # =====================================================
    if is_admin:
        with st.expander("↩️ 연차 취소", expanded=False):
            use_list = []

            for col in USE_COLS:
                value = employee.get(col, None)
                display_value = ""

                if pd.notna(value) and clean_text(value) != "" and clean_text(value).lower() != "none":
                    display_value = format_display_date(value)

                    if display_value == "":
                        display_value = clean_text(value)

                use_list.append({
                    "구분": col,
                    "사용내역": display_value
                })

            use_df = pd.DataFrame(use_list)
            use_df = use_df[use_df["사용내역"] != ""]

            if not use_df.empty:
                cancel_options = [f"{row['구분']} | {row['사용내역']}" for _, row in use_df.iterrows()]
                selected_cancel = st.selectbox("취소할 사용일 선택", cancel_options, key="vac_cancel_select_unique")

                if st.button("선택 사용일 취소", use_container_width=True, key="vac_cancel_btn_unique"):
                    idx = df[df["이름"] == selected_name].index[0]

                    selected_col = selected_cancel.split("|")[0].strip()
                    selected_value = df.loc[idx, selected_col]

                    cancel_amount = parse_cancel_amount(selected_value)

                    current_total = to_number(df.loc[idx, "발생 연차"])
                    current_used = to_number(df.loc[idx, "사용 연차"])

                    new_used = max(0, current_used - cancel_amount)
                    new_remain = current_total - new_used

                    df = df.astype(object)

                    df.loc[idx, selected_col] = ""
                    df.loc[idx, "사용 연차"] = format_leave_number(new_used)
                    df.loc[idx, "잔여 연차"] = format_leave_number(new_remain)

                    save_vacation_data(df)

                    save_vacation_log(
                        action="취소",
                        target_name=str(selected_name),
                        use_date="",
                        used_days="",
                        reason="연차 사용 취소",
                        note="선택 사용일 취소"
                    )

                    st.cache_data.clear()
                    st.success("연차 취소 완료!")
                    st.rerun()
            else:
                st.info("취소할 사용일이 없습니다.")

    # =====================================================
    # 관리자 전용: 직원 관리
    # =====================================================
    if is_admin:
        with st.expander("📁 직원 관리", expanded=False):

            st.markdown("## ➕ 직원 추가")

            with st.form("add_employee_form_unique"):
                new_name = st.text_input("직원 이름", key="new_employee_name_unique")
                new_hire_date = st.date_input("입사일", value=date.today(), key="new_employee_hire_date_unique")

                preview_start, preview_end, preview_service_years, preview_leave_days = calculate_auto_leave_days(
                    new_hire_date,
                    get_target_year()
                )

                st.info(
                    f"자동 계산 결과\n\n"
                    f"- 기산시작일: {preview_start}\n"
                    f"- 기산종료일: {preview_end}\n"
                    f"- 근속년수: {preview_service_years}\n"
                    f"- 발생 연차: {format_leave_number(preview_leave_days)}일"
                )

                submit_add_employee = st.form_submit_button("직원 추가하기")

                if submit_add_employee:
                    new_name = new_name.strip()

                    if new_name == "":
                        st.error("직원 이름을 입력해주세요.")
                    elif new_name in df["이름"].astype(str).tolist():
                        st.error("이미 등록된 직원입니다.")
                    else:
                        hire_date = pd.to_datetime(new_hire_date).date()
                        start_date, end_date, service_years, auto_leave_days = calculate_auto_leave_days(
                            hire_date,
                            get_target_year()
                        )

                        new_row = {
                            "이름": new_name,
                            "입사일": pd.to_datetime(hire_date),
                            "기산시작일": pd.to_datetime(start_date),
                            "기산종료일": pd.to_datetime(end_date),
                            "근속년수": service_years,
                            "발생 연차": float(auto_leave_days),
                            "사용 연차": 0.0,
                            "잔여 연차": float(auto_leave_days),
                        }

                        for col in USE_COLS:
                            new_row[col] = None

                        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

                        save_vacation_data(df)
                        st.cache_data.clear()
                        st.success("직원 추가 완료!")
                        st.rerun()

            st.markdown("---")
            st.markdown("## ✏️ 직원 수정")

            edit_name = st.selectbox("수정할 직원 선택", names, key="edit_employee_select_unique")
            edit_employee = df[df["이름"] == edit_name].iloc[0]

            default_hire_date = pd.to_datetime(edit_employee["입사일"], errors="coerce")
            if pd.isna(default_hire_date):
                default_hire_date = pd.Timestamp(date.today())

            with st.form("edit_employee_form_unique"):
                edited_name = st.text_input("직원 이름 수정", value=str(edit_employee["이름"]), key="edited_name_unique")
                edited_hire_date = st.date_input("입사일 수정", value=default_hire_date.date(), key="edited_hire_date_unique")
                edited_used_leave = st.number_input(
                    "사용 연차 수정",
                    min_value=0.0,
                    step=0.5,
                    value=float(to_number(edit_employee["사용 연차"])),
                    key="edited_used_leave_unique"
                )

                preview_start, preview_end, preview_service_years, preview_leave_days = calculate_auto_leave_days(
                    pd.to_datetime(edited_hire_date).date(),
                    get_target_year()
                )

                st.info(
                    f"자동 계산 결과\n\n"
                    f"- 기산시작일: {preview_start}\n"
                    f"- 기산종료일: {preview_end}\n"
                    f"- 근속년수: {preview_service_years}\n"
                    f"- 발생 연차: {format_leave_number(preview_leave_days)}일"
                )

                submit_edit_employee = st.form_submit_button("직원 정보 수정하기")

                if submit_edit_employee:
                    edited_name = edited_name.strip()

                    if edited_name == "":
                        st.error("직원 이름을 입력해주세요.")
                    else:
                        duplicate_names = [n for n in df["이름"].astype(str).tolist() if n != edit_name]

                        if edited_name in duplicate_names:
                            st.error("같은 이름의 직원이 이미 있습니다.")
                        else:
                            idx = df[df["이름"] == edit_name].index[0]

                            hire_date = pd.to_datetime(edited_hire_date).date()
                            start_date, end_date, service_years, auto_leave_days = calculate_auto_leave_days(
                                hire_date,
                                get_target_year()
                            )

                            new_total = float(auto_leave_days)
                            new_used = float(edited_used_leave)
                            new_remain = new_total - new_used

                            if new_remain < 0:
                                st.error("사용 연차가 발생 연차보다 클 수 없습니다.")
                            else:
                                if "근속년수" not in df.columns:
                                    df["근속년수"] = ""

                                df["근속년수"] = pd.to_numeric(df["근속년수"], errors="coerce").fillna(0)
                                df["근속년수"] = df["근속년수"].astype(int)

                                row_pos = df.index.get_loc(idx)

                                df.loc[idx, "이름"] = edited_name
                                df.loc[idx, "입사일"] = str(hire_date)
                                df.loc[idx, "기산시작일"] = str(start_date)
                                df.loc[idx, "기산종료일"] = str(end_date)
                                df.loc[idx, "발생 연차"] = float(new_total)
                                df.loc[idx, "사용 연차"] = float(new_used)
                                df.loc[idx, "잔여 연차"] = float(new_remain)

                                service_col_pos = list(df.columns).index("근속년수")
                                df.iat[row_pos, service_col_pos] = int(service_years)

                                save_vacation_data(df)
                                st.cache_data.clear()
                                st.success("직원 정보 수정 완료!")
                                st.rerun()

            st.markdown("---")
            st.markdown("## 🗑️ 직원 삭제")

            delete_name = st.selectbox("삭제할 직원 선택", names, key="delete_employee_select_unique")
            confirm_delete = st.checkbox("정말 삭제합니다. 되돌리기 어렵습니다.", key="vac_confirm_delete_unique")

            if st.button("선택 직원 삭제", use_container_width=True, key="vac_delete_btn_unique"):
                if not confirm_delete:
                    st.warning("삭제 확인 체크를 먼저 해주세요.")
                else:
                    before_count = len(df)
                    df = df[df["이름"].astype(str) != str(delete_name)].copy()
                    after_count = len(df)

                    if before_count == after_count:
                        st.error("삭제할 직원을 찾지 못했습니다.")
                    else:
                        save_vacation_data(df)
                        st.cache_data.clear()
                        st.success(f"{delete_name} 직원 삭제 완료!")
                        st.rerun()

    # =====================================================
    # 월별 연차 통계
    # 관리자: 전체 / 직원: 본인만
    # =====================================================
    with st.expander("📅 월별 연차 통계", expanded=False):

        if is_admin:
            stat_base_df = df.copy()
        else:
            stat_base_df = df[df["이름"] == login_name].copy()
        stat_col1, stat_col2 = st.columns(2)
        with stat_col1:
            stat_year = st.number_input(
                "조회 연도",
                min_value=2020,
                max_value=2100,
                value=get_target_year(),
                step=1,
                key="vac_stat_year_unique"
            )

        with stat_col2:
            stat_month = st.selectbox(
                "조회 월",
                list(range(1, 13)),
                index=max(0, datetime.today().month - 1),
                key="vac_stat_month_unique"
            )

        monthly_df, monthly_count, monthly_amount = build_monthly_stats(
            stat_base_df, int(stat_year), int(stat_month)
        )

        metric_col1, metric_col2 = st.columns(2)
        metric_col1.metric("해당 월 사용 건수", monthly_count)
        metric_col2.metric("해당 월 총 사용일수", format_leave_number(monthly_amount))

        if not monthly_df.empty:
            st.dataframe(monthly_df, use_container_width=True)
        else:
            st.info("해당 월 사용 내역이 없습니다.")

    # =====================================================
    # 관리자 전용: 전체 연차 현황
    # =====================================================
    with st.expander("📋 전체 연차 현황", expanded=False):

        if is_admin:
            display_df = df.copy()
        else:
            display_df = df[df["이름"] == login_name].copy()

        for col in ["입사일", "기산시작일", "기산종료일"]:
            if col in display_df.columns:
                display_df[col] = display_df[col].apply(format_display_date)

        for col in USE_COLS:
            if col in display_df.columns:
                display_df[col] = display_df[col].apply(
                    lambda x: (
                        format_display_date(x)
                        if format_display_date(x) != ""
                        else clean_text(x)
                    )
                )

        for col in ["발생 연차", "사용 연차", "잔여 연차"]:
            if col in display_df.columns:
                display_df[col] = display_df[col].apply(format_leave_number)

        basic_cols = [
            "이름", "입사일", "기산시작일", "기산종료일",
            "근속년수", "발생 연차", "사용 연차", "잔여 연차"
        ]

        use_cols = USE_COLS

        basic_cols = [col for col in basic_cols if col in display_df.columns]
        use_cols = [col for col in use_cols if col in display_df.columns]

        st.subheader("📊 기본 연차 정보")

        basic_df = display_df[basic_cols].copy()

        if "잔여 연차" in basic_df.columns:
            styled_basic_df = basic_df.style.map(
                style_remaining_leave,
                subset=["잔여 연차"]
            )
            st.dataframe(styled_basic_df, use_container_width=True, height=400)
        else:
            st.dataframe(basic_df, use_container_width=True, height=400)

        st.subheader("📅 연차 사용 이력")

        if use_cols:
            use_df = display_df[["이름"] + use_cols].set_index("이름")
            st.dataframe(use_df, use_container_width=False, height=600)
        else:
            st.info("사용일 컬럼이 없습니다.")

    # =====================================================
    # 엑셀 다운로드
    # =====================================================
    with st.expander("⬇️ 엑셀 다운로드", expanded=False):
        download_df = load_df("연차관리")
        excel_data = to_excel_bytes({"연차관리": download_df})

        st.download_button(
            label="엑셀 다운로드",
            data=excel_data,
            file_name=f"연차관리_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="vac_download_btn_unique"
        )           

# =========================================================
# 5. 실사 관리 시스템
# =========================================================
INSPECTION_SHEET_NAME = "실사관리"
INSPECTION_COLUMNS = [
    "요청일",
    "운영사",
    "현장명",
    "현장주소",
    "현장연락처",
    "주차면수",
    "상품구분",
    "환경부",
    "자투",
    "신규설치수량",
    "기설치수량",
    "영업담당자",
    "영업담당연락처",
    "요청내용",
    "비고",
    "첨부파일명",
    "첨부파일링크",
    "실사담당자",
    "실사예정일",
    "실사완료일",
    "진행상태",
    "실사결과",
    "특이사항",
    "후속조치",
    "계약여부",
    "계약일",
    "계약수량",
    "계약금액",
    "미계약사유"
]

INSPECTION_STATUS_OPTIONS = [
    "요청접수",
    "담당자배정",
    "일정확정",
    "실사진행",
    "실사완료",
    "계약완료",
    "미계약종결"
]

PRODUCT_OPTIONS = ["아이센서", "전기차충전기", "이전설치"]
ENV_OPTIONS = ["", "대상", "비대상"]
JATU_OPTIONS = ["", "있음", "없음"]
CONTRACT_OPTIONS = ["대기", "계약", "미계약"]


def get_inspection_sheet():
    client = get_gsheet_client()
    spreadsheet = client.open(INSPECTION_SHEET_NAME)
    worksheet = spreadsheet.worksheet("실사복구")
    return worksheet


def safe_int(value, default=0):
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return default
    return int(num)


def show_inspection_flash():
    msg = st.session_state.pop("inspection_flash", "")
    msg_type = st.session_state.pop("inspection_flash_type", "success")


    if msg:
        if msg_type == "success":
            st.success(msg)
        elif msg_type == "warning":
            st.warning(msg)
        elif msg_type == "error":
            st.error(msg)
        else:
            st.info(msg)


def set_inspection_flash(msg, msg_type="success"):
    st.session_state["inspection_flash"] = msg
    st.session_state["inspection_flash_type"] = msg_type


def normalize_inspection_df(df):
    # 컬럼 없으면 자동 생성
    for col in INSPECTION_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    # 컬럼 순서 맞추기
    df = df.reindex(columns=INSPECTION_COLUMNS, fill_value="").copy()

    # 숫자 컬럼 (없는 경우도 대비)
    int_cols = ["주차면수", "신규설치수량", "기설치수량", "계약수량"]
    for col in int_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    # 금액
    if "계약금액" in df.columns:
        df["계약금액"] = pd.to_numeric(df["계약금액"], errors="coerce").fillna(0)

    # 날짜
    date_cols = ["요청일", "실사예정일", "실사완료일", "계약일"]
    for col in date_cols:
        if col in df.columns:
            df[col] = df[col].astype(str)

    # 상태값
    if "진행상태" in df.columns:
        df["진행상태"] = df["진행상태"].astype(str).replace("", "요청접수")

    if "계약여부" in df.columns:
        df["계약여부"] = df["계약여부"].astype(str).replace("", "대기")

    return df

def detect_inspection_duplicates(df):
    keys = ["요청일", "상품구분", "현장명", "현장주소", "영업담당자"]
    existing_keys = [c for c in keys if c in df.columns]

    if not existing_keys:
        return pd.DataFrame()

    dup_df = df[df.duplicated(subset=existing_keys, keep=False)].copy()
    return dup_df

def find_original_inspection_index(full_df, target_row):
    full_df = full_df[INSPECTION_COLUMNS].copy().fillna("")

    for col in INSPECTION_COLUMNS:
        full_df[col] = full_df[col].astype(str).str.strip()

    mask = (
        (full_df["요청일"] == str(target_row["요청일"]).strip()) &
        (full_df["상품구분"] == str(target_row["상품구분"]).strip()) &
        (full_df["현장명"] == str(target_row["현장명"]).strip()) &
        (full_df["현장주소"] == str(target_row["현장주소"]).strip()) &
        (full_df["영업담당자"] == str(target_row["영업담당자"]).strip())
    )

    matched = full_df[mask]

    if matched.empty:
        return None

    return matched.index[0]

# =========================================================
# 차량관리
# =========================================================

VEHICLE_SHEET_NAME = "차량관리"
VEHICLE_REPAIR_SHEET_NAME = "차량정비이력"

VEHICLE_COLUMNS = [
    "차량명", "소유자", "소유형태", "유종", "차종", "모델명", "연식", "차량번호",
    "보험회사", "보험종류", "보험기간", "보험금액", "차량상태", "비고"
]

REPAIR_COLUMNS = [
    "차량번호", "수리일자", "수리내역", "금액", "비고"
]
VEHICLE_SPREADSHEET_ID = "1OBE54H30v_bQ1hxI7VMShtELMiIxdAUKnCXcpzwSlQQ"

def get_vehicle_spreadsheet():
    client = get_gsheet_client()
    return client.open_by_key(VEHICLE_SPREADSHEET_ID)

@st.cache_data(ttl=300)
def load_sheet_as_df(sheet_name, columns):
    try:
        sh = get_vehicle_spreadsheet() 
        ws = sh.worksheet(sheet_name)
        data = ws.get_all_records()

        df = pd.DataFrame(data)

        for col in columns:
            if col not in df.columns:
                df[col] = ""

        df = df[columns].copy()

        # None / NaN / nan 문자 방어
        df = df.where(pd.notnull(df), "")
        df = df.replace(["None", "none", "nan", "NaN", "NaT"], "")

        df = df.astype(str)

        return df

    except Exception as e:
        st.error(f"{sheet_name} 데이터를 불러오는 중 오류가 발생했습니다: {e}")
        return pd.DataFrame(columns=columns)


def save_df_to_sheet(sheet_name, df, columns):
    try:
        sh = get_vehicle_spreadsheet()
        ws = sh.worksheet(sheet_name)

        save_df = df.copy()

        for col in columns:
            if col not in save_df.columns:
                save_df[col] = ""

        save_df = save_df[columns].fillna("").astype(str)

        ws.clear()
        ws.update([columns] + save_df.values.tolist())

        return True

    except Exception as e:
        st.error(f"{sheet_name} 저장 중 오류가 발생했습니다: {e}")
        return False


def parse_money(value):
    try:
        if pd.isna(value):
            return 0
        text = str(value).replace(",", "").replace("₩", "").replace("￦", "").strip()
        if text == "":
            return 0
        return int(float(text))
    except:
        return 0


def parse_insurance_end_date(value):
    try:
        text = str(value).strip()
        if not text:
            return None

        if "~" in text:
            text = text.split("~")[-1].strip()

        text = text.replace(".", "-").replace("/", "-")
        dt = pd.to_datetime(text, errors="coerce")

        if pd.isna(dt):
            return None

        return dt.date()

    except:
        return None


def vehicle_page():
    st.title("🚗 차량관리")

    vehicle_df = load_sheet_as_df("차량관리", VEHICLE_COLUMNS)
    repair_df = load_sheet_as_df("차량정비이력", REPAIR_COLUMNS)

    # 숫자 계산용
    vehicle_temp = vehicle_df.copy()
    repair_temp = repair_df.copy()

    vehicle_temp["보험금액_숫자"] = vehicle_temp["보험금액"].apply(parse_money)
    repair_temp["금액_숫자"] = repair_temp["금액"].apply(parse_money)

    total_vehicle = len(vehicle_df)
    active_vehicle = len(vehicle_df[vehicle_df["차량상태"].astype(str).str.strip() != "매각"])
    sold_vehicle = len(vehicle_df[vehicle_df["차량상태"].astype(str).str.strip() == "매각"])
    total_insurance = int(vehicle_temp["보험금액_숫자"].sum())
    total_repair = int(repair_temp["금액_숫자"].sum())

    # 보험 만료 30일 이내 차량 수
    today = date.today()
    insurance_warning_count = 0
    insurance_expired_count = 0

    for _, row in vehicle_df.iterrows():
        end_date = parse_insurance_end_date(row.get("보험기간", ""))

        if end_date:
            remain_days = (end_date - today).days

            if remain_days < 0:
                insurance_expired_count += 1
            elif remain_days <= 30:
                insurance_warning_count += 1

    st.markdown("""
    <style>
    .vehicle-card {
        background: #ffffff;
        border-radius: 14px;
        padding: 18px 22px;
        border: 1px solid #eef2f7;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
        transition: all 0.22s ease;
        min-height: 105px;
    }
    .vehicle-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 10px 24px rgba(15, 23, 42, 0.10);
    }
    .vehicle-card-label {
        font-size: 13x;
        color: #64748b;
        font-weight: 700;
        margin-bottom: 10px;
    }
    .vehicle-card-value {
        font-size: 22px;
        font-weight: 700;
        color: #0f172a;
        line-height: 1.2;
        white-space: nowrap;
    }
    </style>
    """, unsafe_allow_html=True)

    def vehicle_card(title, value, sub=""):
        st.markdown(f"""
        <div class="vehicle-card">
            <div class="vehicle-card-label">{title}</div>
            <div class="vehicle-card-value">{value}</div>
            <div class="vehicle-card-sub">{sub}</div>
        </div>
        """, unsafe_allow_html=True)

    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)

    with col1:
        vehicle_card("전체 차량", total_vehicle, "등록 차량")

    with col2:
        vehicle_card("운행 차량", active_vehicle, "현재 운행")

    with col3:
        vehicle_card("매각 차량", sold_vehicle, "매각 처리")

    with col4:
        vehicle_card("보험금 합계", f"{total_insurance:,}원", "전체 보험료")

    with col5:
        vehicle_card("정비비 합계", f"{total_repair:,}원", "누적 정비비")

    with col6:
        vehicle_card("보험 만료임박", insurance_warning_count, "30일 이내")

    with col7:
        vehicle_card("보험 만료", insurance_expired_count, "기간 경과")

    if insurance_expired_count > 0:
        st.error(f"🚨 보험이 만료된 차량이 {insurance_expired_count}대 있습니다.")

    if insurance_warning_count > 0:
        st.warning(f"⚠️ 30일 이내 보험 만료 예정 차량이 {insurance_warning_count}대 있습니다.")

    st.divider()

    tab1, tab2, tab3 = st.tabs(["차량 목록", "정비 이력", "보험 만료 경고"])

    # =====================================================
    # 1. 차량 목록
    # =====================================================
    with tab1:
        st.subheader("차량 목록")

        edited_vehicle_df = st.data_editor(
            vehicle_df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="vehicle_editor_main_v3",
        )

        if st.button("💾 차량관리 저장", use_container_width=True):
            if save_df_to_sheet(VEHICLE_SHEET_NAME, edited_vehicle_df, VEHICLE_COLUMNS):
                st.cache_data.clear()
                st.success("차량관리 저장 완료!")
                st.rerun()

    # =====================================================
    # 2. 정비 이력
    # =====================================================
    with tab2:
        st.subheader("정비 이력")

        vehicle_numbers = (
            vehicle_df["차량번호"]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )

        with st.expander("➕ 정비이력 등록", expanded=False):

            with st.form("repair_add_form"):
                r1, r2, r3 = st.columns(3)

                add_vehicle_no = r1.selectbox(
                    "차량번호",
                    vehicle_numbers,
                    key="repair_add_vehicle_no"
                )

                add_repair_date = r2.date_input(
                    "수리일자",
                    value=date.today(),
                    key="repair_add_date"
                )

                add_amount = r3.text_input(
                    "금액",
                    placeholder="예: 120000 또는 무상교체",
                    key="repair_add_amount"
                )

                add_repair_content = st.text_input("수리내역")
                add_note = st.text_input("비고")

                submitted = st.form_submit_button("정비이력 등록")

                # ✅ 반드시 여기 안에!
                if submitted:
                    new_row = pd.DataFrame([{
                        "차량번호": str(add_vehicle_no).strip(),
                        "수리일자": str(add_repair_date),
                        "수리내역": add_repair_content,
                        "금액": add_amount,
                        "비고": add_note
                    }])

                    save_df = pd.concat([repair_df, new_row], ignore_index=True)

                    if save_df_to_sheet(VEHICLE_REPAIR_SHEET_NAME, save_df, REPAIR_COLUMNS):
                        st.cache_data.clear()
                        st.success("정비이력 등록 완료!")
                        st.rerun()

        selected_vehicle = st.selectbox(
            "차량번호 선택",
            ["전체"] + vehicle_numbers
        )

        view_repair_df = repair_df.copy()

        if selected_vehicle != "전체":
            view_repair_df = view_repair_df[
                view_repair_df["차량번호"].astype(str).str.strip() == selected_vehicle
            ]

        with st.expander("📋 정비이력 조회", expanded=False):

            edited_repair_df = st.data_editor(
                view_repair_df,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                key="repair_editor_main_v4"
            )

            if selected_vehicle != "전체":
                st.warning("특정 차량 조회 중에는 전체 정비이력 저장이 비활성화됩니다.")
            else:
                if st.button("💾 정비이력 저장", use_container_width=True):
                    if save_df_to_sheet(VEHICLE_REPAIR_SHEET_NAME, edited_repair_df, REPAIR_COLUMNS):
                        st.cache_data.clear()
                        st.success("정비이력 저장 완료!")
                        st.rerun()

        view_repair_temp = view_repair_df.copy()
        view_repair_temp["금액_숫자"] = view_repair_temp["금액"].apply(parse_money)
        selected_total_repair = int(view_repair_temp["금액_숫자"].sum())

        st.info(f"선택 차량 정비비 합계: {selected_total_repair:,}원")

        repair_edit_df = repair_df.copy()
        repair_edit_df = repair_edit_df.where(pd.notnull(repair_edit_df), "")
        repair_edit_df = repair_edit_df.replace(["None", "none", "nan", "NaN", "NaT"], "")

        for col in REPAIR_COLUMNS:
            if col not in repair_edit_df.columns:
                repair_edit_df[col] = ""

        with st.expander("⚠️ 정비이력 삭제", expanded=False):

            if view_repair_df.empty:
                st.info("삭제할 정비이력이 없습니다.")
            else:
                delete_options = []

                for idx, row in view_repair_df.iterrows():
                    delete_options.append(
                        f"{idx} | {row.get('차량번호', '')} | {row.get('수리일자', '')} | {row.get('수리내역', '')} | {row.get('금액', '')}"
                    )

                selected_delete = st.selectbox(
                    "삭제할 정비이력 선택",
                    delete_options,
                    key="repair_delete_select"
                )

                confirm_delete = st.checkbox(
                    "정말 삭제합니다. 되돌리기 어렵습니다.",
                    key="repair_delete_confirm"
                )

                if st.button("🗑️ 선택 정비이력 삭제", use_container_width=True, key="repair_delete_btn"):
                    if not confirm_delete:
                        st.warning("삭제 확인 체크를 먼저 해주세요.")
                    else:
                        delete_idx = int(selected_delete.split("|")[0].strip())

                        save_df = repair_df[REPAIR_COLUMNS].copy()
                        save_df = save_df.drop(index=delete_idx).reset_index(drop=True)

                        if save_df_to_sheet(VEHICLE_REPAIR_SHEET_NAME, save_df, REPAIR_COLUMNS):
                            st.cache_data.clear()
                            st.success("정비이력이 삭제되었습니다.")
                            st.rerun()

        with st.expander("📊 차량별 정비비 합계", expanded=False):

            if not view_repair_df.empty:

                repair_summary = view_repair_df.copy()
                repair_summary["금액"] = repair_summary["금액"].apply(parse_money)

                summary_df = (
                    repair_summary
                    .groupby("차량번호", as_index=False)["금액"]
                    .sum()
                    .sort_values("금액", ascending=False)
                )

                summary_df["금액"] = summary_df["금액"].apply(lambda x: f"{int(x):,}원")

                st.dataframe(summary_df, use_container_width=True, hide_index=True)

            else:
                st.info("선택한 차량의 정비이력이 없습니다.")


    # =====================================================
    # 3. 보험 만료 경고
    # =====================================================
    with tab3:
        st.subheader("보험 만료 경고")

        today = date.today()
        warning_rows = []

        for _, row in vehicle_df.iterrows():
            end_date = parse_insurance_end_date(row.get("보험기간", ""))

            if end_date:
                remain_days = (end_date - today).days

                if 0 <= remain_days <= 30:
                    warning_rows.append({
                        "차량명": row.get("차량명", ""),
                        "차량번호": row.get("차량번호", ""),
                        "보험회사": row.get("보험회사", ""),
                        "보험기간": row.get("보험기간", ""),
                        "남은일수": remain_days,
                        "비고": row.get("비고", "")
                    })

        if warning_rows:
            st.warning(f"보험 만료 30일 이내 차량이 {len(warning_rows)}대 있습니다.")
            st.dataframe(pd.DataFrame(warning_rows), use_container_width=True, hide_index=True)
        else:
            st.success("보험 만료 임박 차량이 없습니다.")

@st.cache_data(ttl=60)
def load_inspection_data():
    try:
        sheet = get_inspection_sheet()
        values = sheet.get_all_values()

        if not values:
            return pd.DataFrame(columns=INSPECTION_COLUMNS)

        header = values[0]
        data_rows = values[1:]

        # 👉 기존 데이터 그대로 로드
        temp_df = pd.DataFrame(data_rows, columns=header)

        # 👉 새로운 DF 생성
        df = pd.DataFrame(columns=INSPECTION_COLUMNS)

        # 👉 컬럼 매핑 (안전 버전)
        column_map = {
            "요청일": "요청일",
            "운영사": "운영사",
            "현장명": "현장명",
            "단지명": "현장명",
            "현장주소": "현장주소",
            "주소": "현장주소",
            "현장연락처": "현장연락처",
            "전화번호": "현장연락처",
            "주차면수": "주차면수",
            "상품구분": "상품구분",
            "환경부": "환경부",
            "자투": "자투",
            "신규설치수량": "신규설치수량",
            "수량": "신규설치수량",
            "기설치수량": "기설치수량",
            "영업담당자": "영업담당자",
            "영업담당연락처": "영업담당연락처",
            "요청내용": "요청내용",
            "비고": "비고",
            "첨부파일명": "첨부파일명",
            "첨부파일링크": "첨부파일링크",
            "실사담당자": "실사담당자",
            "실사예정일": "실사예정일",
            "실사완료일": "실사완료일",
            "진행상태": "진행상태",
            "실사결과": "실사결과",
            "특이사항": "특이사항",
            "후속조치": "후속조치",
            "계약여부": "계약여부",
            "계약일": "계약일",
            "계약수량": "계약수량",
            "계약금액": "계약금액",
            "미계약사유": "미계약사유",
        }

        # 👉 안전하게 매핑
        for old_col, new_col in column_map.items():
            if old_col in temp_df.columns:
                df[new_col] = temp_df[old_col]

        # 👉 없는 컬럼 채우기
        for col in INSPECTION_COLUMNS:
            if col not in df.columns:
                df[col] = ""

        df = normalize_inspection_df(df)
        return df

    except Exception as e:
        st.error(f"실사 데이터를 불러오지 못했습니다: {e}")
        return pd.DataFrame(columns=INSPECTION_COLUMNS)

def apply_product_filter(df):
    if df is None or df.empty:
        return df

    if "상품구분" not in df.columns:
        return df

    if st.session_state.business == "아이센서":
        return df[df["상품구분"].astype(str).str.strip().isin(["아이센서"])].copy()
    else:
        return df[df["상품구분"].astype(str).str.strip().isin(["전기차충전기", "이전설치"])].copy()        


def save_inspection_data(df, sheet=None):
    if sheet is None:
        sheet = get_inspection_sheet()

    save_df = normalize_inspection_df(df).copy()

    # 임시 컬럼 제거
    if "row_id" in save_df.columns:
        save_df = save_df.drop(columns=["row_id"])

    # 빈 데이터 저장 방지
    if save_df.empty:
        raise Exception("실사 데이터가 비어 있어 저장을 중단했습니다.")

    def clean_cell(x):
        if pd.isna(x):
            return ""
        if isinstance(x, (pd.Timestamp, datetime, date)):
            return x.strftime("%Y-%m-%d")
        text = str(x).strip()
        if text.lower() in ["none", "nan", "nat"]:
            return ""
        return text

    for col in save_df.columns:
        save_df[col] = save_df[col].apply(clean_cell)

    rows = [save_df.columns.tolist()] + save_df.values.tolist()

    if len(rows) <= 1:
        raise Exception("실사 데이터 행이 없어 저장을 중단했습니다.")

    # 절대 먼저 clear 하지 않음
    sheet.update("A1", rows, value_input_option="USER_ENTERED")

    st.cache_data.clear()   


def render_inspection_common_style():
    st.markdown("""
    <style>

    /* 전체 배경 */
    .main {
        background-color: #f8fafc;
    }

    /* 제목 */
    .erp-page-title {
        font-size: 25px !important;
        font-weight: 700 !important;
        color: #0f172a !important;
        margin-bottom: 4px !important;
    }

    .erp-page-desc {
        font-size: 13px !important;
        color: #64748b !important;
        margin-bottom: 20px !important;
    }

    /* 카드 */
    div[data-testid="stMetric"] {
        background: white;
        border-radius: 12px;
        padding: 15px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
    }

    .erp-summary-card {
        transition: all 0.2s ease !important;
    }

    .erp-summary-card:hover {
        transform: translateY(-4px) scale(1.01);
        box-shadow: 0 10px 20px rgba(0,0,0,0.12) !important;
    }      
                      
    /* KPI */
    div[data-testid="stMetricValue"] {
        font-size: 22px !important;
        font-weight: 700 !important;
    }

    div[data-testid="stMetricLabel"] {
        font-size: 12px !important;
        color: #64748b !important;
    }

    /* 버튼 */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
    }

    /* expander */
    details {
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 8px;
        background: white;
    }

    /* 추가 UI */
    .erp-summary-card {
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        background: #ffffff;
        padding: 12px 14px;
        min-height: 88px;
    }

    .erp-summary-label {
        font-size: 12px;
        color: #64748b;
        margin-bottom: 10px;
    }

    .erp-summary-value {
        font-size: 22px;
        font-weight: 700;
        color: #0f172a;
    }

    div[data-testid="stExpander"] {
        border: 1px solid #e5e7eb !important;
        border-radius: 12px !important;
        background: #ffffff !important;
    }

    div[data-testid="stExpander"] details summary p {
        font-size: 16px !important;
        font-weight: 700 !important;
        color: #0f172a !important;
    }

    div[data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;
    }

    div[data-testid="stForm"] {
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        background: #ffffff;
        padding: 14px;
    }

    button[kind="primary"],
    button[kind="secondary"] {
        border-radius: 10px !important;
    }
    /* 버튼 ERP 스타일 */
    .stButton > button {
        border-radius: 10px !important;
        border: 1px solid #cbd5e1 !important;
        background: #ffffff !important;
        color: #0f172a !important;
        font-weight: 700 !important;
        padding: 0.45rem 0.9rem !important;
    }

    .stButton > button:hover {
        border-color: #2563eb !important;
        color: #2563eb !important;
        background: #eff6ff !important;
    }

    /* 입력창 / 선택창 */
    div[data-baseweb="select"] > div,
    input,
    textarea {
        border-radius: 10px !important;
    }

    /* 테이블 헤더 느낌 */
    div[data-testid="stDataFrame"] {
        border: 1px solid #e2e8f0 !important;
        border-radius: 12px !important;
        overflow: hidden !important;
        background: #ffffff !important;
    }

    /* Expander 간격 */
    div[data-testid="stExpander"] {
        margin-bottom: 16px !important;
    }

    /* 구분선 여백 */
    hr {
        margin-top: 22px !important;
        margin-bottom: 22px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# =========================
# 실사관리 페이지
# =========================

def inspection_page():
    render_inspection_common_style()
    render_common_style()

    st.markdown('<div class="erp-page-title">🔎 실사 관리 프로그램</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="erp-page-desc">실사 요청 등록 → 담당자 배정 → 일정 입력 → 결과 작성 → 계약 여부 관리</div>',
        unsafe_allow_html=True
    )

    show_inspection_flash()

    df = load_inspection_data()
    df = normalize_inspection_df(df)

    # ✅ 사업 선택에 따라 실사 데이터 분리
    df = apply_product_filter(df)
    # 🚨 중복 감지
    dup_df = detect_inspection_duplicates(df)

    if not dup_df.empty:
        st.warning(f"⚠️ 중복 데이터 {len(dup_df)}건 발견되었습니다. (구글시트 정리 필요)")

        with st.expander("중복 데이터 확인", expanded=False):
            show_cols = ["요청일", "상품구분", "현장명", "현장주소", "영업담당자"]
        show_cols = [c for c in show_cols if c in dup_df.columns]

        st.dataframe(
            dup_df[show_cols].sort_values(show_cols),
            use_container_width=True,
            hide_index=True
        )
    dedup_keys = ["요청일", "상품구분", "현장명", "현장주소", "영업담당자"]
    existing_keys = [c for c in dedup_keys if c in df.columns]

    if existing_keys:
        df = df.drop_duplicates(subset=existing_keys, keep="first").copy()

    df = df.reset_index(drop=True)
    df["row_id"] = df.index

    total_count = len(df)
    pending_count = len(df[df["진행상태"] == "요청접수"])
    assigned_count = len(df[df["진행상태"].isin(["담당자배정", "일정확정", "실사진행"])])
    done_count = len(df[df["진행상태"] == "실사완료"])
    contract_done_count = len(df[df["계약여부"] == "계약"])

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.markdown(
            f"""
            <div class="erp-summary-card">
                <div class="erp-summary-label">전체 요청</div>
                <div class="erp-summary-value">{total_count}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:
        st.markdown(
            f"""
            <div class="erp-summary-card">
                <div class="erp-summary-label">요청접수</div>
                <div class="erp-summary-value">{pending_count}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c3:
        st.markdown(
            f"""
            <div class="erp-summary-card">
                <div class="erp-summary-label">진행중</div>
                <div class="erp-summary-value">{assigned_count}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c4:
        st.markdown(
            f"""
            <div class="erp-summary-card">
                <div class="erp-summary-label">실사완료</div>
                <div class="erp-summary-value">{done_count}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c5:
        st.markdown(
            f"""
            <div class="erp-summary-card">
                <div class="erp-summary-label">계약완료</div>
                <div class="erp-summary-value">{contract_done_count}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.divider()

    with st.expander("📝 1. 실사 요청 등록", expanded=False):
        st.markdown('<div class="erp-section-title">실사 요청 등록</div>', unsafe_allow_html=True)
        form_ver = st.session_state.inspection_form_version

        with st.form(f"inspection_request_form_new_{form_ver}"):
            c1, c2, c3 = st.columns(3)
            req_date = c1.date_input("요청일", value=date.today(), key=f"req_date_{form_ver}")
            operator_name = c2.text_input("운영사", key=f"operator_name_{form_ver}")
            site_name = c3.text_input("현장명", key=f"site_name_{form_ver}")

            c4, c5, c6 = st.columns(3)
            site_address = c4.text_input("현장주소", key=f"site_address_{form_ver}")
            site_phone = c5.text_input("현장연락처", key=f"site_phone_{form_ver}")
            product_type = c6.selectbox("상품구분", PRODUCT_OPTIONS, key=f"product_type_{form_ver}")
            env_gov = st.selectbox("환경부", ["", "대상", "비대상"], key=f"env_{form_ver}")
            jatu = st.selectbox("자투", ["", "있음", "없음"], key=f"jatu_{form_ver}")

            c7, c8, c9 = st.columns(3)
            parking_count = c7.number_input("주차면수", min_value=0, step=1, value=0, key=f"parking_count_{form_ver}")
            new_qty = c8.number_input("신규설치수량", min_value=0, step=1, value=0, key=f"new_qty_{form_ver}")
            installed_qty = c9.number_input("기설치수량", min_value=0, step=1, value=0, key=f"installed_qty_{form_ver}")

            c10, c11 = st.columns(2)
            sales_manager = c10.text_input("영업담당자", key=f"sales_manager_{form_ver}")
            sales_phone = c11.text_input("영업담당자 연락처", key=f"sales_phone_{form_ver}")

            request_content = st.text_area("요청내용", key=f"request_content_{form_ver}")
            note = st.text_input("비고", key=f"note_{form_ver}")

            st.subheader("첨부파일")
            uploaded_file = st.file_uploader(
                "실사 관련 파일 업로드",
                type=["pdf", "png", "jpg", "jpeg", "xlsx", "xls", "doc", "docx"],
                key=f"insp_uploaded_file_new_{form_ver}"
            )

            submit_request = st.form_submit_button("실사 요청 등록")

            if submit_request:
                if not site_name.strip():
                    st.warning("현장명을 입력해주세요.")
                elif not sales_manager.strip():
                    st.warning("영업담당자를 입력해주세요.")
                else:
                    attachment_name = ""
                    attachment_link = ""

                    if uploaded_file is not None:
                        try:
                            attachment_name, attachment_link = upload_file_to_drive(
                                uploaded_file,
                                folder_id="13W2N1v9IBiuZEstmTrvt57Zg8XQiHt7J"
                            )
                        except Exception as e:
                            st.error(str(e))
                            attachment_name = ""
                            attachment_link = ""

                    new_row = pd.DataFrame([{
                        "요청일": str(req_date),
                        "운영사": operator_name.strip(),
                        "현장명": site_name.strip(),
                        "현장주소": site_address.strip(),
                        "현장연락처": site_phone.strip(),
                        "주차면수": int(parking_count),
                        "상품구분": product_type,
                        "환경부": env_gov,
                        "자투": jatu,
                        "신규설치수량": int(new_qty),
                        "기설치수량": int(installed_qty),
                        "영업담당자": sales_manager.strip(),
                        "영업담당연락처": sales_phone.strip(),
                        "요청내용": request_content.strip(),
                        "비고": note.strip(),
                        "첨부파일명": attachment_name,
                        "첨부파일링크": attachment_link,
                        "실사담당자": "",
                        "실사예정일": "",
                        "실사완료일": "",
                        "진행상태": "요청접수",
                        "실사결과": "",
                        "특이사항": "",
                        "후속조치": "",
                        "계약여부": "대기",
                        "계약일": "",
                        "계약수량": 0,
                        "계약금액": 0,
                        "미계약사유": ""
                    }])

                    save_df = df[INSPECTION_COLUMNS].copy() if not df.empty else pd.DataFrame(columns=INSPECTION_COLUMNS)
                    save_df = pd.concat([save_df, new_row], ignore_index=True)
                    save_inspection_data(full_df)

                    set_inspection_flash("실사 요청이 등록되었습니다.", "success")
                    st.session_state.inspection_form_version += 1
                    st.rerun()

    st.divider()

    with st.expander("📋 2. 전체 실사 현황", expanded=False):
        st.markdown('<div class="erp-section-title">전체 실사 현황</div>', unsafe_allow_html=True)
        st.markdown('<div class="erp-soft-box">전체 실사 데이터를 확인합니다.</div>', unsafe_allow_html=True)

        status_list = ["전체"] + INSPECTION_STATUS_OPTIONS
        product_list = ["전체"] + PRODUCT_OPTIONS
        contract_list = ["전체"] + CONTRACT_OPTIONS

        f1, f2, f3, f4 = st.columns(4)
        status_filter = f1.selectbox("진행상태", status_list, key="insp_filter_status_new")
        product_filter = f2.selectbox("상품구분", product_list, key="insp_filter_product_new")
        contract_filter = f3.selectbox("계약여부", contract_list, key="insp_filter_contract_new")
        keyword = f4.text_input("검색", placeholder="현장명 / 주소 / 담당자 / 운영사", key="insp_filter_keyword_new")

        filtered_df = df.copy()

        if status_filter != "전체":
            filtered_df = filtered_df[filtered_df["진행상태"] == status_filter]

        if product_filter != "전체":
            filtered_df = filtered_df[filtered_df["상품구분"] == product_filter]

        if contract_filter != "전체":
            filtered_df = filtered_df[filtered_df["계약여부"] == contract_filter]

        if keyword.strip():
            kw = keyword.strip()
            filtered_df = filtered_df[
                filtered_df["현장명"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["현장주소"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["영업담당자"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["실사담당자"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["운영사"].astype(str).str.contains(kw, case=False, na=False)
            ].copy()

        show_df = filtered_df[[
            "요청일",
            "상품구분",
            "환경부",
            "자투",
            "현장명",
            "현장주소",
            "현장연락처",
            "운영사",
            "신규설치수량",
            "기설치수량",
            "주차면수",

            "영업담당자",
            "영업담당연락처",

            "실사담당자",
            "실사예정일",
            "진행상태",
            "계약여부",

            "첨부파일링크"
        ]].copy()

        def status_style(val):
            if str(val) == "요청접수":
                return "background-color: #fef3c7; color: #92400e; font-weight: 600;"
            elif str(val) in ["담당자배정", "일정확정", "실사진행"]:
                return "background-color: #dbeafe; color: #1d4ed8; font-weight: 600;"
            elif str(val) == "실사완료":
                return "background-color: #dcfce7; color: #166534; font-weight: 600;"
            elif str(val) == "계약완료":
                return "background-color: #dcfce7; color: #166534; font-weight: 700;"
            elif str(val) == "미계약종결":
                return "background-color: #fee2e2; color: #991b1b; font-weight: 600;"
            return ""

        def contract_style(val):
            if str(val) == "계약":
                return "background-color: #dcfce7; color: #166534; font-weight: 700;"
            elif str(val) == "미계약":
                return "background-color: #fee2e2; color: #991b1b; font-weight: 700;"
            elif str(val) == "대기":
                return "background-color: #f3f4f6; color: #374151; font-weight: 600;"
            return ""

        if show_df.empty:
            st.info("조건에 맞는 실사 내역이 없습니다.")
        else:
            show_df["첨부파일열기"] = show_df["첨부파일링크"]

            for col in show_df.columns:
                show_df[col] = show_df[col].apply(
                    lambda x: "" if pd.isna(x) or str(x).strip().lower() in ["none", "nan", "nat"] else x
                )
            styled_df = show_df.style.map(status_style, subset=["진행상태"]).map(contract_style, subset=["계약여부"])

            st.dataframe(
                styled_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "현장명": st.column_config.TextColumn("현장명", width="medium"),
                    "현장주소": st.column_config.TextColumn("현장주소", width="medium"),
                    "운영사": st.column_config.TextColumn("운영사", width="small"),
                    "신규설치수량": st.column_config.NumberColumn("신규설치수량", width="small"),
                    "기설치수량": st.column_config.NumberColumn("기설치수량", width="small"),
                    "주차면수": st.column_config.NumberColumn("주차면수", width="small"),
                    "영업담당자": st.column_config.TextColumn("영업담당자", width="small"),
                    "영업담당연락처": st.column_config.TextColumn("영업담당연락처", width="medium"),
                    "실사담당자": st.column_config.TextColumn("실사담당자", width="small"),
                    "실사예정일": st.column_config.TextColumn("실사예정일", width="small"),
                    "진행상태": st.column_config.TextColumn("진행상태", width="small"),
                    "계약여부": st.column_config.TextColumn("계약여부", width="small"),
                    "첨부파일링크": None,
                    "첨부파일열기": st.column_config.LinkColumn(
                        "첨부파일 열기",
                        display_text="열기",
                        width="small"
                    )
                }
            )

            st.caption("상태와 계약여부는 색상으로 구분되며, 첨부파일은 '열기'로 바로 확인할 수 있습니다.")

    st.divider()

    with st.expander("🧑‍🔧 3. 담당자 배정 / 일정 입력", expanded=False):
        if df.empty:
            st.info("배정할 실사 요청이 없습니다.")
        else:
            assign_options = [
                f"{row['row_id']} | {row['요청일']} | {row['현장명']} | {row['상품구분']} | {row['영업담당자']}"
                for _, row in df.iterrows()
            ]

            selected_assign = st.selectbox(
                "배정할 실사 선택",
                assign_options,
                key="insp_assign_select_new"
            )
            assign_idx = int(selected_assign.split("|")[0].strip())
            assign_row = df.loc[df["row_id"] == assign_idx].iloc[0]

            assign_date_raw = str(assign_row["실사예정일"]).strip()
            parsed_assign_date = pd.to_datetime(assign_date_raw, errors="coerce")
            default_assign_date = parsed_assign_date.date() if pd.notna(parsed_assign_date) else date.today()

            with st.form(f"inspection_assign_form_{assign_idx}"):
                a1, a2, a3 = st.columns(3)

                inspector = a1.text_input(
                    "실사담당자",
                    value=str(assign_row["실사담당자"])
                )

                inspect_date = a2.date_input(
                    "실사예정일",
                    value=default_assign_date
                )

                current_status = str(assign_row["진행상태"]).strip()
                default_status_index = (
                    INSPECTION_STATUS_OPTIONS.index(current_status)
                    if current_status in INSPECTION_STATUS_OPTIONS
                    else 0
                )

                inspect_status = a3.selectbox(
                    "진행상태",
                    INSPECTION_STATUS_OPTIONS,
                    index=default_status_index
                )

                b1, b2 = st.columns(2)
                assign_submit = b1.form_submit_button("배정 / 일정 저장")
                delete_submit = b2.form_submit_button("담당자 배정 삭제")

                if assign_submit:
                    target_row = df.loc[df["row_id"] == assign_idx].iloc[0]

                    full_df = load_inspection_data()
                    full_df = full_df[INSPECTION_COLUMNS].copy()

                    original_idx = find_original_inspection_index(full_df, target_row)

                    if original_idx is None:
                        st.error("원본 데이터를 찾지 못했습니다.")
                        st.stop()

                    full_df.loc[original_idx, "실사담당자"] = inspector.strip()
                    full_df.loc[original_idx, "실사예정일"] = str(inspect_date)
                    full_df.loc[original_idx, "진행상태"] = inspect_status

                    save_inspection_data(full_df)
                    st.cache_data.clear()

                    set_inspection_flash("담당자 배정 및 일정 저장 완료!", "success")
                    st.rerun()

                if delete_submit:
                    target_row = df.loc[df["row_id"] == assign_idx].iloc[0]

                    full_df = load_inspection_data()
                    full_df = full_df[INSPECTION_COLUMNS].copy()

                    original_idx = find_original_inspection_index(full_df, target_row)

                    if original_idx is None:
                        st.error("원본 데이터를 찾지 못했습니다.")
                        st.stop()

                    full_df.loc[original_idx, "실사담당자"] = ""
                    full_df.loc[original_idx, "실사예정일"] = ""
                    full_df.loc[original_idx, "진행상태"] = "요청접수"

                    save_inspection_data(full_df)
                    st.cache_data.clear()

                    set_inspection_flash("담당자 배정이 삭제되었습니다.", "success")
                    st.rerun()

    st.divider()

    with st.expander("📝 4. 실사 결과 입력", expanded=False):
        if df.empty:
            st.info("입력할 실사 내역이 없습니다.")
        else:
            result_options = [
                f"{row['row_id']} | {row['현장명']} | {row['실사담당자']} | {row['진행상태']}"
                for _, row in df.iterrows()
            ]

            selected_result = st.selectbox("결과 입력 대상 선택", result_options, key="insp_result_select_new")
            result_idx = int(selected_result.split("|")[0].strip())
            result_row = df.loc[df["row_id"] == result_idx].iloc[0]

            complete_date_raw = str(result_row["실사완료일"]).strip()
            parsed_complete_date = pd.to_datetime(complete_date_raw, errors="coerce")
            default_complete_date = parsed_complete_date.date() if pd.notna(parsed_complete_date) else date.today()

            with st.form(f"inspection_result_form_{result_idx}"):
                r1, r2 = st.columns(2)
                result_text = r1.text_area("실사결과", value=str(result_row["실사결과"]))
                special_note = r2.text_area("특이사항", value=str(result_row["특이사항"]))

                r3, r4 = st.columns(2)
                follow_up = r3.text_area("후속조치", value=str(result_row["후속조치"]))
                complete_date = r4.date_input("실사완료일", value=default_complete_date)

                result_status = st.selectbox(
                    "진행상태",
                    INSPECTION_STATUS_OPTIONS,
                    index=INSPECTION_STATUS_OPTIONS.index(result_row["진행상태"]) if result_row["진행상태"] in INSPECTION_STATUS_OPTIONS else 0
                )

                result_submit = st.form_submit_button("실사 결과 저장")

                if result_submit:
                    target_row = df.loc[df["row_id"] == result_idx].iloc[0]

                    full_df = load_inspection_data()
                    full_df = full_df[INSPECTION_COLUMNS].copy()

                    original_idx = find_original_inspection_index(full_df, target_row)

                    if original_idx is None:
                        st.error("원본 데이터를 찾지 못했습니다.")
                        st.stop()

                    full_df.loc[original_idx, "실사결과"] = result_text.strip()
                    full_df.loc[original_idx, "특이사항"] = special_note.strip()
                    full_df.loc[original_idx, "후속조치"] = follow_up.strip()
                    full_df.loc[original_idx, "실사완료일"] = str(complete_date)
                    full_df.loc[original_idx, "진행상태"] = result_status

                    save_inspection_data(full_df)

                    set_inspection_flash("실사 결과 저장 완료!", "success")
                    st.rerun()

    st.divider()

    with st.expander("💰 5. 계약 여부 입력", expanded=False):
        if df.empty:
            st.info("계약 처리할 내역이 없습니다.")
        else:
            contract_options = [
                f"{row['row_id']} | {row['현장명']} | {row['상품구분']} | 현재:{row['계약여부']}"
                for _, row in df.iterrows()
            ]

            selected_contract = st.selectbox("계약 처리 대상 선택", contract_options, key="insp_contract_select_new")
            contract_idx = int(selected_contract.split("|")[0].strip())
            contract_row = df.loc[df["row_id"] == contract_idx].iloc[0]

            contract_date_raw = str(contract_row["계약일"]).strip()
            parsed_contract_date = pd.to_datetime(contract_date_raw, errors="coerce")
            default_contract_date = parsed_contract_date.date() if pd.notna(parsed_contract_date) else date.today()

            contract_qty_default = safe_int(contract_row["계약수량"], 0)
            contract_amount_default = safe_int(contract_row["계약금액"], 0)

            with st.form(f"inspection_contract_form_{contract_idx}"):
                ct1, ct2, ct3 = st.columns(3)
                contract_status = ct1.selectbox(
                    "계약여부",
                    CONTRACT_OPTIONS,
                    index=CONTRACT_OPTIONS.index(contract_row["계약여부"]) if contract_row["계약여부"] in CONTRACT_OPTIONS else 0
                )
                contract_date = ct2.date_input("계약일", value=default_contract_date)
                contract_qty = ct3.number_input("계약수량", min_value=0, step=1, value=contract_qty_default)

                ct4, ct5 = st.columns(2)
                contract_amount = ct4.number_input("계약금액", min_value=0, step=10000, value=contract_amount_default)
                fail_reason = ct5.text_input("미계약사유", value=str(contract_row["미계약사유"]))

                contract_submit = st.form_submit_button("계약 정보 저장")

                if contract_submit:
                    target_row = df.loc[df["row_id"] == contract_idx].iloc[0]

                    full_df = load_inspection_data()
                    full_df = full_df[INSPECTION_COLUMNS].copy()

                    original_idx = find_original_inspection_index(full_df, target_row)

                    if original_idx is None:
                        st.error("원본 데이터를 찾지 못했습니다.")
                        st.stop()

                    full_df.loc[original_idx, "계약여부"] = contract_status
                    full_df.loc[original_idx, "계약일"] = str(contract_date) if contract_status == "계약" else ""
                    full_df.loc[original_idx, "계약수량"] = int(contract_qty) if contract_status == "계약" else 0
                    full_df.loc[original_idx, "계약금액"] = int(contract_amount) if contract_status == "계약" else 0
                    full_df.loc[original_idx, "미계약사유"] = fail_reason.strip() if contract_status == "미계약" else ""

                    if contract_status == "계약":
                        full_df.loc[original_idx, "진행상태"] = "계약완료"
                    elif contract_status == "미계약":
                        full_df.loc[original_idx, "진행상태"] = "미계약종결"

                    save_inspection_data(full_df)

    st.divider()

    with st.expander("✏️ 6. 상세 보기 / 수정", expanded=False):
        if df.empty:
            st.info("조회할 내역이 없습니다.")
        else:
            view_options = [
                f"{row['row_id']} | {row['현장명']} | {row['상품구분']} | {row['진행상태']}"
                for _, row in df.iterrows()
            ]

            selected_view = st.selectbox("조회 대상 선택", view_options, key="insp_view_select_new")
            view_idx = int(selected_view.split("|")[0].strip())
            view_row = df.loc[df["row_id"] == view_idx].iloc[0]

            is_edit_mode = (
                st.session_state.get("inspection_edit_mode", False)
                and st.session_state.get("inspection_edit_target") == view_idx
            )

            if not is_edit_mode:
                st.markdown("""
                <style>
                .detail-title {
                    font-size: 22px;
                    font-weight: 800;
                    color: #0f172a;
                    margin-bottom: 14px;
                }
                .detail-summary-wrap {
                    display: grid;
                    grid-template-columns: repeat(4, 1fr);
                    gap: 12px;
                    margin-bottom: 18px;
                }
                .detail-summary-card {
                    border: 1px solid #e5e7eb;
                    border-radius: 14px;
                    background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
                    padding: 14px 16px;
                    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.06);
                }
                .detail-summary-label {
                    font-size: 12px;
                    color: #64748b;
                    margin-bottom: 8px;
                }
                .detail-summary-value {
                    font-size: 20px;
                    font-weight: 700;
                    color: #0f172a;
                    line-height: 1.25;
                }
                .detail-section {
                    margin-top: 10px;
                    margin-bottom: 14px;
                    padding: 14px;
                    border: 1px solid #e5e7eb;
                    border-radius: 14px;
                    background: #ffffff;
                    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.05);
                }
                .detail-section-title {
                    font-size: 16px;
                    font-weight: 800;
                    color: #0f172a;
                    margin-bottom: 12px;
                }
                .detail-label {
                    font-size: 13px;
                    font-weight: 700;
                    color: #334155;
                    margin-bottom: 6px;
                }
                .detail-box {
                    border: 1px solid #dbe3ee;
                    border-radius: 12px;
                    background: #f8fbff;
                    padding: 12px 14px;
                    min-height: 46px;
                    font-size: 14px;
                    color: #0f172a;
                    display: flex;
                    align-items: center;
                }
                .detail-textarea {
                    border: 1px solid #e5e7eb;
                    border-radius: 12px;
                    background: #f8fafc;
                    padding: 14px;
                    white-space: pre-wrap;
                    line-height: 1.7;
                    font-size: 14px;
                    color: #111827;
                    min-height: 72px;
                }
                </style>
                """, unsafe_allow_html=True)

                st.markdown('<div class="detail-title">📄 실사 상세보기</div>', unsafe_allow_html=True)

                st.markdown(
                    f"""
                    <div class="detail-summary-wrap">
                        <div class="detail-summary-card">
                            <div class="detail-summary-label">진행상태</div>
                            <div class="detail-summary-value">{str(view_row["진행상태"]).strip() or "-"}</div>
                        </div>
                        <div class="detail-summary-card">
                            <div class="detail-summary-label">계약여부</div>
                            <div class="detail-summary-value">{str(view_row["계약여부"]).strip() or "-"}</div>
                        </div>
                        <div class="detail-summary-card">
                            <div class="detail-summary-label">상품구분</div>
                            <div class="detail-summary-value">{str(view_row["상품구분"]).strip() or "-"}</div>
                        </div>
                        <div class="detail-summary-card">
                            <div class="detail-summary-label">영업담당자</div>
                            <div class="detail-summary-value">{str(view_row["영업담당자"]).strip() or "-"}</div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                # 1. 기본 정보
                st.markdown('<div class="detail-section">', unsafe_allow_html=True)
                st.markdown('<div class="detail-section-title">기본 정보</div>', unsafe_allow_html=True)

                row1_col1, row1_col2, row1_col3 = st.columns(3)
                with row1_col1:
                    st.markdown('<div class="detail-label">요청일</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["요청일"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row1_col2:
                    st.markdown('<div class="detail-label">운영사</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["운영사"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row1_col3:
                    st.markdown('<div class="detail-label">현장명</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["현장명"]).strip() or "-"}</div>', unsafe_allow_html=True)

                row2_col1, row2_col2, row2_col3 = st.columns(3)
                with row2_col1:
                    st.markdown('<div class="detail-label">현장주소</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["현장주소"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row2_col2:
                    st.markdown('<div class="detail-label">현장연락처</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["현장연락처"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row2_col3:
                    st.markdown('<div class="detail-label">상품구분</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["상품구분"]).strip() or "-"}</div>', unsafe_allow_html=True)

                row3_col1, row3_col2, row3_col3 = st.columns(3)
                with row3_col1:
                    st.markdown('<div class="detail-label">환경부</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["환경부"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row3_col2:
                    st.markdown('<div class="detail-label">자투</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["자투"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row3_col3:
                    st.markdown('<div class="detail-label">주차면수</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["주차면수"]).strip() or "0"}</div>', unsafe_allow_html=True)

                row4_col1, row4_col2, row4_col3 = st.columns(3)
                with row4_col1:
                    st.markdown('<div class="detail-label">신규설치수량</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["신규설치수량"]).strip() or "0"}</div>', unsafe_allow_html=True)
                with row4_col2:
                    st.markdown('<div class="detail-label">기설치수량</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["기설치수량"]).strip() or "0"}</div>', unsafe_allow_html=True)
                with row4_col3:
                    st.markdown('<div class="detail-label">실사담당자</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["실사담당자"]).strip() or "-"}</div>', unsafe_allow_html=True)

                row5_col1, row5_col2, row5_col3 = st.columns(3)
                with row5_col1:
                    st.markdown('<div class="detail-label">영업담당자</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["영업담당자"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row5_col2:
                    st.markdown('<div class="detail-label">영업담당자 연락처</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["영업담당연락처"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row5_col3:
                    st.markdown('<div class="detail-label">실사예정일</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["실사예정일"]).strip() or "-"}</div>', unsafe_allow_html=True)

                row6_col1, row6_col2, row6_col3 = st.columns(3)
                with row6_col1:
                    st.markdown('<div class="detail-label">실사완료일</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["실사완료일"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row6_col2:
                    st.markdown('<div class="detail-label">계약일</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["계약일"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row6_col3:
                    st.markdown('<div class="detail-label">계약수량</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["계약수량"]).strip() or "0"}</div>', unsafe_allow_html=True)

                row7_col1, row7_col2 = st.columns(2)
                with row7_col1:
                    st.markdown('<div class="detail-label">계약금액</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["계약금액"]).strip() or "0"}</div>', unsafe_allow_html=True)
                with row7_col2:
                    st.markdown('<div class="detail-label">미계약사유</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["미계약사유"]).strip() or "-"}</div>', unsafe_allow_html=True)

                st.markdown('</div>', unsafe_allow_html=True)

                # 2. 내용 정보
                st.markdown('<div class="detail-section">', unsafe_allow_html=True)
                st.markdown('<div class="detail-section-title">내용 정보</div>', unsafe_allow_html=True)

                st.markdown('<div class="detail-label">요청내용</div>', unsafe_allow_html=True)
                request_text = str(view_row["요청내용"]).strip()
                st.markdown(
                    f'<div class="detail-textarea">{request_text if request_text else "요청내용이 없습니다."}</div>',
                    unsafe_allow_html=True
                )

                st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)

                st.markdown('<div class="detail-label">비고</div>', unsafe_allow_html=True)
                note_text = str(view_row["비고"]).strip()
                st.markdown(
                    f'<div class="detail-textarea">{note_text if note_text else "비고가 없습니다."}</div>',
                    unsafe_allow_html=True
                )

                st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)

                st.markdown('<div class="detail-label">실사결과</div>', unsafe_allow_html=True)
                result_text = str(view_row["실사결과"]).strip()
                st.markdown(
                    f'<div class="detail-textarea">{result_text if result_text else "실사결과가 없습니다."}</div>',
                    unsafe_allow_html=True
                )

                st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)

                st.markdown('<div class="detail-label">특이사항</div>', unsafe_allow_html=True)
                special_text = str(view_row["특이사항"]).strip()
                st.markdown(
                    f'<div class="detail-textarea">{special_text if special_text else "특이사항이 없습니다."}</div>',
                    unsafe_allow_html=True
                )

                st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)

                st.markdown('<div class="detail-label">후속조치</div>', unsafe_allow_html=True)
                follow_text = str(view_row["후속조치"]).strip()
                st.markdown(
                    f'<div class="detail-textarea">{follow_text if follow_text else "후속조치가 없습니다."}</div>',
                    unsafe_allow_html=True
                )

                st.markdown('</div>', unsafe_allow_html=True)

                # 3. 첨부파일
                st.markdown('<div class="detail-section">', unsafe_allow_html=True)
                st.markdown('<div class="detail-section-title">첨부파일</div>', unsafe_allow_html=True)

                if str(view_row["첨부파일링크"]).strip():
                    file_name = str(view_row["첨부파일명"]).strip() or "첨부파일 열기"
                    st.link_button(file_name, str(view_row["첨부파일링크"]))
                else:
                    st.markdown('<div class="detail-box">첨부파일이 없습니다.</div>', unsafe_allow_html=True)

                st.markdown('</div>', unsafe_allow_html=True)

                st.markdown("---")

                if st.button("수정", use_container_width=True, key=f"insp_edit_mode_btn_{view_idx}"):
                    st.session_state.inspection_edit_mode = True
                    st.session_state.inspection_edit_target = view_idx
                    st.rerun()   

            else:

                st.markdown("### 📌 기본 정보 수정")

                req_date_raw = str(view_row["요청일"]).strip()
                parsed_req_date = pd.to_datetime(req_date_raw, errors="coerce")
                default_req_date = parsed_req_date.date() if pd.notna(parsed_req_date) else date.today()

                with st.form(f"inspection_edit_form_{view_idx}"):

                    e1, e2, e3 = st.columns(3)
                    edit_req_date = e1.date_input("요청일 수정", value=default_req_date)
                    edit_operator = e2.text_input("운영사 수정", value=str(view_row["운영사"]))
                    edit_name = e3.text_input("현장명 수정", value=str(view_row["현장명"]))

                    e4, e5, e6 = st.columns(3)
                    edit_addr = e4.text_input("현장주소 수정", value=str(view_row["현장주소"]))
                    edit_phone = e5.text_input("현장연락처 수정", value=str(view_row["현장연락처"]))
                    edit_product = e6.selectbox(
                        "상품구분 수정",
                        PRODUCT_OPTIONS,
                        index=PRODUCT_OPTIONS.index(view_row["상품구분"]) if view_row["상품구분"] in PRODUCT_OPTIONS else 0
                    )

                    e6_1, e6_2 = st.columns(2)
                    current_env = str(view_row["환경부"]).strip() if "환경부" in view_row.index else ""
                    current_jatu = str(view_row["자투"]).strip() if "자투" in view_row.index else ""

                    env_index = ENV_OPTIONS.index(current_env) if current_env in ENV_OPTIONS else 0
                    jatu_index = JATU_OPTIONS.index(current_jatu) if current_jatu in JATU_OPTIONS else 0

                    env_gov = e6_1.selectbox("환경부 수정", ENV_OPTIONS, index=env_index, key=f"edit_env_{view_idx}")
                    jatu = e6_2.selectbox("자투 수정", JATU_OPTIONS, index=jatu_index, key=f"edit_jatu_{view_idx}")

                    e7, e8, e9 = st.columns(3)
                    edit_parking = e7.number_input("주차면수 수정", min_value=0, step=1, value=safe_int(view_row["주차면수"], 0))
                    edit_new_qty = e8.number_input("신규설치수량 수정", min_value=0, step=1, value=safe_int(view_row["신규설치수량"], 0))
                    edit_old_qty = e9.number_input("기설치수량 수정", min_value=0, step=1, value=safe_int(view_row["기설치수량"], 0))

                    e10, e11 = st.columns(2)
                    edit_sales = e10.text_input("영업담당자 수정", value=str(view_row["영업담당자"]))
                    edit_sales_phone = e11.text_input("영업담당자 연락처 수정", value=str(view_row["영업담당연락처"]))

                    edit_request = st.text_area("요청내용 수정", value=str(view_row["요청내용"]))
                    edit_note = st.text_input("비고 수정", value=str(view_row["비고"]))

                    st.subheader("첨부파일 수정")
                    edit_uploaded_file = st.file_uploader(
                        "새 첨부파일 업로드",
                        type=["pdf", "png", "jpg", "jpeg", "xlsx", "xls", "doc", "docx"],
                        key=f"insp_edit_uploaded_file_{view_idx}"
                    )

                    current_file_name = str(view_row["첨부파일명"]).strip()
                    current_file_link = str(view_row["첨부파일링크"]).strip()

                    delete_current_file = False

                    if current_file_link:
                        st.caption(f"현재 첨부파일: {current_file_name if current_file_name else '첨부파일'}")
                        delete_current_file = st.checkbox(
                            "기존 첨부파일 삭제",
                            key=f"delete_current_file_{view_idx}"
                        )

                    s1, s2 = st.columns(2)
                    save_submit = s1.form_submit_button("기본 정보 수정 저장", use_container_width=True)
                    cancel_submit = s2.form_submit_button("취소", use_container_width=True)

                    if save_submit:
                        target_row = df.loc[df["row_id"] == view_idx].iloc[0]

                        full_df = load_inspection_data()
                        full_df = full_df[INSPECTION_COLUMNS].copy()

                        original_idx = find_original_inspection_index(full_df, target_row)

                        if original_idx is None:
                            st.error("원본 데이터를 찾지 못했습니다.")
                            st.stop()

                        full_df.loc[original_idx, "요청일"] = str(edit_req_date)
                        full_df.loc[original_idx, "운영사"] = edit_operator.strip()
                        full_df.loc[original_idx, "현장명"] = edit_name.strip()
                        full_df.loc[original_idx, "현장주소"] = edit_addr.strip()
                        full_df.loc[original_idx, "현장연락처"] = edit_phone.strip()
                        full_df.loc[original_idx, "상품구분"] = edit_product
                        full_df.loc[original_idx, "주차면수"] = int(edit_parking)
                        full_df.loc[original_idx, "신규설치수량"] = int(edit_new_qty)
                        full_df.loc[original_idx, "기설치수량"] = int(edit_old_qty)
                        full_df.loc[original_idx, "영업담당자"] = edit_sales.strip()
                        full_df.loc[original_idx, "영업담당연락처"] = edit_sales_phone.strip()
                        full_df.loc[original_idx, "요청내용"] = edit_request.strip()
                        full_df.loc[original_idx, "비고"] = edit_note.strip()
                        full_df.loc[original_idx, "환경부"] = env_gov
                        full_df.loc[original_idx, "자투"] = jatu

                        save_inspection_data(full_df)

                        if delete_current_file:
                            full_df.loc[original_idx, "첨부파일명"] = ""
                            full_df.loc[original_idx, "첨부파일링크"] = ""

                        if edit_uploaded_file is not None:
                            try:
                                new_attachment_name, new_attachment_link = upload_file_to_drive(
                                    edit_uploaded_file,
                                    folder_id="13W2N1v9IBiuZEstmTrvt57Zg8XQiHt7J"
                                )
                                full_df.loc[original_idx, "첨부파일명"] = new_attachment_name
                                full_df.loc[original_idx, "첨부파일링크"] = new_attachment_link
                            except Exception as e:
                                st.error(f"첨부파일 업로드 실패: {e}")
                                st.stop()

                        save_inspection_data(full_df)

                        st.session_state.inspection_edit_mode = False
                        st.session_state.inspection_edit_target = None
                        set_inspection_flash("기본 정보 수정 완료!", "success")
                        st.rerun()

                    if cancel_submit:
                        st.session_state.inspection_edit_mode = False
                        st.session_state.inspection_edit_target = None
                        st.rerun()

    st.divider()

    with st.expander("🗑️ 7. 실사 요청 삭제", expanded=False):
        if df.empty:
            st.info("삭제할 내역이 없습니다.")
        else:
            delete_options = [
                f"{row['row_id']} | {row['현장명']} | {row['상품구분']} | {row['영업담당자']}"
                for _, row in df.iterrows()
            ]

            selected_delete = st.selectbox("삭제할 내역 선택", delete_options, key="insp_delete_select_new")
            confirm_delete = st.checkbox("정말 삭제합니다. 되돌리기 어렵습니다.", key="insp_delete_confirm_new")

            if st.button("선택 내역 삭제", key="insp_delete_btn_new"):
                if not confirm_delete:
                    st.warning("삭제 확인 체크를 먼저 해주세요.")
                else:
                    delete_idx = int(selected_delete.split("|")[0].strip())
                    target_row = df.loc[df["row_id"] == delete_idx].iloc[0]

                    full_df = load_inspection_data()
                    full_df = full_df[INSPECTION_COLUMNS].copy()

                    original_idx = find_original_inspection_index(full_df, target_row)

                    if original_idx is None:
                        st.error("원본 데이터를 찾지 못했습니다.")
                        st.stop()

                    full_df = full_df.drop(index=original_idx).reset_index(drop=True)

                    save_inspection_data(full_df)

                    set_inspection_flash("삭제가 완료되었습니다.", "success")
                    st.rerun()

# =========================================================
# 유지보수 관리 시스템
# =========================================================
MAINTENANCE_SHEET_NAME = "아이센서유지보수"

MAINTENANCE_COLUMNS = [
    "코드번호",
    "단지명",
    "연락처",
    "지역",
    "영업담당자",
    "수량",
    "단가",
    "계약시작일",
    "계약종료일",
    "총계약금액",
    "계약상태",
    "청구주기",
    "비고",
    "첨부파일명",
    "첨부파일링크"
]


def get_maintenance_sheet():
    client = get_gsheet_client()   # 🔥 이거 반드시 필요
    return client.open(MAINTENANCE_SHEET_NAME).sheet1

def get_maintenance_payment_sheet():
    client = get_gsheet_client()   # 🔥 이것도
    return client.open(MAINTENANCE_PAYMENT_SHEET_NAME).sheet1

def ensure_maintenance_sheet_header(sheet):
    values = sheet.get_all_values()

    if not values:
        end_col = chr(64 + len(MAINTENANCE_COLUMNS)) if len(MAINTENANCE_COLUMNS) <= 26 else "O"
        sheet.update(f"A1:{end_col}1", [MAINTENANCE_COLUMNS])
        return

    header = [str(x).strip() for x in values[0]]

    if header != MAINTENANCE_COLUMNS:
        existing = pd.DataFrame(values[1:], columns=header if header else None)

        for col in MAINTENANCE_COLUMNS:
            if col not in existing.columns:
                existing[col] = ""

        existing = existing[MAINTENANCE_COLUMNS]
        save_maintenance_data(existing, sheet=sheet)


def maintenance_safe_int(value, default=0):
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return default
    return int(num)


def maintenance_safe_float(value, default=0.0):
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return default
    return float(num)


def format_currency(value):
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return "0"
    return f"{int(round(float(num))):,}"

def style_unpaid_amount(val):
    num = pd.to_numeric(val, errors="coerce")
    if pd.isna(num):
        return ""
    if num > 0:
        return "background-color: #f8d7da; color: #842029; font-weight: bold;"
    return ""

def get_contract_expiring_soon(df, within_days=60):
    if df.empty:
        return pd.DataFrame(columns=df.columns)

    temp_df = df.copy()
    temp_df["계약종료일_dt"] = pd.to_datetime(temp_df["계약종료일"], errors="coerce")

    today = pd.Timestamp(date.today())
    limit_day = today + pd.Timedelta(days=within_days)

    temp_df = temp_df[
        (temp_df["계약상태"].astype(str).str.strip() == "진행중") &
        (temp_df["계약종료일_dt"].notna()) &
        (temp_df["계약종료일_dt"] >= today) &
        (temp_df["계약종료일_dt"] <= limit_day)
    ].copy()

    return temp_df.sort_values("계약종료일_dt")


def calculate_total_contract_amount(qty, unit_price):
    return int(maintenance_safe_int(qty, 0) * maintenance_safe_float(unit_price, 0))


@st.cache_data(ttl=60)
def load_maintenance_data():
    sheet = get_maintenance_sheet()
    ensure_maintenance_sheet_header(sheet)

    data = sheet.get_all_records()
    df = pd.DataFrame(data)

    if df.empty:
        return pd.DataFrame(columns=MAINTENANCE_COLUMNS)

    for col in MAINTENANCE_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[MAINTENANCE_COLUMNS].copy()

    df["수량"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0).astype(int)
    df["단가"] = pd.to_numeric(df["단가"], errors="coerce").fillna(0)
    df["총계약금액"] = pd.to_numeric(df["총계약금액"], errors="coerce").fillna(0)
    df["계약시작일"] = df["계약시작일"].astype(str)
    df["계약종료일"] = df["계약종료일"].astype(str)
    df["계약상태"] = df["계약상태"].astype(str).replace("", "진행중")
    df["청구주기"] = df["청구주기"].astype(str).replace("", "매월")

    return df


def save_maintenance_data(df, sheet=None):
    if sheet is None:
        sheet = get_maintenance_sheet()

    save_df = df.copy()

    if "row_id" in save_df.columns:
        save_df = save_df.drop(columns=["row_id"])

    for col in MAINTENANCE_COLUMNS:
        if col not in save_df.columns:
            save_df[col] = ""

    save_df = save_df[MAINTENANCE_COLUMNS].fillna("")

    save_df["수량"] = pd.to_numeric(save_df["수량"], errors="coerce").fillna(0).astype(int)
    save_df["단가"] = pd.to_numeric(save_df["단가"], errors="coerce").fillna(0)
    save_df["총계약금액"] = pd.to_numeric(save_df["총계약금액"], errors="coerce").fillna(0)

    rows = [save_df.columns.tolist()] + save_df.astype(str).values.tolist()

    old_values = sheet.get_all_values()
    old_data_rows = max(0, len(old_values) - 1)

    sheet.update("A1", rows)

    new_data_rows = len(save_df)
    total_cols = len(MAINTENANCE_COLUMNS)

    if old_data_rows > new_data_rows:
        blank_rows = old_data_rows - new_data_rows
        start_row = new_data_rows + 2
        end_row = old_data_rows + 1
        end_col_letter = chr(64 + total_cols) if total_cols <= 26 else "O"
        clear_range = f"A{start_row}:{end_col_letter}{end_row}"
        empty_values = [[""] * total_cols for _ in range(blank_rows)]
        sheet.update(clear_range, empty_values)

    st.cache_data.clear()

MAINTENANCE_PAYMENT_SHEET_NAME = "아이센서유지보수_수금관리"

MAINTENANCE_PAYMENT_COLUMNS = [
    "코드번호",
    "단지명",
    "기준년월",
    "청구금액",
    "발행여부",
    "발행일",
    "입금여부",
    "입금일",
    "미수금",
    "영업담당자",
    "계약상태",
    "비고"
]


def get_maintenance_payment_sheet():
    client = get_gsheet_client()
    return client.open(MAINTENANCE_PAYMENT_SHEET_NAME).sheet1


def ensure_maintenance_payment_sheet_header(sheet):
    values = sheet.get_all_values()

    if not values:
        sheet.update("A1:L1", [MAINTENANCE_PAYMENT_COLUMNS])
        return

    header = [str(x).strip() for x in values[0]]

    if header != MAINTENANCE_PAYMENT_COLUMNS:
        existing = pd.DataFrame(values[1:], columns=header if header else None)

        for col in MAINTENANCE_PAYMENT_COLUMNS:
            if col not in existing.columns:
                existing[col] = ""

        existing = existing[MAINTENANCE_PAYMENT_COLUMNS]
        save_maintenance_payment_data(existing, sheet=sheet)


@st.cache_data(ttl=60)
def load_maintenance_payment_data():
    sheet = get_maintenance_payment_sheet()
    ensure_maintenance_payment_sheet_header(sheet)

    data = sheet.get_all_records()
    df = pd.DataFrame(data)

    if df.empty:
        return pd.DataFrame(columns=MAINTENANCE_PAYMENT_COLUMNS)

    for col in MAINTENANCE_PAYMENT_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[MAINTENANCE_PAYMENT_COLUMNS].copy()

    df["청구금액"] = pd.to_numeric(df["청구금액"], errors="coerce").fillna(0)
    df["미수금"] = pd.to_numeric(df["미수금"], errors="coerce").fillna(0)

    df["발행여부"] = df["발행여부"].astype(str).replace("", "미발행")
    df["입금여부"] = df["입금여부"].astype(str).replace("", "미입금")
    df["기준년월"] = df["기준년월"].astype(str)
    df["발행일"] = df["발행일"].astype(str)
    df["입금일"] = df["입금일"].astype(str)
    df["계약상태"] = df["계약상태"].astype(str)

    return df


def save_maintenance_payment_data(df, sheet=None):
    if sheet is None:
        sheet = get_maintenance_payment_sheet()

    save_df = df.copy()

    if "row_id" in save_df.columns:
        save_df = save_df.drop(columns=["row_id"])

    for col in MAINTENANCE_PAYMENT_COLUMNS:
        if col not in save_df.columns:
            save_df[col] = ""

    save_df = save_df[MAINTENANCE_PAYMENT_COLUMNS].fillna("")

    save_df["청구금액"] = pd.to_numeric(save_df["청구금액"], errors="coerce").fillna(0)
    save_df["미수금"] = pd.to_numeric(save_df["미수금"], errors="coerce").fillna(0)

    rows = [save_df.columns.tolist()] + save_df.astype(str).values.tolist()

    old_values = sheet.get_all_values()
    old_data_rows = max(0, len(old_values) - 1)

    sheet.update("A1", rows)

    new_data_rows = len(save_df)

    if old_data_rows > new_data_rows:
        blank_rows = old_data_rows - new_data_rows
        start_row = new_data_rows + 2
        end_row = old_data_rows + 1
        clear_range = f"A{start_row}:L{end_row}"
        empty_values = [[""] * len(MAINTENANCE_PAYMENT_COLUMNS) for _ in range(blank_rows)]
        sheet.update(clear_range, empty_values)

    st.cache_data.clear()

def maintenance_page():
    render_inspection_common_style()

    st.markdown('<div class="erp-page-title">아이센서 유지보수관리 프로그램</div>', unsafe_allow_html=True)
    st.markdown('<div class="erp-page-desc">유지보수 계약등록, 월별 청구/수금, 미수금 관리</div>', unsafe_allow_html=True)
    st.markdown("""
    <style>
    .maintenance-alert-box {
        border-radius: 14px;
        padding: 14px 18px;
        margin: 8px 0 18px 0;
        font-size: 15px;
        font-weight: 600;
    }
    .maintenance-alert-danger {
        background: #fff1f2;
        border: 1px solid #fecdd3;
        color: #9f1239;
    }
    .maintenance-alert-warning {
        background: #fffbeb;
        border: 1px solid #fde68a;
        color: #92400e;
    }
    .maintenance-guide-box {
        border: 1px solid #e5e7eb;
        background: #ffffff;
        border-radius: 14px;
        padding: 14px 18px;
        margin: 10px 0 18px 0;
        color: #334155;
        font-size: 14px;
        line-height: 1.7;
    }
    .section-gap {
        margin-top: 18px;
        margin-bottom: 18px;
    }
    </style>
    """, unsafe_allow_html=True)

    try:
        df = load_maintenance_data()
        pay_df = load_maintenance_payment_data()
    except Exception as e:
        st.error(f"유지보수 데이터를 불러오지 못했습니다: {e}")
        return

    df = df.reset_index(drop=True)
    df["row_id"] = df.index

    pay_df = pay_df.reset_index(drop=True)
    pay_df["row_id"] = pay_df.index

    current_ym = make_year_month(date.today().year, date.today().month)

    total_count = len(df)
    active_count = len(df[df["계약상태"].astype(str).str.strip() == "진행중"]) if not df.empty else 0
    total_qty = int(df["수량"].sum()) if not df.empty else 0
    total_amount = int(pd.to_numeric(df["총계약금액"], errors="coerce").fillna(0).sum()) if not df.empty else 0

    unpaid_df = pay_df[pay_df["입금여부"].astype(str).str.strip() != "입금완료"].copy() if not pay_df.empty else pd.DataFrame()
    total_unpaid = int(pd.to_numeric(unpaid_df["미수금"], errors="coerce").fillna(0).sum()) if not unpaid_df.empty else 0

    expiring_df = get_contract_expiring_soon(df, within_days=60)
    expiring_count = len(expiring_df)

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("전체 계약", total_count)
    c2.metric("진행중 계약", active_count)
    c3.metric("총 수량", total_qty)
    c4.metric("전체 계약금액", format_currency(total_amount))
    c5.metric("전체 미수금", format_currency(total_unpaid))
    c6.metric("60일 내 종료예정", expiring_count)
    if total_unpaid > 0:
        st.markdown(
            f'<div class="maintenance-alert-box maintenance-alert-danger">⚠️ 현재 미수금 총액: {format_currency(total_unpaid)} 원</div>',
            unsafe_allow_html=True
        )

    if expiring_count > 0:
        st.markdown(
            f'<div class="maintenance-alert-box maintenance-alert-warning">⏰ 60일 이내 종료 예정 계약이 {expiring_count}건 있습니다.</div>',
            unsafe_allow_html=True
        )

    st.markdown("""
    <div class="maintenance-guide-box">
    • 유지보수 계약 등록 → 월별 청구 생성 → 청구/수금 처리 순서로 사용하시면 됩니다.<br>
    • 미수금이 있는 건은 월별 청구/수금 현황에서 빨간색으로 표시됩니다.<br>
    • 계약 종료 예정 건은 별도 섹션에서 빠르게 확인할 수 있습니다.
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    with st.expander("📝 1. 유지보수 계약 등록", expanded=False):
        with st.form("maintenance_add_form"):
            a1, a2, a3 = st.columns(3)
            code_no = a1.text_input("코드번호", key="mt_code_no")
            site_name = a2.text_input("단지명", key="mt_site_name")
            phone = a3.text_input("연락처", key="mt_phone")

            a4, a5, a6 = st.columns(3)
            region = a4.text_input("지역", key="mt_region")
            sales_manager = a5.text_input("영업담당자", key="mt_sales_manager")
            qty = a6.number_input("수량", min_value=0, step=1, value=0, key="mt_qty")

            a7, a8, a9 = st.columns(3)
            unit_price = a7.number_input("단가", min_value=0, step=1000, value=0, key="mt_unit_price")
            start_date = a8.date_input("계약시작일", value=date.today(), key="mt_start_date")
            end_date = a9.date_input("계약종료일", value=date.today(), key="mt_end_date")

            b1, b2, b3 = st.columns(3)
            contract_status = b1.selectbox("계약상태", ["진행중", "종료", "해지"], key="mt_contract_status")
            billing_cycle = b2.selectbox("청구주기", ["매월", "분기", "반기", "연간"], key="mt_billing_cycle")
            note = b3.text_input("비고", key="mt_note")

            uploaded_file = st.file_uploader("계약서 첨부", type=None, key="mt_uploaded_file")

            total_amount_preview = calculate_total_contract_amount(qty, unit_price)
            st.info(f"자동 계산 총계약금액: {format_currency(total_amount_preview)} 원")

            submitted = st.form_submit_button("등록하기")

            if submitted:
                if not site_name.strip():
                    st.warning("단지명을 입력해주세요.")
                else:
                    attachment_name = ""
                    attachment_link = ""

                    if uploaded_file is not None:
                        try:
                            attachment_name, attachment_link = upload_file_to_drive(uploaded_file)
                        except Exception as e:
                            st.error(f"첨부파일 업로드 실패: {e}")
                            st.stop()

                    new_row = pd.DataFrame([{
                        "코드번호": code_no.strip(),
                        "단지명": site_name.strip(),
                        "연락처": phone.strip(),
                        "지역": region.strip(),
                        "영업담당자": sales_manager.strip(),
                        "수량": int(qty),
                        "단가": float(unit_price),
                        "계약시작일": str(start_date),
                        "계약종료일": str(end_date),
                        "총계약금액": calculate_total_contract_amount(qty, unit_price),
                        "계약상태": contract_status,
                        "청구주기": billing_cycle,
                        "비고": note.strip(),
                        "첨부파일명": attachment_name,
                        "첨부파일링크": attachment_link
                    }])

                    save_df = df[MAINTENANCE_COLUMNS].copy() if not df.empty else pd.DataFrame(columns=MAINTENANCE_COLUMNS)
                    save_df = pd.concat([save_df, new_row], ignore_index=True)
                    save_maintenance_data(save_df)

                    st.success("유지보수 계약 등록 완료!")
                    st.rerun()

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("🧾 2. 월별 청구 생성", expanded=False):
        g1, g2 = st.columns(2)
        gen_year = g1.number_input("청구 생성 연도", min_value=2024, max_value=2100, value=date.today().year, step=1, key="mt_gen_year")
        gen_month = g2.selectbox("청구 생성 월", list(range(1, 13)), index=date.today().month - 1, key="mt_gen_month")

        target_ym = make_year_month(gen_year, gen_month)
        st.info(f"{target_ym} 기준으로 진행중 계약의 월 청구를 생성합니다. 이미 같은 월 자료가 있으면 중복 생성되지 않습니다.")

        if st.button("해당 월 청구 일괄 생성", key="mt_generate_claim_btn"):
            new_claim_df = generate_monthly_claim_rows(df, pay_df, gen_year, gen_month)

            if new_claim_df.empty:
                st.warning("새로 생성할 청구 자료가 없습니다. 이미 생성되었거나 진행중 계약이 없습니다.")
            else:
                base_df = pay_df[MAINTENANCE_PAYMENT_COLUMNS].copy() if not pay_df.empty else pd.DataFrame(columns=MAINTENANCE_PAYMENT_COLUMNS)
                base_df = pd.concat([base_df, new_claim_df], ignore_index=True)
                save_maintenance_payment_data(base_df)
                st.success(f"{target_ym} 청구자료 {len(new_claim_df)}건 생성 완료!")
                st.rerun()

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("💰 3. 월별 청구/수금 처리", expanded=False):
        if pay_df.empty:
            st.info("생성된 청구자료가 없습니다.")
        else:
            payment_options = [
                f"{row['row_id']} | {row['기준년월']} | {row['코드번호']} | {row['단지명']}"
                for _, row in pay_df.iterrows()
            ]

            selected_payment = st.selectbox("처리할 청구 선택", payment_options, key="mt_payment_select")
            pay_idx = int(selected_payment.split("|")[0].strip())
            pay_row = pay_df.loc[pay_df["row_id"] == pay_idx].iloc[0]

            with st.form(f"mt_payment_form_{pay_idx}"):
                p1, p2, p3 = st.columns(3)
                edit_claim_amount = p1.number_input("청구금액", min_value=0, step=1000, value=int(maintenance_safe_float(pay_row["청구금액"], 0)))
                issue_status = p2.selectbox(
                    "발행여부",
                    ["미발행", "발행완료"],
                    index=0 if str(pay_row["발행여부"]) != "발행완료" else 1
                )
                deposit_status = p3.selectbox(
                    "입금여부",
                    ["미입금", "입금완료"],
                    index=0 if str(pay_row["입금여부"]) != "입금완료" else 1
                )

                p4, p5, p6 = st.columns(3)

                default_issue_date = pd.to_datetime(pay_row["발행일"], errors="coerce")
                if pd.isna(default_issue_date):
                    default_issue_date = pd.Timestamp(date.today())

                default_deposit_date = pd.to_datetime(pay_row["입금일"], errors="coerce")
                if pd.isna(default_deposit_date):
                    default_deposit_date = pd.Timestamp(date.today())

                issue_date = p4.date_input("발행일", value=default_issue_date.date(), key=f"mt_issue_date_{pay_idx}")
                deposit_date = p5.date_input("입금일", value=default_deposit_date.date(), key=f"mt_deposit_date_{pay_idx}")
                payment_note = p6.text_input("비고", value=str(pay_row["비고"]), key=f"mt_payment_note_{pay_idx}")

                unpaid_preview = calculate_unpaid_amount(edit_claim_amount, deposit_status)
                st.info(f"처리 후 미수금: {format_currency(unpaid_preview)} 원")

                payment_submit = st.form_submit_button("저장하기")

                if payment_submit:
                    save_pay_df = pay_df[MAINTENANCE_PAYMENT_COLUMNS].copy()

                    save_pay_df.loc[pay_idx, "청구금액"] = float(edit_claim_amount)
                    save_pay_df.loc[pay_idx, "발행여부"] = issue_status
                    save_pay_df.loc[pay_idx, "발행일"] = str(issue_date) if issue_status == "발행완료" else ""
                    save_pay_df.loc[pay_idx, "입금여부"] = deposit_status
                    save_pay_df.loc[pay_idx, "입금일"] = str(deposit_date) if deposit_status == "입금완료" else ""
                    save_pay_df.loc[pay_idx, "미수금"] = calculate_unpaid_amount(edit_claim_amount, deposit_status)
                    save_pay_df.loc[pay_idx, "비고"] = payment_note.strip()

                    save_maintenance_payment_data(save_pay_df)
                    st.success("청구/수금 처리 완료!")
                    st.rerun()

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("📋 4. 월별 청구/수금 현황", expanded=False):
        f1, f2, f3 = st.columns(3)
        filter_year = f1.number_input("조회 연도", min_value=2024, max_value=2100, value=date.today().year, step=1, key="mt_filter_year")
        filter_month = f2.selectbox("조회 월", list(range(1, 13)), index=date.today().month - 1, key="mt_filter_month")
        filter_deposit = f3.selectbox("입금상태", ["전체", "미입금", "입금완료"], key="mt_filter_deposit")

        target_ym = make_year_month(filter_year, filter_month)
        view_df = pay_df.copy()

        if not view_df.empty:
            view_df = view_df[view_df["기준년월"] == target_ym]

            if filter_deposit != "전체":
                view_df = view_df[view_df["입금여부"] == filter_deposit]

        month_claim_total = int(pd.to_numeric(view_df["청구금액"], errors="coerce").fillna(0).sum()) if not view_df.empty else 0
        month_unpaid_total = int(pd.to_numeric(view_df["미수금"], errors="coerce").fillna(0).sum()) if not view_df.empty else 0
        month_paid_count = len(view_df[view_df["입금여부"] == "입금완료"]) if not view_df.empty else 0

        m1, m2, m3 = st.columns(3)
        m1.metric("월 청구금액", format_currency(month_claim_total))
        m2.metric("월 미수금", format_currency(month_unpaid_total))
        m3.metric("입금완료 건수", month_paid_count)

        if view_df.empty:
            st.info("조건에 맞는 수금자료가 없습니다.")
        else:
            show_pay_df = view_df.copy()

            show_pay_df["청구금액"] = pd.to_numeric(show_pay_df["청구금액"], errors="coerce").fillna(0)
            show_pay_df["미수금"] = pd.to_numeric(show_pay_df["미수금"], errors="coerce").fillna(0)

            display_df = show_pay_df[
                [
                    "기준년월", "코드번호", "단지명", "청구금액",
                    "발행여부", "발행일", "입금여부", "입금일",
                    "미수금", "영업담당자", "계약상태", "비고"
                ]
            ].copy()

            # 숫자 포맷
            display_df["청구금액"] = display_df["청구금액"].apply(format_currency)
            display_df["미수금"] = display_df["미수금"].apply(format_currency)

            # 👉 스타일 적용 (핵심)
            styled_df = display_df.style.map(
                style_unpaid_amount,
                subset=["미수금"]
            )

            st.dataframe(styled_df, use_container_width=True, hide_index=True) 

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("📄 5. 계약 현황", expanded=False):
        f1, f2, f3 = st.columns(3)
        region_filter = f1.selectbox(
            "지역 선택",
            ["전체"] + sorted([x for x in df["지역"].astype(str).unique().tolist() if str(x).strip() != ""]),
            key="mt_region_filter"
        )
        status_filter = f2.selectbox(
            "계약상태 선택",
            ["전체", "진행중", "종료", "해지"],
            key="mt_status_filter"
        )
        keyword = f3.text_input("검색", placeholder="코드번호 / 단지명 / 담당자", key="mt_keyword")

        filtered_df = df.copy()

        if region_filter != "전체":
            filtered_df = filtered_df[filtered_df["지역"] == region_filter]

        if status_filter != "전체":
            filtered_df = filtered_df[filtered_df["계약상태"] == status_filter]

        if keyword.strip():
            kw = keyword.strip()
            filtered_df = filtered_df[
                filtered_df["코드번호"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["단지명"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["영업담당자"].astype(str).str.contains(kw, case=False, na=False)
            ]

        if filtered_df.empty:
            st.info("조건에 맞는 계약이 없습니다.")
        else:
            show_df = filtered_df.copy()
            show_df["단가"] = show_df["단가"].apply(format_currency)
            show_df["총계약금액"] = show_df["총계약금액"].apply(format_currency)

            st.dataframe(
                show_df[
                    [
                        "코드번호", "단지명", "연락처", "지역", "영업담당자",
                        "수량", "단가", "계약시작일", "계약종료일",
                        "총계약금액", "계약상태", "청구주기", "비고"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("📅 5-1. 계약 종료 예정", expanded=False):
        soon_df = get_contract_expiring_soon(df, within_days=60)

        if soon_df.empty:
            st.info("60일 이내 종료 예정 계약이 없습니다.")
        else:
            display_soon_df = soon_df[
                [
                    "코드번호", "단지명", "지역", "영업담당자",
                    "계약시작일", "계약종료일", "계약상태", "청구주기", "비고"
                ]
            ].copy()

            st.dataframe(display_soon_df, use_container_width=True, hide_index=True)

    st.divider()        
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("🔍 6. 상세 보기", expanded=False):
        if df.empty:
            st.info("조회할 계약이 없습니다.")
        else:
            view_options = [
                f"{row['row_id']} | {row['코드번호']} | {row['단지명']} | {row['영업담당자']}"
                for _, row in df.iterrows()
            ]
            selected_view = st.selectbox("조회할 계약 선택", view_options, key="mt_view_select")
            view_idx = int(selected_view.split("|")[0].strip())
            row = df.loc[df["row_id"] == view_idx].iloc[0]

            v1, v2 = st.columns(2)

            with v1:
                st.subheader("기본 정보")
                st.write(f"**코드번호**: {row['코드번호']}")
                st.write(f"**단지명**: {row['단지명']}")
                st.write(f"**연락처**: {row['연락처']}")
                st.write(f"**지역**: {row['지역']}")
                st.write(f"**영업담당자**: {row['영업담당자']}")
                st.write(f"**수량**: {row['수량']}")
                st.write(f"**단가**: {format_currency(row['단가'])} 원")
                st.write(f"**총계약금액**: {format_currency(row['총계약금액'])} 원")

            with v2:
                st.subheader("계약 정보")
                st.write(f"**계약시작일**: {row['계약시작일']}")
                st.write(f"**계약종료일**: {row['계약종료일']}")
                st.write(f"**계약상태**: {row['계약상태']}")
                st.write(f"**청구주기**: {row['청구주기']}")
                st.write(f"**비고**: {row['비고']}")

                if str(row["첨부파일링크"]).strip():
                    st.markdown(f"[📎 첨부파일 열기]({row['첨부파일링크']})")
                else:
                    st.caption("첨부파일 없음")

            st.subheader("월별 수금 이력")
            history_df = pay_df[pay_df["코드번호"].astype(str) == str(row["코드번호"]).strip()].copy()

            if history_df.empty:
                st.info("이 계약의 수금 이력이 없습니다.")
            else:
                history_df["청구금액"] = history_df["청구금액"].apply(format_currency)
                history_df["미수금"] = history_df["미수금"].apply(format_currency)
                st.dataframe(
                    history_df[
                        ["기준년월", "청구금액", "발행여부", "발행일", "입금여부", "입금일", "미수금", "비고"]
                    ],
                    use_container_width=True,
                    hide_index=True
                )

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("✏️ 7. 계약 수정", expanded=False):
        if df.empty:
            st.info("수정할 계약이 없습니다.")
        else:
            edit_options = [
                f"{row['row_id']} | {row['코드번호']} | {row['단지명']} | {row['영업담당자']}"
                for _, row in df.iterrows()
            ]
            selected_edit = st.selectbox("수정할 계약 선택", edit_options, key="mt_edit_select")
            edit_idx = int(selected_edit.split("|")[0].strip())
            edit_row = df.loc[df["row_id"] == edit_idx].iloc[0]

            default_start = pd.to_datetime(edit_row["계약시작일"], errors="coerce")
            if pd.isna(default_start):
                default_start = pd.Timestamp(date.today())

            default_end = pd.to_datetime(edit_row["계약종료일"], errors="coerce")
            if pd.isna(default_end):
                default_end = pd.Timestamp(date.today())

            with st.form(f"maintenance_edit_form_{edit_idx}"):
                e1, e2, e3 = st.columns(3)
                edit_code = e1.text_input("코드번호 수정", value=str(edit_row["코드번호"]))
                edit_site = e2.text_input("단지명 수정", value=str(edit_row["단지명"]))
                edit_phone = e3.text_input("연락처 수정", value=str(edit_row["연락처"]))

                e4, e5, e6 = st.columns(3)
                edit_region = e4.text_input("지역 수정", value=str(edit_row["지역"]))
                edit_sales = e5.text_input("영업담당자 수정", value=str(edit_row["영업담당자"]))
                edit_qty = e6.number_input("수량 수정", min_value=0, step=1, value=int(edit_row["수량"]))

                e7, e8, e9 = st.columns(3)
                edit_unit_price = e7.number_input("단가 수정", min_value=0, step=1000, value=int(maintenance_safe_float(edit_row["단가"], 0)))
                edit_start = e8.date_input("계약시작일 수정", value=default_start.date())
                edit_end = e9.date_input("계약종료일 수정", value=default_end.date())

                e10, e11, e12 = st.columns(3)
                edit_status = e10.selectbox(
                    "계약상태 수정",
                    ["진행중", "종료", "해지"],
                    index=["진행중", "종료", "해지"].index(edit_row["계약상태"]) if str(edit_row["계약상태"]) in ["진행중", "종료", "해지"] else 0
                )
                edit_cycle = e11.selectbox(
                    "청구주기 수정",
                    ["매월", "분기", "반기", "연간"],
                    index=["매월", "분기", "반기", "연간"].index(edit_row["청구주기"]) if str(edit_row["청구주기"]) in ["매월", "분기", "반기", "연간"] else 0
                )
                edit_note = e12.text_input("비고 수정", value=str(edit_row["비고"]))

                new_file = st.file_uploader("새 첨부파일 업로드", type=None, key=f"mt_new_file_{edit_idx}")

                total_amount_preview = calculate_total_contract_amount(edit_qty, edit_unit_price)
                st.info(f"수정 후 총계약금액: {format_currency(total_amount_preview)} 원")

                edit_submit = st.form_submit_button("수정 저장")

                if edit_submit:
                    save_df = df[MAINTENANCE_COLUMNS].copy()

                    old_code = str(save_df.loc[edit_idx, "코드번호"]).strip()
                    new_code = edit_code.strip()

                    save_df.loc[edit_idx, "코드번호"] = new_code
                    save_df.loc[edit_idx, "단지명"] = edit_site.strip()
                    save_df.loc[edit_idx, "연락처"] = edit_phone.strip()
                    save_df.loc[edit_idx, "지역"] = edit_region.strip()
                    save_df.loc[edit_idx, "영업담당자"] = edit_sales.strip()
                    save_df.loc[edit_idx, "수량"] = int(edit_qty)
                    save_df.loc[edit_idx, "단가"] = float(edit_unit_price)
                    save_df.loc[edit_idx, "계약시작일"] = str(edit_start)
                    save_df.loc[edit_idx, "계약종료일"] = str(edit_end)
                    save_df.loc[edit_idx, "총계약금액"] = calculate_total_contract_amount(edit_qty, edit_unit_price)
                    save_df.loc[edit_idx, "계약상태"] = edit_status
                    save_df.loc[edit_idx, "청구주기"] = edit_cycle
                    save_df.loc[edit_idx, "비고"] = edit_note.strip()

                    if new_file is not None:
                        try:
                            new_attachment_name, new_attachment_link = upload_file_to_drive(new_file)
                            save_df.loc[edit_idx, "첨부파일명"] = new_attachment_name
                            save_df.loc[edit_idx, "첨부파일링크"] = new_attachment_link
                        except Exception as e:
                            st.error(f"첨부파일 업로드 실패: {e}")
                            st.stop()

                    save_maintenance_data(save_df)

                    if not pay_df.empty:
                        save_pay_df = pay_df[MAINTENANCE_PAYMENT_COLUMNS].copy()
                        save_pay_df.loc[save_pay_df["코드번호"].astype(str) == old_code, "코드번호"] = new_code
                        save_pay_df.loc[save_pay_df["코드번호"].astype(str) == new_code, "단지명"] = edit_site.strip()
                        save_pay_df.loc[save_pay_df["코드번호"].astype(str) == new_code, "영업담당자"] = edit_sales.strip()
                        save_pay_df.loc[save_pay_df["코드번호"].astype(str) == new_code, "계약상태"] = edit_status
                        save_maintenance_payment_data(save_pay_df)

                    st.success("수정 완료!")
                    st.rerun()

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("🗑️ 8. 계약 삭제", expanded=False):
        if df.empty:
            st.info("삭제할 계약이 없습니다.")
        else:
            delete_options = [
                f"{row['row_id']} | {row['코드번호']} | {row['단지명']} | {row['영업담당자']}"
                for _, row in df.iterrows()
            ]
            selected_delete = st.selectbox("삭제할 계약 선택", delete_options, key="mt_delete_select")
            confirm_delete = st.checkbox("정말 삭제합니다. 되돌리기 어렵습니다.", key="mt_delete_confirm")

            if st.button("선택 계약 삭제", key="mt_delete_btn"):
                if not confirm_delete:
                    st.warning("삭제 확인 체크를 먼저 해주세요.")
                else:
                    delete_idx = int(selected_delete.split("|")[0].strip())
                    delete_code = str(df.loc[delete_idx, "코드번호"]).strip()

                    save_df = df[MAINTENANCE_COLUMNS].copy()
                    save_df = save_df.drop(index=delete_idx).reset_index(drop=True)
                    save_maintenance_data(save_df)

                    if not pay_df.empty and delete_code != "":
                        save_pay_df = pay_df[MAINTENANCE_PAYMENT_COLUMNS].copy()
                        save_pay_df = save_pay_df[save_pay_df["코드번호"].astype(str) != delete_code].reset_index(drop=True)
                        save_maintenance_payment_data(save_pay_df)

                    st.success("삭제 완료!")
                    st.rerun()   


def make_year_month(year, month):
    return f"{int(year)}-{int(month):02d}"


def calculate_unpaid_amount(claim_amount, deposit_status):
    claim = maintenance_safe_float(claim_amount, 0)
    if str(deposit_status).strip() == "입금완료":
        return 0
    return claim

def can_generate_claim_by_cycle(start_date_value, billing_cycle, target_year, target_month):
    """
    계약시작일과 청구주기를 기준으로 해당 년/월에 청구 생성 대상인지 판단
    """
    start_dt = pd.to_datetime(start_date_value, errors="coerce")

    if pd.isna(start_dt):
        return False

    start_index = start_dt.year * 12 + start_dt.month
    target_index = int(target_year) * 12 + int(target_month)

    if target_index < start_index:
        return False

    diff = target_index - start_index
    cycle = str(billing_cycle).strip()

    if cycle == "매월":
        return True
    elif cycle == "분기":
        return diff % 3 == 0
    elif cycle == "반기":
        return diff % 6 == 0
    elif cycle == "연간":
        return diff % 12 == 0

    # 값이 비어 있거나 예외면 매월로 처리
    return True


def is_contract_active_for_month(start_date_value, end_date_value, target_year, target_month):
    """
    해당 년/월이 계약기간 안에 있는지 판단
    """
    start_dt = pd.to_datetime(start_date_value, errors="coerce")
    end_dt = pd.to_datetime(end_date_value, errors="coerce")

    if pd.isna(start_dt):
        return False

    month_start = pd.Timestamp(year=int(target_year), month=int(target_month), day=1)
    month_end = month_start + pd.offsets.MonthEnd(0)

    if month_end < start_dt:
        return False

    if not pd.isna(end_dt) and month_start > end_dt:
        return False

    return True

def generate_monthly_claim_rows(contract_df, payment_df, target_year, target_month):
    target_ym = make_year_month(target_year, target_month)

    existing_keys = set()
    if not payment_df.empty:
        for _, row in payment_df.iterrows():
            key = (str(row["코드번호"]).strip(), str(row["기준년월"]).strip())
            existing_keys.add(key)

    new_rows = []

    active_df = contract_df[contract_df["계약상태"].astype(str).isin(["진행중"])].copy()

    for _, row in active_df.iterrows():
        code_no = str(row["코드번호"]).strip()
        site_name = str(row["단지명"]).strip()
        billing_cycle = str(row["청구주기"]).strip()
        start_date_value = row["계약시작일"]
        end_date_value = row["계약종료일"]

        if code_no == "" and site_name == "":
            continue

        # 1) 계약기간 안에 있어야 함
        if not is_contract_active_for_month(start_date_value, end_date_value, target_year, target_month):
            continue

        # 2) 청구주기에 맞는 월만 생성
        if not can_generate_claim_by_cycle(start_date_value, billing_cycle, target_year, target_month):
            continue

        key = (code_no, target_ym)
        if key in existing_keys:
            continue

        claim_amount = calculate_total_contract_amount(row["수량"], row["단가"])

        new_rows.append({
            "코드번호": code_no,
            "단지명": site_name,
            "기준년월": target_ym,
            "청구금액": claim_amount,
            "발행여부": "미발행",
            "발행일": "",
            "입금여부": "미입금",
            "입금일": "",
            "미수금": claim_amount,
            "영업담당자": str(row["영업담당자"]).strip(),
            "계약상태": str(row["계약상태"]).strip(),
            "비고": f"{billing_cycle} 자동생성"
        })

    return pd.DataFrame(new_rows)   


# =========================================================
# 페이지들
# =========================================================
def page_dashboard():
    render_inspection_common_style()
    render_common_style()
    st.title("📊 통합 대시보드")
    st.caption("영업 / 계약 / 시공 / 실사 / 유지보수 / 연차 / 오늘 할 일을 한 화면에서 확인합니다.")

    today = date.today()
    today_str = str(today)

    # =========================
    # 1. 기본 데이터 불러오기
    # =========================
    try:
        sales_df = apply_role_filter(load_df("영업현황")) if st.session_state.business == "아이센서" else pd.DataFrame()
        possible_df = apply_role_filter(load_df("가능단지")) if st.session_state.business == "아이센서" else pd.DataFrame()
        bid_df = apply_role_filter(load_df("입찰공고")) if st.session_state.business == "아이센서" else pd.DataFrame()
        contract_df = apply_role_filter(load_df("계약단지")) if st.session_state.business == "아이센서" else apply_role_filter(load_df("계약접수현황"))

        task_df = apply_author_filter(load_tasks_df())
        schedule_common_df = apply_author_filter(load_schedule_df())

    except Exception as e:
        st.error(f"대시보드 기본 데이터를 불러오지 못했습니다: {e}")
        return

    # =========================
    # 2. 상단 핵심 KPI
    # =========================
    st.subheader("📌 핵심 현황")

    k1, k2, k3, k4, k5 = st.columns(5)

    with k1:
        ui_card("영업현황", len(sales_df) if not sales_df.empty else 0, "전체 영업 데이터", "info")

    with k2:
        ui_card("계약/접수", len(contract_df) if not contract_df.empty else 0, "계약·접수 현황", "success")

    with k3:
        ui_card("가능단지", len(possible_df) if not possible_df.empty else 0, "가능 단지 수", "info")

    with k4:
        ui_card("입찰공고", len(bid_df) if not bid_df.empty else 0, "입찰 공고 수", "warning")

    with k5:
        ui_card("오늘 할 일", len(task_df) if not task_df.empty else 0, "등록된 할 일", "danger")

    st.divider()

    # =========================
    # 3. 시공 일정 요약
    # =========================
    st.subheader("🛠 시공 일정 요약")

    try:
        construction_df = load_schedule_data()

        if construction_df is None:
            construction_df = pd.DataFrame()
        else:
            construction_df = construction_df.copy()
            construction_df = apply_product_filter(construction_df)

        if construction_df.empty:
            c1, c2, c3, c4 = st.columns(4)

            with c1:
                ui_card("전체 일정", 0, "전체 시공 일정")

            with c2:
                ui_card("오늘 일정", 0, "오늘 시공 예정",)

            with c3:
                ui_card("진행중", 0, "미완료 일정",)

            with c4:
                ui_card("완료", 0, "완료된 일정",)

            st.info("등록된 시공 일정이 없습니다.")

        else:
            today_construction = construction_df[construction_df["날짜"].astype(str) == today_str]
            progress_construction = construction_df[construction_df["상태"].astype(str) == "진행중"]
            done_construction = construction_df[construction_df["상태"].astype(str) == "완료"]

            c1, c2, c3, c4 = st.columns(4)

            with c1:
                ui_card("전체 일정", len(construction_df), "전체 시공 일정")

            with c2:
                ui_card("오늘 일정", len(today_construction), "오늘 시공 예정")

            with c3:
                ui_card("진행중", len(progress_construction), "미완료 일정")

            with c4:
                ui_card("완료", len(done_construction), "완료된 일정")

            # st.info(
            #     f"시공 현황: 전체 {total_count}건 / "
            #     f"오늘 {today_count}건 / "
            #     f"진행중 {progress_count}건 / "
            #     f"완료 {done_count}건"
            # )

            if not today_construction.empty:
                st.subheader("📅 오늘 시공 일정")
                st.dataframe(
                    today_construction[["날짜","상품구분", "설치현장", "시공담당", "수량", "비고", "상태"]],
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("오늘 시공 일정은 없습니다.")

    except Exception as e:
        st.warning(f"시공 일정 요약을 불러오지 못했습니다: {e}")

    st.divider()

    # =========================
    # 4. 실사 관리 요약
    # =========================
    st.subheader("🔎 실사 관리 요약")

    try:
        inspection_df = load_inspection_data()
        inspection_df = normalize_inspection_df(inspection_df)
        inspection_df = apply_product_filter(inspection_df)


        if inspection_df.empty:
            i1, i2, i3, i4, i5 = st.columns(5)

            with i1:
                ui_card("전체 요청", 0,"전체 실사 요청",)

            with i2:
                ui_card("요청접수", 0,"신규 요청",)

            with i3:
                ui_card("진행중", 0,"배정/일정/진행",)

            with i4:
                ui_card("실사완료", 0,"완료된 실사",)

            with i5:
                ui_card("계약완료", 0,"계약 전환",)

            st.info("등록된 실사 요청이 없습니다.")

        else:
            pending_count = len(inspection_df[inspection_df["진행상태"] == "요청접수"])
            working_count = len(inspection_df[inspection_df["진행상태"].isin(["담당자배정", "일정확정", "실사진행"])])
            done_count = len(inspection_df[inspection_df["진행상태"] == "실사완료"])
            contract_done_count = len(inspection_df[inspection_df["계약여부"] == "계약"])

            i1, i2, i3, i4, i5 = st.columns(5)

            with i1:
                ui_card("전체 요청", len(inspection_df),"전체 실사 요청")

            with i2:
                ui_card("요청접수", pending_count,"신규 요청")

            with i3:
                ui_card("진행중", working_count,"배정/일정/진행")

            with i4:
                ui_card("실사완료", done_count,"완료된 실사")

            with i5:
                ui_card("계약완료", contract_done_count,"계약 전환")

            # st.info(
            #     f"실사 현황: 전체 {len(inspection_df)}건 / "
            #     f"요청 {pending_count}건 / "
            #     f"진행중 {working_count}건 / "
            #     f"완료 {done_count}건"
            # )

            recent_insp = inspection_df.tail(10).copy()
            st.subheader("최근 실사 요청")
            st.dataframe(
                recent_insp[[
                    "요청일", "현장명", "상품구분", "영업담당자",
                    "실사담당자", "실사예정일", "진행상태", "계약여부"
                ]],
                use_container_width=True,
                hide_index=True
            )

    except Exception as e:
        st.warning(f"실사 관리 요약을 불러오지 못했습니다: {e}")

    st.divider()

    # =========================
    # 5. 유지보수 / 수금 요약
    # =========================
    if st.session_state.business == "아이센서":

        st.subheader("📡 유지보수 / 수금 요약")

        try:
            maintenance_df = load_maintenance_data()
            payment_df = load_maintenance_payment_data()

            active_count = 0
            total_unpaid = 0
            expiring_count = 0

            if not maintenance_df.empty:
                active_count = len(maintenance_df[maintenance_df["계약상태"].astype(str).str.strip() == "진행중"])
                expiring_count = len(get_contract_expiring_soon(maintenance_df, within_days=60))

            if not payment_df.empty:
                unpaid_df = payment_df[payment_df["입금여부"].astype(str).str.strip() != "입금완료"].copy()
                total_unpaid = int(pd.to_numeric(unpaid_df["미수금"], errors="coerce").fillna(0).sum())
            else:
                unpaid_df = pd.DataFrame()

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("전체 유지보수 계약", len(maintenance_df) if not maintenance_df.empty else 0)
            m2.metric("진행중 계약", active_count)
            m3.metric("전체 미수금", f"{total_unpaid:,} 원")
            m4.metric("60일 내 종료예정", expiring_count)

            if total_unpaid > 0:
                st.warning(f"현재 유지보수 미수금이 {total_unpaid:,}원 있습니다.")

            if not unpaid_df.empty:
                st.subheader("💰 미입금 현황")
                show_unpaid = unpaid_df[[
                    "기준년월", "코드번호", "단지명", "청구금액",
                    "입금여부", "미수금", "영업담당자"
                ]].copy()

                st.dataframe(show_unpaid.head(20), use_container_width=True, hide_index=True)

        except Exception as e:
            st.warning(f"유지보수 / 수금 요약을 불러오지 못했습니다: {e}")

        st.divider()

    # =========================
    # 6. 연차 요약
    # =========================
    if st.session_state.get("business") == "아이센서":

        st.subheader("👥 연차 요약")

        try:
            vacation_df = load_df("연차관리")
            # ✅ 대시보드 연차 권한 필터
            login_role = str(st.session_state.get("role", "")).strip()
            login_name = str(st.session_state.get("display_name", "")).strip()

            if login_role != "관리자" and "이름" in vacation_df.columns:
                vacation_df = vacation_df[
                    vacation_df["이름"].astype(str).str.strip() == login_name
                ].copy()

            if vacation_df.empty:
                v1, v2, v3 = st.columns(3)
                v1.metric("직원 수", 0)
                v2.metric("잔여 5일 이하", 0)
                v3.metric("잔여 0일 이하", 0)
                st.info("연차 데이터가 없습니다.")
            else:
                for col in ["발생 연차", "사용 연차", "잔여 연차"]:
                    if col in vacation_df.columns:
                        vacation_df[col] = pd.to_numeric(vacation_df[col], errors="coerce").fillna(0)

                low_leave_df = vacation_df[vacation_df["잔여 연차"] <= 5].copy() if "잔여 연차" in vacation_df.columns else pd.DataFrame()
                zero_leave_df = vacation_df[vacation_df["잔여 연차"] <= 0].copy() if "잔여 연차" in vacation_df.columns else pd.DataFrame()

                v1, v2, v3 = st.columns(3)
                v1.metric("직원 수", len(vacation_df))
                v2.metric("잔여 5일 이하", len(low_leave_df))
                v3.metric("잔여 0일 이하", len(zero_leave_df))

                if login_role == "관리자":
                    st.subheader("📊 연차 사용 현황 그래프")

                    leave_chart_df = vacation_df[["이름", "사용 연차", "잔여 연차"]].copy()
                    leave_chart_df = leave_chart_df.set_index("이름")

                    st.bar_chart(leave_chart_df)
                else:
                    st.info("본인 연차 요약만 표시됩니다.")

        except Exception as e:
            st.warning(f"연차 요약을 불러오지 못했습니다: {e}")

    st.divider()

    # =========================
    # 7. 오늘 할 일 / 일정
    # =========================
    st.subheader("✅ 오늘 할 일 / 일정")

    t1, t2 = st.columns(2)

    with t1:
        st.subheader("오늘 할 일")

        if task_df.empty:
            st.info("등록된 할 일이 없습니다.")
        else:
            st.dataframe(
                task_df[["등록일시", "작성자", "사업", "할일"]].tail(10),
                use_container_width=True,
                hide_index=True
            )

    with t2:
        st.subheader("일정 관리")

        if schedule_common_df.empty:
            st.info("등록된 일정이 없습니다.")
        else:
            temp_schedule = schedule_common_df.copy()
            temp_schedule["날짜_dt"] = pd.to_datetime(temp_schedule["날짜"], errors="coerce")

            today_schedule = temp_schedule[temp_schedule["날짜"].astype(str) == today_str]

            if today_schedule.empty:
                st.info("오늘 등록된 일정은 없습니다.")
            else:
                st.dataframe(
                    today_schedule[["등록일시", "작성자", "사업", "일정명", "날짜"]],
                    use_container_width=True,
                    hide_index=True
                )

    st.divider()

    st.success("통합 대시보드 로딩 완료")


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

def is_admin():
    role = str(st.session_state.get("role", "")).strip()
    return role in ["관리자", "admin", "ADMIN"]


def is_admin():
    role = str(st.session_state.get("role", "")).strip()
    return role in ["관리자", "admin", "ADMIN"]

def current_user_name():
    return st.session_state.display_name or st.session_state.username

def logout():
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.session_state.display_name = ""
    st.session_state.department = ""
    st.session_state.position = ""
    st.session_state.user_code = ""
    st.rerun()

# =========================================================
# 메인
# =========================================================
def main():
    init_files()

    st.sidebar.title("메뉴")
    st.sidebar.write(f"로그인: {st.session_state.username}")
    st.sidebar.write(f"이름: {current_user_name()}")
    st.sidebar.write(f"권한: {'관리자' if is_admin() else '담당자'}")

    if st.session_state.get("department"):
        st.sidebar.write(f"부서: {st.session_state.department}")

    if st.session_state.get("position"):
        st.sidebar.write(f"직급: {st.session_state.position}")

    if st.session_state.get("user_code"):
        st.sidebar.write(f"코드: {st.session_state.user_code}")

    selected_business = st.sidebar.selectbox("사업 선택", list(BUSINESS_CONFIG.keys()))
    st.session_state.business = selected_business

    if st.sidebar.button("로그아웃"):
        logout()

    if st.sidebar.button("🏠 홈 ", use_container_width=True):
        st.session_state["sidebar_menu_group"] = "📊 통합"
        st.session_state["sidebar_menu_item"] = "대시보드"
        st.rerun()

    render_header()

    # =====================================================
    # 사이드바 메뉴 그룹화
    # =====================================================
    if st.session_state.business == "아이센서":

        menu_groups = {
            "📊 통합": ["대시보드"],
            "📁 영업관리": [
                "영업현황",
                "가능단지",
                "입찰공고단지",
                "계약단지",
            ],
            "🛠 시공/실사관리": [
                "실사 관리",
                "시공 일정",
            ],
            "👥 인사관리": [
                "연차 관리",
            ],
            "✅ 업무관리": [
                "오늘 할 일",
                "일정 관리",
                "영업 알림",
            ],
        }

        # ✅ 관리자만 유지보수 추가
        if is_admin():
            menu_groups["🔬 유지보수관리"] = [
                "라우터 관리",
                "아이센서 유지보수관리",
                "차량관리"
            ]

        if is_admin():
            menu_groups["⚙ 관리자"] = [
                "데이터 가져오기",
                "관리자 도구",
            ]

    else:
        menu_groups = {
            "📊 통합": [
                "대시보드",
            ],
            "📁 영업관리": [
                "계약접수현황",
            ],
            "🛠 시공/실사관리": [
                "실사 관리",
                "시공 일정",
            ],
            "✅ 업무관리": [
                "오늘 할 일",
                "일정 관리",
                "영업 알림",
            ],
        }

        if is_admin():
            menu_groups["⚙ 관리자"] = [
                "데이터 가져오기",
                "관리자 도구",
            ]

    st.sidebar.markdown("### 업무 구분")

    selected_group = st.sidebar.selectbox(
        "카테고리",
        list(menu_groups.keys()),
        key="sidebar_menu_group"
    )

    menu = st.sidebar.radio(
        "메뉴 선택",
        menu_groups[selected_group],
        key="sidebar_menu_item"
    )

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

    elif menu == "연차 관리":
        vacation_page()

    elif menu == "시공 일정":
        schedule_page()

    elif menu == "실사 관리":
        inspection_page()

    elif menu == "아이센서 유지보수관리":
        maintenance_page()

    elif menu == "차량관리":
        vehicle_page()            

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
