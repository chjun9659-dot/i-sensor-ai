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
            "오늘 할 일",
            "일정 관리",
            "영업 알림",
        ],
    },
    "전기차 충전기": {
        "sheets": {
            # 여기에 충전기 '계약서 접수현황' 시트 URL만 넣으시면 됩니다.
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
        if not df[col].isna().all():
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
    df = remove_empty_columns(df)
    df = df.fillna("")

    for col in df.columns:
        if any(key in str(col) for key in ["날짜", "일자", "마감", "공고", "접수", "실사"]):
            df[col] = df[col].apply(format_date_value)
        else:
            df[col] = df[col].apply(lambda x: x if isinstance(x, (int, float)) else normalize_text(x))
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
def convert_google_sheet_url_to_csv(url: str) -> str:
    sheet_match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    gid_match = re.search(r"gid=(\d+)", url)

    if not sheet_match:
        raise ValueError("구글 스프레드시트 문서 ID를 찾을 수 없습니다.")

    sheet_id = sheet_match.group(1)
    gid = gid_match.group(1) if gid_match else "0"
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}"


@st.cache_data(ttl=60)
def load_google_sheet_data(business_name: str, sheet_name: str, url: str) -> pd.DataFrame:
    if not url or "여기에_" in url:
        return pd.DataFrame()

    csv_url = convert_google_sheet_url_to_csv(url)

    try:
        df = pd.read_csv(csv_url, header=1)
    except Exception:
        df = pd.read_csv(csv_url, header=0)

    return preprocess_df(df)


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
    COMMON_FILE_MAP[key]
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
    return get_best_column(df, ["아파트 명", "아파트명", "단지명", "현장명"])


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
    if s in ["진행중", "상담중", "검토중", "운영사 변경중", "필", "상", "접수중"]:
        return "background-color: #fff4cc; color: #7a5c00; font-weight: bold;"
    if s in ["부결", "실패", "보류", "미진행", "하", "불가", "미접수"]:
        return "background-color: #f8d7da; color: #8a1f2d; font-weight: bold;"
    return ""


def convert_number_display(value):
    if pd.isna(value):
        return ""
    try:
        # 문자열 숫자도 처리
        num = float(value)
        # 정수처럼 보이면 정수로 표시
        if num.is_integer():
            return int(num)
        # 소수는 너무 길지 않게
        return round(num, 2)
    except Exception:
        return value


def prepare_display_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()

    # 완전히 비어 있는 컬럼 제거
    keep_cols = []
    for col in df.columns:
        col_name = str(col).strip()
        if col_name.startswith("빈컬럼"):
            # 값이 하나라도 있으면 유지, 전부 비었으면 제거
            if not df[col].astype(str).replace("", pd.NA).isna().all():
                keep_cols.append(col)
        else:
            keep_cols.append(col)

    df = df[keep_cols].copy()

    # 숫자 표시 정리
    for col in df.columns:
        df[col] = df[col].apply(convert_number_display)

    return df


def styled_dataframe(df: pd.DataFrame):
    df = prepare_display_df(df)

    target_cols = [
        c for c in df.columns
        if c in ["진행여부", "결과", "시공여부", "확인", "낙찰여부", "단지 반응도", "가능성", "구분", "상태"]
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
            site_col = get_best_column(df, ["아파트명", "아파트 명", "단지명"])
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
