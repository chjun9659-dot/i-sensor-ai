import os
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="윤우 영업관리 시스템", layout="wide")

# =========================================================
# 기본 경로 / 파일 설정
# =========================================================
DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)

FILE_MAP = {
    "영업현황": os.path.join(DATA_DIR, "sales_status.csv"),
    "가능단지": os.path.join(DATA_DIR, "possible_sites.csv"),
    "입찰공고": os.path.join(DATA_DIR, "bid_sites.csv"),
    "계약단지": os.path.join(DATA_DIR, "contract_sites.csv"),
}

# =========================================================
# 로그인 사용자 (임시)
# =========================================================
USERS = {
    "admin": {"pw": "1234", "role": "admin", "name": "관리자"},
    "user1": {"pw": "1234", "role": "user", "name": "영업1"},
    "user2": {"pw": "1234", "role": "user", "name": "영업2"},
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

# 일정 / 할일
if "tasks" not in st.session_state:
    st.session_state.tasks = []

if "schedule" not in st.session_state:
    st.session_state.schedule = []

# =========================================================
# 공통 함수
# =========================================================
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    cols = []
    for c in df.columns:
        c = str(c).replace("\n", " ").strip()
        if c.startswith("Unnamed"):
            c = ""
        cols.append(c)
    df.columns = cols
    return df


def remove_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    keep_cols = []
    for col in df.columns:
        if str(col).strip() == "":
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
        if "날짜" in str(col) or "일자" in str(col) or "마감" in str(col) or "공고" in str(col):
            df[col] = df[col].apply(format_date_value)
        else:
            df[col] = df[col].apply(lambda x: x if isinstance(x, (int, float)) else normalize_text(x))
    return df


def save_df(key: str, df: pd.DataFrame):
    path = FILE_MAP[key]
    df.to_csv(path, index=False, encoding="utf-8-sig")


def load_df(key: str) -> pd.DataFrame:
    path = FILE_MAP[key]
    if os.path.exists(path):
        try:
            return pd.read_csv(path, encoding="utf-8-sig").fillna("")
        except Exception:
            return pd.read_csv(path).fillna("")
    return pd.DataFrame()


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


def read_excel_sheet(uploaded_file, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    raw = raw.fillna("")

    if len(raw) >= 2:
        header = list(raw.iloc[1].values)
        header = [str(x).strip() for x in header]
        header = make_unique_columns(header)
        df = raw.iloc[2:].copy()
        df.columns = header
    else:
        df = raw.copy()

    df = preprocess_df(df)
    return df


def import_sensor_excel(uploaded_file):
    target_sheets = {
        "영업현황": "영업현황",
        "가능단지": "가능단지 모음",
        "입찰공고": "입찰공고단지",
        "계약단지": "계약단지",
    }

    results = {}

    for save_key, sheet_name in target_sheets.items():
        try:
            df = read_excel_sheet(uploaded_file, sheet_name)
            save_df(save_key, df)
            results[save_key] = len(df)
        except Exception as e:
            results[save_key] = f"오류: {e}"

    return results


def filtered_df(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    result = df.copy()
    for col, value in filters.items():
        if col in result.columns and value and value != "전체":
            result = result[result[col].astype(str) == str(value)]
    return result


def keyword_filter(df: pd.DataFrame, keyword: str) -> pd.DataFrame:
    if not keyword:
        return df
    keyword = keyword.strip()
    if not keyword:
        return df

    mask = df.astype(str).apply(
        lambda row: row.str.contains(keyword, case=False, na=False).any(),
        axis=1
    )
    return df[mask]


def login():
    st.title("🔐 윤우 영업관리 로그인")

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


def get_best_column(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def get_site_column(df):
    return get_best_column(df, ["아파트 명", "아파트명", "단지명"])


def get_manager_column(df):
    return get_best_column(df, ["담당자", "영업담당"])


def get_code_column(df):
    return get_best_column(df, ["관리코드", "관리 코드"])


def can_edit_row(df, idx):
    if is_admin():
        return True

    manager_col = get_manager_column(df)
    if manager_col is None:
        return True

    assigned = str(df.loc[idx, manager_col]).strip()
    me = current_user_name().strip()

    if assigned == "" or assigned == "미지정":
        return True

    return assigned == me


def duplicate_check_message(df, code_value, site_value):
    code_col = get_code_column(df)
    site_col = get_site_column(df)
    manager_col = get_manager_column(df)

    if code_col and code_value.strip():
        dup = df[df[code_col].astype(str).str.strip() == code_value.strip()]
        if not dup.empty:
            담당 = ""
            if manager_col:
                담당 = str(dup.iloc[0][manager_col]).strip()
            return f"이미 존재하는 관리코드입니다. 현재 담당자: {담당 or '미지정'}"

    if site_col and site_value.strip():
        dup = df[df[site_col].astype(str).str.strip() == site_value.strip()]
        if not dup.empty:
            담당 = ""
            if manager_col:
                담당 = str(dup.iloc[0][manager_col]).strip()
            return f"이미 등록된 단지입니다. 현재 담당자: {담당 or '미지정'}"

    return ""


def style_status_value(val):
    s = str(val).strip()
    if s in ["계약", "완료", "시공완료", "유찰", "낙찰", "진행완료"]:
        return "background-color: #dff0d8; color: #1b5e20; font-weight: bold;"
    if s in ["진행중", "상담중", "검토중", "운영사 변경중", "필", "상"]:
        return "background-color: #fff4cc; color: #7a5c00; font-weight: bold;"
    if s in ["부결", "실패", "보류", "미진행", "하", "불가"]:
        return "background-color: #f8d7da; color: #8a1f2d; font-weight: bold;"
    return ""


def styled_dataframe(df: pd.DataFrame):
    target_cols = [c for c in df.columns if c in ["진행여부", "결과", "시공여부", "확인", "낙찰여부", "단지 반응도", "가능성"]]

    if not target_cols:
        st.dataframe(df, use_container_width=True, height=500)
        return

    styled = df.style
    for col in target_cols:
        styled = styled.map(style_status_value, subset=[col])

    st.dataframe(styled, use_container_width=True, height=500)


def to_excel_bytes(df_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output.getvalue()


def download_section(title, df, filename_prefix):
    st.subheader(title)
    if df.empty:
        st.info("다운로드할 데이터가 없습니다.")
        return

    csv_data = df.to_csv(index=False, encoding="utf-8-sig")
    st.download_button(
        label="CSV 다운로드",
        data=csv_data,
        file_name=f"{filename_prefix}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
    )

    excel_data = to_excel_bytes({title: df})
    st.download_button(
        label="엑셀 다운로드",
        data=excel_data,
        file_name=f"{filename_prefix}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =========================================================
# 화면 함수
# =========================================================
def page_import():
    st.title("📥 데이터 가져오기")
    st.write("`온도 감지센서.xlsx` 파일을 업로드하면 영업현황 / 가능단지 / 입찰공고 / 계약단지 데이터를 프로그램으로 가져옵니다.")

    uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        st.info("업로드한 파일에서 데이터를 읽을 준비가 되었습니다.")

        if st.button("엑셀 데이터 가져오기"):
            results = import_sensor_excel(uploaded_file)
            st.success("가져오기가 완료되었습니다.")
            st.write("가져오기 결과")
            st.json(results)

    st.divider()

    st.subheader("현재 저장된 데이터 현황")
    for key in FILE_MAP.keys():
        df = load_df(key)
        st.write(f"- {key}: {len(df)}건")

    st.divider()
    st.subheader("전체 데이터 백업 다운로드")

    export_dict = {}
    for key in FILE_MAP.keys():
        export_dict[key] = load_df(key)

    if any(len(df) > 0 for df in export_dict.values()):
        all_excel = to_excel_bytes(export_dict)
        st.download_button(
            label="전체 데이터 엑셀 백업 다운로드",
            data=all_excel,
            file_name=f"윤우영업관리_전체백업_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def page_dashboard():
    st.title("📊 대시보드")

    sales_df = load_df("영업현황")
    possible_df = load_df("가능단지")
    bid_df = load_df("입찰공고")
    contract_df = load_df("계약단지")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("영업현황", len(sales_df))
    c2.metric("가능단지", len(possible_df))
    c3.metric("입찰공고", len(bid_df))
    c4.metric("계약단지", len(contract_df))

    st.divider()

    if not sales_df.empty and "담당자" in sales_df.columns:
        st.subheader("영업현황 담당자별 건수")
        count_df = sales_df["담당자"].astype(str).replace("", "미지정").value_counts().reset_index()
        count_df.columns = ["담당자", "건수"]
        st.dataframe(count_df, use_container_width=True)

    if not possible_df.empty and "결과" in possible_df.columns:
        st.subheader("가능단지 결과 현황")
        result_df = possible_df["결과"].astype(str).replace("", "미입력").value_counts().reset_index()
        result_df.columns = ["결과", "건수"]
        st.dataframe(result_df, use_container_width=True)

    if not contract_df.empty and "영업담당" in contract_df.columns:
        st.subheader("계약단지 담당자별 건수")
        contract_count = contract_df["영업담당"].astype(str).replace("", "미지정").value_counts().reset_index()
        contract_count.columns = ["영업담당", "건수"]
        st.dataframe(contract_count, use_container_width=True)

    st.divider()
    st.subheader("📝 오늘 할 일")
    if len(st.session_state.tasks) == 0:
        st.info("등록된 오늘 할 일이 없습니다.")
    else:
        for task in st.session_state.tasks:
            st.write(f"- {task}")

    st.divider()
    st.subheader("📅 최근 7일 일정")
    if len(st.session_state.schedule) == 0:
        st.info("등록된 일정이 없습니다.")
    else:
        sch_df = pd.DataFrame(st.session_state.schedule)
        sch_df["date"] = pd.to_datetime(sch_df["date"], errors="coerce")
        today = pd.to_datetime(datetime.today().date())
        week_df = sch_df[
            (sch_df["date"] >= today) &
            (sch_df["date"] <= today + pd.Timedelta(days=7))
        ].sort_values("date")

        if week_df.empty:
            st.info("최근 7일 일정이 없습니다.")
        else:
            st.dataframe(week_df, use_container_width=True)


def page_sales_status():
    st.title("📞 영업현황")

    df = load_df("영업현황")
    if df.empty:
        st.warning("영업현황 데이터가 없습니다. 먼저 '데이터 가져오기'에서 엑셀을 불러와 주세요.")
        return

    col1, col2, col3, col4 = st.columns(4)

    담당자 = "전체"
    지역 = "전체"
    상품 = "전체"
    진행여부 = "전체"

    if "담당자" in df.columns:
        담당자 = col1.selectbox("담당자", ["전체"] + sorted([x for x in df["담당자"].astype(str).unique() if x != ""]))
    if "지역" in df.columns:
        지역 = col2.selectbox("지역", ["전체"] + sorted([x for x in df["지역"].astype(str).unique() if x != ""]))
    if "상품" in df.columns:
        상품 = col3.selectbox("상품", ["전체"] + sorted([x for x in df["상품"].astype(str).unique() if x != ""]))
    if "진행여부" in df.columns:
        진행여부 = col4.selectbox("진행여부", ["전체"] + sorted([x for x in df["진행여부"].astype(str).unique() if x != ""]))

    keyword = st.text_input("검색", placeholder="아파트명, 관리코드, 주소, 내용 등 검색")

    df2 = filtered_df(
        df,
        {
            "담당자": 담당자,
            "지역": 지역,
            "상품": 상품,
            "진행여부": 진행여부,
        },
    )
    df2 = keyword_filter(df2, keyword)

    st.write(f"조회 건수: {len(df2)}건")
    styled_dataframe(df2)

    download_section("영업현황_필터결과", df2, "영업현황")

    st.divider()
    st.subheader("진행여부 / 확인 수정")

    site_col = get_site_column(df)
    if not site_col:
        st.info("단지명 컬럼을 찾지 못했습니다.")
        return

    selectable_df = df.copy()
    if not is_admin():
        manager_col = get_manager_column(df)
        if manager_col:
            me = current_user_name().strip()
            selectable_df = df[
                (df[manager_col].astype(str).str.strip() == me) |
                (df[manager_col].astype(str).str.strip() == "") |
                (df[manager_col].astype(str).str.strip() == "미지정")
            ]

    if selectable_df.empty:
        st.warning("수정 가능한 데이터가 없습니다.")
        return

    selected_site = st.selectbox("대상 단지 선택", selectable_df[site_col].astype(str).tolist())

    target_idx_list = df[df[site_col].astype(str) == str(selected_site)].index
    if len(target_idx_list) == 0:
        st.warning("선택한 단지를 찾을 수 없습니다.")
        return

    row_index = target_idx_list[0]

    if not can_edit_row(df, row_index):
        st.error("이 단지는 현재 본인 담당이 아니므로 수정할 수 없습니다.")
        return

    c1, c2 = st.columns(2)
    new_status = c1.text_input("새 진행여부")
    new_check = c2.text_input("새 확인 내용")

    if st.button("영업현황 저장"):
        if "진행여부" in df.columns and new_status.strip():
            df.loc[row_index, "진행여부"] = new_status.strip()
        if "확인" in df.columns and new_check.strip():
            df.loc[row_index, "확인"] = new_check.strip()

        save_df("영업현황", df)
        st.success("수정되었습니다.")
        st.rerun()

    st.divider()
    st.subheader("담당자 변경")
    if is_admin():
        manager_col = get_manager_column(df)
        if manager_col:
            selected_site_admin = st.selectbox(
                "담당자 변경 대상",
                df[site_col].astype(str).tolist(),
                key="sales_manager_change_target"
            )
            new_manager = st.text_input("새 담당자", key="sales_new_manager")

            if st.button("담당자 변경 저장"):
                idx = df[df[site_col].astype(str) == str(selected_site_admin)].index
                if len(idx) > 0 and new_manager.strip():
                    df.loc[idx[0], manager_col] = new_manager.strip()
                    save_df("영업현황", df)
                    st.success("담당자가 변경되었습니다.")
                    st.rerun()
    else:
        st.info("담당자 변경은 관리자만 가능합니다.")


def page_possible_sites():
    st.title("✅ 가능단지")

    df = load_df("가능단지")
    if df.empty:
        st.warning("가능단지 데이터가 없습니다. 먼저 '데이터 가져오기'에서 엑셀을 불러와 주세요.")
        return

    col1, col2, col3, col4 = st.columns(4)

    담당자 = "전체"
    지역 = "전체"
    상품 = "전체"
    결과 = "전체"

    if "영업담당" in df.columns:
        담당자 = col1.selectbox("영업담당", ["전체"] + sorted([x for x in df["영업담당"].astype(str).unique() if x != ""]))
    if "지역" in df.columns:
        지역 = col2.selectbox("지역", ["전체"] + sorted([x for x in df["지역"].astype(str).unique() if x != ""]))
    if "상품" in df.columns:
        상품 = col3.selectbox("상품", ["전체"] + sorted([x for x in df["상품"].astype(str).unique() if x != ""]))
    if "결과" in df.columns:
        결과 = col4.selectbox("결과", ["전체"] + sorted([x for x in df["결과"].astype(str).unique() if x != ""]))

    keyword = st.text_input("검색", placeholder="아파트명, 관리코드, 진행사항, 비고 등 검색", key="possible_search")

    df2 = filtered_df(
        df,
        {
            "영업담당": 담당자,
            "지역": 지역,
            "상품": 상품,
            "결과": 결과,
        },
    )
    df2 = keyword_filter(df2, keyword)

    st.write(f"조회 건수: {len(df2)}건")
    styled_dataframe(df2)

    download_section("가능단지_필터결과", df2, "가능단지")

    st.divider()
    st.subheader("결과 수정")

    site_col = get_best_column(df, ["아파트명", "아파트 명", "단지명"])
    manager_col = get_manager_column(df)

    if site_col:
        selectable_df = df.copy()
        if not is_admin() and manager_col:
            me = current_user_name().strip()
            selectable_df = df[
                (df[manager_col].astype(str).str.strip() == me) |
                (df[manager_col].astype(str).str.strip() == "") |
                (df[manager_col].astype(str).str.strip() == "미지정")
            ]

        if selectable_df.empty:
            st.warning("수정 가능한 데이터가 없습니다.")
            return

        selected_site = st.selectbox("단지 선택", selectable_df[site_col].astype(str).tolist())
        new_result = st.text_input("새 결과값 (예: 계약, 부결, 진행중)", key="possible_result")

        idx = df[df[site_col].astype(str) == str(selected_site)].index
        if len(idx) > 0 and not can_edit_row(df, idx[0]):
            st.error("이 단지는 현재 본인 담당이 아니므로 수정할 수 없습니다.")
            return

        if st.button("가능단지 저장"):
            if len(idx) > 0:
                row_index = idx[0]
                if "결과" in df.columns and new_result.strip():
                    df.loc[row_index, "결과"] = new_result.strip()
                    save_df("가능단지", df)
                    st.success("수정되었습니다.")
                    st.rerun()


def page_bid_sites():
    st.title("📝 입찰공고단지")

    df = load_df("입찰공고")
    if df.empty:
        st.warning("입찰공고 데이터가 없습니다. 먼저 '데이터 가져오기'에서 엑셀을 불러와 주세요.")
        return

    col1, col2, col3, col4 = st.columns(4)

    담당자 = "전체"
    지역 = "전체"
    상품 = "전체"
    판매형태 = "전체"

    if "영업담당" in df.columns:
        담당자 = col1.selectbox("영업담당", ["전체"] + sorted([x for x in df["영업담당"].astype(str).unique() if x != ""]))
    if "지역" in df.columns:
        지역 = col2.selectbox("지역", ["전체"] + sorted([x for x in df["지역"].astype(str).unique() if x != ""]))
    if "상품" in df.columns:
        상품 = col3.selectbox("상품", ["전체"] + sorted([x for x in df["상품"].astype(str).unique() if x != ""]))
    if "판매형태" in df.columns:
        판매형태 = col4.selectbox("판매형태", ["전체"] + sorted([x for x in df["판매형태"].astype(str).unique() if x != ""]))

    keyword = st.text_input("검색", placeholder="아파트명, 관리코드, 특이사항 등 검색", key="bid_search")

    df2 = filtered_df(
        df,
        {
            "영업담당": 담당자,
            "지역": 지역,
            "상품": 상품,
            "판매형태": 판매형태,
        },
    )
    df2 = keyword_filter(df2, keyword)

    st.write(f"조회 건수: {len(df2)}건")
    styled_dataframe(df2)

    download_section("입찰공고_필터결과", df2, "입찰공고")


def page_contract_sites():
    st.title("📦 계약단지")

    df = load_df("계약단지")
    if df.empty:
        st.warning("계약단지 데이터가 없습니다. 먼저 '데이터 가져오기'에서 엑셀을 불러와 주세요.")
        return

    col1, col2, col3, col4 = st.columns(4)

    담당자 = "전체"
    지역 = "전체"
    상품 = "전체"
    시공여부 = "전체"

    if "영업담당" in df.columns:
        담당자 = col1.selectbox("영업담당", ["전체"] + sorted([x for x in df["영업담당"].astype(str).unique() if x != ""]))
    if "지역" in df.columns:
        지역 = col2.selectbox("지역", ["전체"] + sorted([x for x in df["지역"].astype(str).unique() if x != ""]))
    if "상품" in df.columns:
        상품 = col3.selectbox("상품", ["전체"] + sorted([x for x in df["상품"].astype(str).unique() if x != ""]))
    if "시공여부" in df.columns:
        시공여부 = col4.selectbox("시공여부", ["전체"] + sorted([x for x in df["시공여부"].astype(str).unique() if x != ""]))

    keyword = st.text_input("검색", placeholder="아파트명, 관리코드, 주소 등 검색", key="contract_search")

    df2 = filtered_df(
        df,
        {
            "영업담당": 담당자,
            "지역": 지역,
            "상품": 상품,
            "시공여부": 시공여부,
        },
    )
    df2 = keyword_filter(df2, keyword)

    st.write(f"조회 건수: {len(df2)}건")
    styled_dataframe(df2)

    download_section("계약단지_필터결과", df2, "계약단지")

    st.divider()
    st.subheader("시공여부 / 세금계산서 발행 수정")

    site_col = get_best_column(df, ["아파트명", "아파트 명", "단지명"])
    manager_col = get_manager_column(df)

    if site_col:
        selectable_df = df.copy()
        if not is_admin() and manager_col:
            me = current_user_name().strip()
            selectable_df = df[
                (df[manager_col].astype(str).str.strip() == me) |
                (df[manager_col].astype(str).str.strip() == "") |
                (df[manager_col].astype(str).str.strip() == "미지정")
            ]

        if selectable_df.empty:
            st.warning("수정 가능한 데이터가 없습니다.")
            return

        selected_site = st.selectbox("단지 선택", selectable_df[site_col].astype(str).tolist(), key="contract_site")
        target_idx = df[df[site_col].astype(str) == str(selected_site)].index
        if len(target_idx) == 0:
            return

        row_index = target_idx[0]
        if not can_edit_row(df, row_index):
            st.error("이 단지는 현재 본인 담당이 아니므로 수정할 수 없습니다.")
            return

        col_a, col_b = st.columns(2)
        new_build = col_a.text_input("새 시공여부", key="new_build")
        new_tax = col_b.text_input("새 세금계산서 발행", key="new_tax")

        if st.button("계약단지 저장"):
            if "시공여부" in df.columns and new_build.strip():
                df.loc[row_index, "시공여부"] = new_build.strip()
            if "세금계산서 발행" in df.columns and new_tax.strip():
                df.loc[row_index, "세금계산서 발행"] = new_tax.strip()

            save_df("계약단지", df)
            st.success("수정되었습니다.")
            st.rerun()

    st.divider()
    st.subheader("계약단지 담당자 변경")
    if is_admin():
        if site_col and manager_col:
            selected_site_admin = st.selectbox(
                "담당자 변경 대상",
                df[site_col].astype(str).tolist(),
                key="contract_manager_change_target"
            )
            new_manager = st.text_input("새 담당자", key="contract_new_manager")

            if st.button("계약단지 담당자 변경 저장"):
                idx = df[df[site_col].astype(str) == str(selected_site_admin)].index
                if len(idx) > 0 and new_manager.strip():
                    df.loc[idx[0], manager_col] = new_manager.strip()
                    save_df("계약단지", df)
                    st.success("담당자가 변경되었습니다.")
                    st.rerun()
    else:
        st.info("담당자 변경은 관리자만 가능합니다.")


def page_new_sales_entry():
    st.title("➕ 영업현황 신규 등록")
    st.write("새 영업건을 직접 추가할 수 있습니다.")

    df = load_df("영업현황")

    관리코드 = st.text_input("관리코드")
    아파트명 = st.text_input("아파트 명")
    대표전화 = st.text_input("대표전화")
    지역 = st.text_input("지역")
    주소 = st.text_input("주소")
    상품 = st.selectbox("상품", ["", "화재예방시스템", "충전기", "번들(충+화)"])
    영업형태 = st.text_input("영업형태")
    담당자 = st.text_input("담당자", value=current_user_name())
    수량 = st.text_input("수량")
    날짜 = st.text_input("날짜", value=datetime.now().strftime("%Y-%m-%d"))
    티엠결과 = st.text_area("티엠결과")
    가능성 = st.selectbox("가능성", ["", "상", "중", "하"])
    방문현황 = st.text_input("방문현황")
    진행여부 = st.text_input("진행여부")
    내용 = st.text_area("내용")
    확인 = st.text_input("확인")
    계약날짜 = st.text_input("계약날짜")

    if not df.empty:
        dup_msg = duplicate_check_message(df, 관리코드, 아파트명)
        if dup_msg:
            st.warning(dup_msg)

    if st.button("신규 등록 저장"):
        if df.empty:
            df = pd.DataFrame(columns=[
                "관리코드", "아파트 명", "대표전화", "지역", "주소", "상품", "영업형태",
                "담당자", "수량", "날짜", "티엠결과", "가능성", "방문현황",
                "진행여부", "내용", "확인", "계약날짜"
            ])

        dup_msg = duplicate_check_message(df, 관리코드, 아파트명)
        if dup_msg:
            st.error(dup_msg)
            return

        new_row = {
            "관리코드": 관리코드.strip(),
            "아파트 명": 아파트명.strip(),
            "대표전화": 대표전화.strip(),
            "지역": 지역.strip(),
            "주소": 주소.strip(),
            "상품": 상품.strip(),
            "영업형태": 영업형태.strip(),
            "담당자": 담당자.strip(),
            "수량": 수량.strip(),
            "날짜": 날짜.strip(),
            "티엠결과": 티엠결과.strip(),
            "가능성": 가능성.strip(),
            "방문현황": 방문현황.strip(),
            "진행여부": 진행여부.strip(),
            "내용": 내용.strip(),
            "확인": 확인.strip(),
            "계약날짜": 계약날짜.strip(),
        }

        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_df("영업현황", df)
        st.success("영업현황에 신규 등록되었습니다.")
        st.rerun()


def page_tasks():
    st.title("📝 오늘 할 일")

    new_task = st.text_input("할 일 입력")

    if st.button("추가"):
        if new_task.strip():
            st.session_state.tasks.append(new_task.strip())
            st.success("추가 완료")
            st.rerun()

    st.write("---")

    if len(st.session_state.tasks) == 0:
        st.info("등록된 할 일이 없습니다.")
    else:
        for i, task in enumerate(st.session_state.tasks):
            col1, col2 = st.columns([8, 1])
            col1.write(task)
            if col2.button("삭제", key=f"task_{i}"):
                st.session_state.tasks.pop(i)
                st.rerun()


def page_schedule():
    st.title("📅 일정 관리")

    title = st.text_input("일정 제목")
    selected_date = st.date_input("날짜")

    if st.button("일정 추가"):
        if title.strip():
            st.session_state.schedule.append({
                "title": title.strip(),
                "date": str(selected_date)
            })
            st.success("일정 추가 완료")
            st.rerun()

    st.write("---")

    if len(st.session_state.schedule) == 0:
        st.info("등록된 일정이 없습니다.")
    else:
        df = pd.DataFrame(st.session_state.schedule)
        st.dataframe(df, use_container_width=True)

        st.subheader("일정 삭제")
        option_list = [
            f"{i + 1}. {item['date']} | {item['title']}"
            for i, item in enumerate(st.session_state.schedule)
        ]
        selected_item = st.selectbox("삭제할 일정 선택", option_list)

        if st.button("선택 일정 삭제"):
            idx = option_list.index(selected_item)
            st.session_state.schedule.pop(idx)
            st.success("삭제 완료")
            st.rerun()


def page_week():
    st.title("📈 주간 일정")

    if not st.session_state.schedule:
        st.info("일정 없음")
        return

    df = pd.DataFrame(st.session_state.schedule)
    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    today = pd.to_datetime(datetime.today().date())

    week_df = df[
        (df["date"] >= today) &
        (df["date"] <= today + pd.Timedelta(days=7))
    ]

    if week_df.empty:
        st.info("최근 7일 일정이 없습니다.")
    else:
        st.dataframe(week_df.sort_values("date"), use_container_width=True)


def page_admin_tools():
    st.title("🛠 관리자 도구")

    if not is_admin():
        st.error("이 메뉴는 관리자만 사용할 수 있습니다.")
        return

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


# =========================================================
# 메인
# =========================================================
def main():
    st.sidebar.title("메뉴")
    st.sidebar.write(f"로그인: {st.session_state.username}")
    st.sidebar.write(f"권한: {st.session_state.role}")

    menus = [
        "대시보드",
        "데이터 가져오기",
        "영업현황",
        "가능단지",
        "입찰공고단지",
        "계약단지",
        "영업현황 신규 등록",
        "오늘 할 일",
        "일정 관리",
        "주간 일정",
    ]

    if is_admin():
        menus.append("관리자 도구")

    menu = st.sidebar.radio("선택", menus)

    if st.sidebar.button("로그아웃"):
        logout()

    if menu == "대시보드":
        page_dashboard()
    elif menu == "데이터 가져오기":
        page_import()
    elif menu == "영업현황":
        page_sales_status()
    elif menu == "가능단지":
        page_possible_sites()
    elif menu == "입찰공고단지":
        page_bid_sites()
    elif menu == "계약단지":
        page_contract_sites()
    elif menu == "영업현황 신규 등록":
        page_new_sales_entry()
    elif menu == "오늘 할 일":
        page_tasks()
    elif menu == "일정 관리":
        page_schedule()
    elif menu == "주간 일정":
        page_week()
    elif menu == "관리자 도구":
        page_admin_tools()


if not st.session_state.logged_in:
    login()
else:
    main()
