import os
import re
import io
from datetime import datetime

import pandas as pd
import streamlit as st


def clean_text(value):
    if pd.isna(value):
        return ""
    if isinstance(value, str):
        return re.sub(r"[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F]", "", value)
    return value


st.set_page_config(page_title="윤우 영업관리 시스템", layout="wide")

# =========================================================
# 구글 스프레드시트 자동 연동 설정
# =========================================================
GOOGLESHEET_URLS = {
    "영업현황": "https://docs.google.com/spreadsheets/d/1CWTvHC1r6i5wjcZoFJa5kAm-6T0SKao1EdjI8jwVQG8/edit?gid=167508641#gid=167508641",
    "가능단지": "https://docs.google.com/spreadsheets/d/1CWTvHC1r6i5wjcZoFJa5kAm-6T0SKao1EdjI8jwVQG8/edit?gid=1108943027#gid=1108943027",
    "입찰공고": "https://docs.google.com/spreadsheets/d/1CWTvHC1r6i5wjcZoFJa5kAm-6T0SKao1EdjI8jwVQG8/edit?gid=243967548#gid=243967548",
    "계약단지": "https://docs.google.com/spreadsheets/d/1CWTvHC1r6i5wjcZoFJa5kAm-6T0SKao1EdjI8jwVQG8/edit?gid=2071693391#gid=2071693391",
}

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
    "할일": os.path.join(DATA_DIR, "tasks.csv"),
    "일정": os.path.join(DATA_DIR, "schedule.csv"),
    "세금알림": os.path.join(DATA_DIR, "tax_alerts.csv"),
    "입대의알림": os.path.join(DATA_DIR, "meeting_alerts.csv"),
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


# =========================================================
# 구글 시트 관련 함수
# =========================================================
def convert_google_sheet_url_to_csv(url: str) -> str:
    """
    구글 스프레드시트 edit 링크를 csv 읽기용 링크로 변환
    """
    sheet_match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    gid_match = re.search(r"gid=(\d+)", url)

    if not sheet_match:
        raise ValueError("구글 스프레드시트 문서 ID를 찾을 수 없습니다.")

    sheet_id = sheet_match.group(1)
    gid = gid_match.group(1) if gid_match else "0"

    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}"


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
        if "날짜" in str(col) or "일자" in str(col) or "마감" in str(col) or "공고" in str(col):
            df[col] = df[col].apply(format_date_value)
        else:
            df[col] = df[col].apply(lambda x: x if isinstance(x, (int, float)) else normalize_text(x))
    return df


@st.cache_data(ttl=60)
def load_google_sheet_data(sheet_name: str) -> pd.DataFrame:
    if sheet_name not in GOOGLESHEET_URLS:
        return pd.DataFrame()

    csv_url = convert_google_sheet_url_to_csv(GOOGLESHEET_URLS[sheet_name])

    try:
        df = pd.read_csv(csv_url, header=1)
    except Exception:
        df = pd.read_csv(csv_url, header=0)

    df = preprocess_df(df)
    return df


# =========================================================
# 공통 파일 함수
# =========================================================
def save_df(key: str, df: pd.DataFrame):
    path = FILE_MAP[key]
    df.to_csv(path, index=False, encoding="utf-8-sig")


def load_local_df(key: str) -> pd.DataFrame:
    path = FILE_MAP[key]
    if os.path.exists(path):
        try:
            return pd.read_csv(path, encoding="utf-8-sig").fillna("")
        except Exception:
            return pd.read_csv(path).fillna("")
    return pd.DataFrame()


def load_df(key: str) -> pd.DataFrame:
    """
    4개 메인 데이터는 구글시트 우선 사용
    나머지는 로컬 CSV 사용
    """
    if key in GOOGLESHEET_URLS:
        try:
            df = load_google_sheet_data(key)
            if not df.empty:
                save_df(key, df)
                return df
        except Exception as e:
            st.warning(f"{key} 구글시트 불러오기 실패, 로컬 백업 사용: {e}")

    return load_local_df(key)


def init_tasks_file():
    if not os.path.exists(FILE_MAP["할일"]):
        df = pd.DataFrame(columns=["등록일시", "작성자", "할일"])
        save_df("할일", df)


def init_schedule_file():
    if not os.path.exists(FILE_MAP["일정"]):
        df = pd.DataFrame(columns=["등록일시", "작성자", "일정명", "날짜"])
        save_df("일정", df)


def init_tax_alert_file():
    if not os.path.exists(FILE_MAP["세금알림"]):
        df = pd.DataFrame(columns=["등록일시", "작성자", "단지명", "예정일", "상태", "비고"])
        save_df("세금알림", df)


def init_meeting_alert_file():
    if not os.path.exists(FILE_MAP["입대의알림"]):
        df = pd.DataFrame(columns=["등록일시", "작성자", "단지명", "입대의일자", "상태", "비고"])
        save_df("입대의알림", df)


def load_tasks_df():
    init_tasks_file()
    df = load_local_df("할일")
    needed = ["등록일시", "작성자", "할일"]
    for col in needed:
        if col not in df.columns:
            df[col] = ""
    return df[needed].copy()


def save_tasks_df(df: pd.DataFrame):
    save_df("할일", df[["등록일시", "작성자", "할일"]].copy())


def load_schedule_df():
    init_schedule_file()
    df = load_local_df("일정")
    needed = ["등록일시", "작성자", "일정명", "날짜"]
    for col in needed:
        if col not in df.columns:
            df[col] = ""
    return df[needed].copy()


def save_schedule_df(df: pd.DataFrame):
    save_df("일정", df[["등록일시", "작성자", "일정명", "날짜"]].copy())


def load_tax_alert_df():
    init_tax_alert_file()
    df = load_local_df("세금알림")
    needed = ["등록일시", "작성자", "단지명", "예정일", "상태", "비고"]
    for col in needed:
        if col not in df.columns:
            df[col] = ""
    return df[needed].copy()


def save_tax_alert_df(df: pd.DataFrame):
    save_df("세금알림", df[["등록일시", "작성자", "단지명", "예정일", "상태", "비고"]].copy())


def load_meeting_alert_df():
    init_meeting_alert_file()
    df = load_local_df("입대의알림")
    needed = ["등록일시", "작성자", "단지명", "입대의일자", "상태", "비고"]
    for col in needed:
        if col not in df.columns:
            df[col] = ""
    return df[needed].copy()


def save_meeting_alert_df(df: pd.DataFrame):
    save_df("입대의알림", df[["등록일시", "작성자", "단지명", "입대의일자", "상태", "비고"]].copy())


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
# 엑셀 다운로드 함수 (복구 완료)
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
                safe_sheet_name = clean_text(safe_sheet_name)
                if not safe_sheet_name:
                    safe_sheet_name = f"Sheet{written_count + 1}"
                safe_sheet_name = safe_sheet_name[:31]

                df2.to_excel(writer, index=False, sheet_name=safe_sheet_name)
                written_count += 1

            except Exception as e:
                error_df = pd.DataFrame({
                    "오류": [f"{sheet_name} 시트 저장 실패"],
                    "상세": [str(e)]
                })
                fail_sheet_name = f"오류{written_count + 1}"[:31]
                error_df.to_excel(writer, index=False, sheet_name=fail_sheet_name)
                written_count += 1

        if written_count == 0:
            pd.DataFrame({"안내": ["저장할 데이터가 없습니다."]}).to_excel(
                writer, index=False, sheet_name="안내"
            )

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

        if not excel_data:
            st.error(f"{title} 다운로드 데이터 생성 실패")
            return

        st.download_button(
            label=f"📥 {title} 다운로드",
            data=excel_data,
            file_name=f"{file_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_{file_name}_{title}"
        )

    except Exception as e:
        st.error(f"{title} 다운로드 오류: {e}")


# =========================================================
# 화면 함수
# =========================================================
def page_import():
    st.title("📥 데이터 가져오기")
    st.write("기존 엑셀 가져오기 기능도 유지됩니다.")
    st.write("현재 메인 데이터는 구글 스프레드시트에서 자동 불러옵니다.")

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
        df = load_df(key) if key in GOOGLESHEET_URLS else load_local_df(key)
        st.write(f"- {key}: {len(df)}건")

    st.divider()
    st.subheader("전체 데이터 백업 다운로드")

    export_dict = {}
    for key in FILE_MAP.keys():
        export_dict[key] = load_df(key) if key in GOOGLESHEET_URLS else load_local_df(key)

    if any(len(df) > 0 for df in export_dict.values()):
        all_excel = to_excel_bytes(export_dict)

        if all_excel:
            st.download_button(
                label="전체 데이터 엑셀 백업 다운로드",
                data=all_excel,
                file_name=f"윤우영업관리_전체백업_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_all_backup"
            )
        else:
            st.error("전체 데이터 백업 파일 생성 실패")
    else:
        st.info("백업할 데이터가 없습니다.")

    st.divider()
    st.subheader("구글시트 연결 확인")

    for sheet_name in GOOGLESHEET_URLS.keys():
        try:
            df = load_google_sheet_data(sheet_name)
            st.success(f"{sheet_name}: 연결 성공 ({len(df)}건)")
        except Exception as e:
            st.error(f"{sheet_name}: 연결 실패 - {e}")


def page_dashboard():
    st.title("📊 대시보드")

    sales_df = load_df("영업현황")
    possible_df = load_df("가능단지")
    bid_df = load_df("입찰공고")
    contract_df = load_df("계약단지")
    tasks_df = load_tasks_df()
    schedule_df = load_schedule_df()
    tax_df = load_tax_alert_df()
    meeting_df = load_meeting_alert_df()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("영업현황", len(sales_df))
    c2.metric("가능단지", len(possible_df))
    c3.metric("입찰공고", len(bid_df))
    c4.metric("계약단지", len(contract_df))

    st.divider()

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

    if not sales_df.empty:
        manager_col = get_manager_column(sales_df)
        if manager_col:
            st.subheader("영업현황 담당자별 건수")
            count_df = sales_df[manager_col].astype(str).replace("", "미지정").value_counts().reset_index()
            count_df.columns = ["담당자", "건수"]
            st.dataframe(count_df, use_container_width=True)

    if not possible_df.empty and "결과" in possible_df.columns:
        st.subheader("가능단지 결과 현황")
        result_df = possible_df["결과"].astype(str).replace("", "미입력").value_counts().reset_index()
        result_df.columns = ["결과", "건수"]
        st.dataframe(result_df, use_container_width=True)

    if not contract_df.empty:
        manager_col = get_manager_column(contract_df)
        if manager_col:
            st.subheader("계약단지 담당자별 건수")
            contract_count = contract_df[manager_col].astype(str).replace("", "미지정").value_counts().reset_index()
            contract_count.columns = ["영업담당", "건수"]
            st.dataframe(contract_count, use_container_width=True)

    st.divider()
    st.subheader("📝 오늘 할 일")

    if tasks_df.empty:
        st.info("등록된 할 일이 없습니다.")
    else:
        show_tasks = tasks_df.tail(10).iloc[::-1]
        st.dataframe(show_tasks, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("📅 최근 7일 일정")

    if schedule_df.empty:
        st.info("등록된 일정이 없습니다.")
    else:
        temp_df = schedule_df.copy()
        temp_df["날짜"] = pd.to_datetime(temp_df["날짜"], errors="coerce")
        today = pd.to_datetime(datetime.today().date())
        week_df = temp_df[
            (temp_df["날짜"] >= today) &
            (temp_df["날짜"] <= today + pd.Timedelta(days=7))
        ].sort_values("날짜")

        if week_df.empty:
            st.info("최근 7일 일정이 없습니다.")
        else:
            st.dataframe(week_df, use_container_width=True, hide_index=True)


def page_sales_status():
    st.title("📞 영업현황")

    df = load_df("영업현황")
    if df.empty:
        st.warning("영업현황 데이터가 없습니다.")
        return

    col1, col2, col3, col4 = st.columns(4)

    담당자 = "전체"
    지역 = "전체"
    상품 = "전체"
    진행여부 = "전체"

    manager_col = get_manager_column(df)

    if manager_col:
        담당자 = col1.selectbox("담당자", ["전체"] + sorted([x for x in df[manager_col].astype(str).unique() if x != ""]))
    if "지역" in df.columns:
        지역 = col2.selectbox("지역", ["전체"] + sorted([x for x in df["지역"].astype(str).unique() if x != ""]))
    if "상품" in df.columns:
        상품 = col3.selectbox("상품", ["전체"] + sorted([x for x in df["상품"].astype(str).unique() if x != ""]))
    if "진행여부" in df.columns:
        진행여부 = col4.selectbox("진행여부", ["전체"] + sorted([x for x in df["진행여부"].astype(str).unique() if x != ""]))

    keyword = st.text_input("검색", placeholder="아파트명, 관리코드, 주소, 내용 등 검색")

    filters = {}
    if manager_col:
        filters[manager_col] = 담당자
    if "지역" in df.columns:
        filters["지역"] = 지역
    if "상품" in df.columns:
        filters["상품"] = 상품
    if "진행여부" in df.columns:
        filters["진행여부"] = 진행여부

    df2 = filtered_df(df, filters)
    df2 = keyword_filter(df2, keyword)

    st.write(f"조회 건수: {len(df2)}건")
    styled_dataframe(df2)

    download_section("영업현황_필터결과", df2, "영업현황")


def page_possible_sites():
    st.title("✅ 가능단지")

    df = load_df("가능단지")
    if df.empty:
        st.warning("가능단지 데이터가 없습니다.")
        return

    col1, col2, col3, col4 = st.columns(4)

    담당자 = "전체"
    지역 = "전체"
    상품 = "전체"
    결과 = "전체"

    manager_col = get_manager_column(df)

    if manager_col:
        담당자 = col1.selectbox("영업담당", ["전체"] + sorted([x for x in df[manager_col].astype(str).unique() if x != ""]))
    if "지역" in df.columns:
        지역 = col2.selectbox("지역", ["전체"] + sorted([x for x in df["지역"].astype(str).unique() if x != ""]))
    if "상품" in df.columns:
        상품 = col3.selectbox("상품", ["전체"] + sorted([x for x in df["상품"].astype(str).unique() if x != ""]))
    if "결과" in df.columns:
        결과 = col4.selectbox("결과", ["전체"] + sorted([x for x in df["결과"].astype(str).unique() if x != ""]))

    keyword = st.text_input("검색", placeholder="아파트명, 관리코드, 진행사항, 비고 등 검색", key="possible_search")

    filters = {}
    if manager_col:
        filters[manager_col] = 담당자
    if "지역" in df.columns:
        filters["지역"] = 지역
    if "상품" in df.columns:
        filters["상품"] = 상품
    if "결과" in df.columns:
        filters["결과"] = 결과

    df2 = filtered_df(df, filters)
    df2 = keyword_filter(df2, keyword)

    st.write(f"조회 건수: {len(df2)}건")
    styled_dataframe(df2)

    download_section("가능단지_필터결과", df2, "가능단지")


def page_bid_sites():
    st.title("📝 입찰공고단지")

    df = load_df("입찰공고")
    if df.empty:
        st.warning("입찰공고 데이터가 없습니다.")
        return

    col1, col2, col3, col4 = st.columns(4)

    담당자 = "전체"
    지역 = "전체"
    상품 = "전체"
    판매형태 = "전체"

    manager_col = get_manager_column(df)

    if manager_col:
        담당자 = col1.selectbox("영업담당", ["전체"] + sorted([x for x in df[manager_col].astype(str).unique() if x != ""]))
    if "지역" in df.columns:
        지역 = col2.selectbox("지역", ["전체"] + sorted([x for x in df["지역"].astype(str).unique() if x != ""]))
    if "상품" in df.columns:
        상품 = col3.selectbox("상품", ["전체"] + sorted([x for x in df["상품"].astype(str).unique() if x != ""]))
    if "판매형태" in df.columns:
        판매형태 = col4.selectbox("판매형태", ["전체"] + sorted([x for x in df["판매형태"].astype(str).unique() if x != ""]))

    keyword = st.text_input("검색", placeholder="아파트명, 관리코드, 특이사항 등 검색", key="bid_search")

    filters = {}
    if manager_col:
        filters[manager_col] = 담당자
    if "지역" in df.columns:
        filters["지역"] = 지역
    if "상품" in df.columns:
        filters["상품"] = 상품
    if "판매형태" in df.columns:
        filters["판매형태"] = 판매형태

    df2 = filtered_df(df, filters)
    df2 = keyword_filter(df2, keyword)

    st.write(f"조회 건수: {len(df2)}건")
    styled_dataframe(df2)

    download_section("입찰공고_필터결과", df2, "입찰공고")


def page_contract_sites():
    st.title("📦 계약단지")

    df = load_df("계약단지")
    if df.empty:
        st.warning("계약단지 데이터가 없습니다.")
        return

    col1, col2, col3, col4 = st.columns(4)

    담당자 = "전체"
    지역 = "전체"
    상품 = "전체"
    시공여부 = "전체"

    manager_col = get_manager_column(df)

    if manager_col:
        담당자 = col1.selectbox("영업담당", ["전체"] + sorted([x for x in df[manager_col].astype(str).unique() if x != ""]))
    if "지역" in df.columns:
        지역 = col2.selectbox("지역", ["전체"] + sorted([x for x in df["지역"].astype(str).unique() if x != ""]))
    if "상품" in df.columns:
        상품 = col3.selectbox("상품", ["전체"] + sorted([x for x in df["상품"].astype(str).unique() if x != ""]))
    if "시공여부" in df.columns:
        시공여부 = col4.selectbox("시공여부", ["전체"] + sorted([x for x in df["시공여부"].astype(str).unique() if x != ""]))

    keyword = st.text_input("검색", placeholder="아파트명, 관리코드, 주소 등 검색", key="contract_search")

    filters = {}
    if manager_col:
        filters[manager_col] = 담당자
    if "지역" in df.columns:
        filters["지역"] = 지역
    if "상품" in df.columns:
        filters["상품"] = 상품
    if "시공여부" in df.columns:
        filters["시공여부"] = 시공여부

    df2 = filtered_df(df, filters)
    df2 = keyword_filter(df2, keyword)

    st.write(f"조회 건수: {len(df2)}건")
    styled_dataframe(df2)

    download_section("계약단지_필터결과", df2, "계약단지")


def page_new_sales_entry():
    st.title("➕ 영업현황 신규 등록")
    st.info("현재 메인 영업현황은 구글시트 자동연동 방식입니다.")
    st.write("신규 등록은 구글 스프레드시트에서 직접 입력하시는 것을 권장드립니다.")


def page_tasks():
    st.title("📝 오늘 할 일")

    df = load_tasks_df()

    new_task = st.text_input("할 일 입력")

    if st.button("할 일 추가"):
        if new_task.strip():
            new_row = {
                "등록일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "작성자": current_user_name(),
                "할일": new_task.strip(),
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            save_tasks_df(df)
            st.success("추가 완료")
            st.rerun()

    st.write("---")

    if df.empty:
        st.info("등록된 할 일이 없습니다.")
    else:
        view_df = df.iloc[::-1].reset_index(drop=True)
        st.dataframe(view_df, use_container_width=True, hide_index=True)

        st.subheader("할 일 삭제")
        option_list = [
            f"{idx} | {row['작성자']} | {row['할일']}"
            for idx, row in df.iterrows()
        ]
        selected_item = st.selectbox("삭제할 할 일 선택", option_list)

        if st.button("선택 할 일 삭제"):
            delete_idx = int(selected_item.split(" | ")[0])
            df = df.drop(index=delete_idx).reset_index(drop=True)
            save_tasks_df(df)
            st.success("삭제 완료")
            st.rerun()


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
                "일정명": title.strip(),
                "날짜": str(selected_date),
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            save_schedule_df(df)
            st.success("일정 추가 완료")
            st.rerun()

    st.write("---")

    if df.empty:
        st.info("등록된 일정이 없습니다.")
    else:
        temp_df = df.copy()
        temp_df["날짜정렬"] = pd.to_datetime(temp_df["날짜"], errors="coerce")
        temp_df = temp_df.sort_values(["날짜정렬", "등록일시"], ascending=[True, False]).drop(columns=["날짜정렬"])
        st.dataframe(temp_df, use_container_width=True, hide_index=True)

        st.subheader("일정 삭제")
        option_list = [
            f"{idx} | {row['날짜']} | {row['일정명']}"
            for idx, row in df.iterrows()
        ]
        selected_item = st.selectbox("삭제할 일정 선택", option_list)

        if st.button("선택 일정 삭제"):
            delete_idx = int(selected_item.split(" | ")[0])
            df = df.drop(index=delete_idx).reset_index(drop=True)
            save_schedule_df(df)
            st.success("삭제 완료")
            st.rerun()


def page_week():
    st.title("📈 주간 일정")

    df = load_schedule_df()
    if df.empty:
        st.info("일정 없음")
        return

    df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
    today = pd.to_datetime(datetime.today().date())

    week_df = df[
        (df["날짜"] >= today) &
        (df["날짜"] <= today + pd.Timedelta(days=7))
    ]

    if week_df.empty:
        st.info("최근 7일 일정이 없습니다.")
    else:
        week_df = week_df.sort_values("날짜")
        st.dataframe(week_df, use_container_width=True, hide_index=True)


def page_alerts():
    st.title("🚨 영업 알림")

    tax_df = load_tax_alert_df()
    meeting_df = load_meeting_alert_df()
    schedule_df = load_schedule_df()

    st.subheader("1. 알림 요약")
    c1, c2, c3 = st.columns(3)

    tax_temp = tax_df.copy()
    if not tax_temp.empty:
        tax_temp["예정일_dt"] = pd.to_datetime(tax_temp["예정일"], errors="coerce")
        tax_temp["상태표시"] = tax_temp.apply(lambda r: make_alert_status(r["예정일_dt"], r["상태"]), axis=1)
        tax_pending = tax_temp[tax_temp["상태표시"].isin(["지남", "오늘", "긴급", "임박"])]
    else:
        tax_pending = pd.DataFrame()

    meeting_temp = meeting_df.copy()
    if not meeting_temp.empty:
        meeting_temp["입대의일자_dt"] = pd.to_datetime(meeting_temp["입대의일자"], errors="coerce")
        meeting_temp["상태표시"] = meeting_temp.apply(lambda r: make_alert_status(r["입대의일자_dt"], r["상태"]), axis=1)
        meeting_pending = meeting_temp[meeting_temp["상태표시"].isin(["지남", "오늘", "긴급", "임박"])]
    else:
        meeting_pending = pd.DataFrame()

    schedule_temp = schedule_df.copy()
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
    if tax_df.empty:
        st.info("등록된 세금계산서 알림이 없습니다.")
    else:
        view_tax = tax_df.copy()
        view_tax["예정일_dt"] = pd.to_datetime(view_tax["예정일"], errors="coerce")
        view_tax["D-Day"] = view_tax["예정일_dt"].apply(get_d_day_label)
        view_tax["상태표시"] = view_tax.apply(lambda r: make_alert_status(r["예정일_dt"], r["상태"]), axis=1)
        view_tax = view_tax.sort_values(["예정일_dt", "등록일시"], ascending=[True, False])
        show_alert_table(view_tax[["단지명", "예정일", "D-Day", "상태", "상태표시", "비고", "작성자"]])

        tax_options = [
            f"{idx} | {row['단지명']} | {row['예정일']} | {row['상태']}"
            for idx, row in tax_df.iterrows()
        ]
        selected_tax = st.selectbox("상태 변경할 세금계산서 알림", tax_options, key="tax_select")
        new_tax_status = st.selectbox("변경 상태", ["예정", "완료"], key="tax_status_change")

        col_a, col_b = st.columns(2)
        if col_a.button("세금계산서 상태 저장"):
            idx = int(selected_tax.split(" | ")[0])
            tax_df.loc[idx, "상태"] = new_tax_status
            save_tax_alert_df(tax_df)
            st.success("세금계산서 상태가 변경되었습니다.")
            st.rerun()

        if col_b.button("세금계산서 알림 삭제"):
            idx = int(selected_tax.split(" | ")[0])
            tax_df = tax_df.drop(index=idx).reset_index(drop=True)
            save_tax_alert_df(tax_df)
            st.success("삭제되었습니다.")
            st.rerun()

    st.divider()

    st.subheader("3. 입대의 일정 알림 등록")
    col1, col2, col3 = st.columns(3)
    meeting_site = col1.text_input("단지명", key="meeting_site")
    meeting_date = col2.date_input("입대의 날짜", key="meeting_date")
    meeting_note = col3.text_input("비고", key="meeting_note")

    if st.button("입대의 알림 추가"):
        if meeting_site.strip():
            new_row = {
                "등록일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "작성자": current_user_name(),
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
    if meeting_df.empty:
        st.info("등록된 입대의 알림이 없습니다.")
    else:
        view_meeting = meeting_df.copy()
        view_meeting["입대의일자_dt"] = pd.to_datetime(view_meeting["입대의일자"], errors="coerce")
        view_meeting["D-Day"] = view_meeting["입대의일자_dt"].apply(get_d_day_label)
        view_meeting["상태표시"] = view_meeting.apply(lambda r: make_alert_status(r["입대의일자_dt"], r["상태"]), axis=1)
        view_meeting = view_meeting.sort_values(["입대의일자_dt", "등록일시"], ascending=[True, False])
        show_alert_table(view_meeting[["단지명", "입대의일자", "D-Day", "상태", "상태표시", "비고", "작성자"]])

        meeting_options = [
            f"{idx} | {row['단지명']} | {row['입대의일자']} | {row['상태']}"
            for idx, row in meeting_df.iterrows()
        ]
        selected_meeting = st.selectbox("상태 변경할 입대의 알림", meeting_options, key="meeting_select")
        new_meeting_status = st.selectbox("변경 상태", ["예정", "완료"], key="meeting_status_change")

        col_a, col_b = st.columns(2)
        if col_a.button("입대의 상태 저장"):
            idx = int(selected_meeting.split(" | ")[0])
            meeting_df.loc[idx, "상태"] = new_meeting_status
            save_meeting_alert_df(meeting_df)
            st.success("입대의 상태가 변경되었습니다.")
            st.rerun()

        if col_b.button("입대의 알림 삭제"):
            idx = int(selected_meeting.split(" | ")[0])
            meeting_df = meeting_df.drop(index=idx).reset_index(drop=True)
            save_meeting_alert_df(meeting_df)
            st.success("삭제되었습니다.")
            st.rerun()

    st.divider()

    st.subheader("4. 일반 일정 임박 알림")
    if schedule_df.empty:
        st.info("등록된 일정이 없습니다.")
    else:
        view_schedule = schedule_df.copy()
        view_schedule["날짜_dt"] = pd.to_datetime(view_schedule["날짜"], errors="coerce")
        view_schedule["D-Day"] = view_schedule["날짜_dt"].apply(get_d_day_label)
        view_schedule["상태표시"] = view_schedule["날짜_dt"].apply(lambda x: make_alert_status(x, ""))
        view_schedule = view_schedule.sort_values(["날짜_dt", "등록일시"], ascending=[True, False])

        urgent_schedule = view_schedule[
            view_schedule["상태표시"].isin(["지남", "오늘", "긴급", "임박"])
        ]

        if urgent_schedule.empty:
            st.info("임박한 일정이 없습니다.")
        else:
            show_alert_table(urgent_schedule[["일정명", "날짜", "D-Day", "상태표시", "작성자"]])


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
    init_tasks_file()
    init_schedule_file()
    init_tax_alert_file()
    init_meeting_alert_file()

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
        "영업 알림",
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
    elif menu == "영업 알림":
        page_alerts()
    elif menu == "관리자 도구":
        page_admin_tools()


if not st.session_state.logged_in:
    login()
else:
    main()
