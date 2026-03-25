import streamlit as st
import pandas as pd
import os
from datetime import datetime
from pathlib import Path

st.set_page_config(page_title="위험도 자동 분석기", layout="wide")

# -----------------------------
# 1. 기본 설정
# -----------------------------
USERS = {
    "admin": {"password": "1234", "role": "admin", "complex": "전체"},
    "행정": {"password": "1234", "role": "staff", "complex": "전체"},
    "무등산자이": {"password": "1234", "role": "client", "complex": "무등산자이"},
}

DATA_FOLDER = Path("data")
LOG_FOLDER = Path("risk_logs")
LOGO_PATH = "logo.png"

DATA_FOLDER.mkdir(exist_ok=True)
LOG_FOLDER.mkdir(exist_ok=True)

# -----------------------------
# 2. 세션 상태 초기화
# -----------------------------
DEFAULT_SESSION = {
    "logged_in": False,
    "username": "",
    "role": "",
    "user_complex": "전체",
    "dashboard_filter": "전체",
    "selected_saved_file": "선택 안 함",
    "selected_admin_user": "전체",
    "last_alert_message": "",
}

for key, value in DEFAULT_SESSION.items():
    if key not in st.session_state:
        st.session_state[key] = value

# -----------------------------
# 3. 공통 함수
# -----------------------------
def show_logo():
    if os.path.exists(LOGO_PATH):
        col1, col2, col3 = st.columns([2, 3, 2])
        with col2:
            st.image(LOGO_PATH, width=180)


def get_role_name(role):
    if role == "admin":
        return "관리자"
    if role == "staff":
        return "행정"
    if role == "client":
        return "고객"
    return role


def login():
    show_logo()
    st.title("🔐 로그인")
    st.write("아이디와 비밀번호를 입력하세요.")

    username = st.text_input("아이디")
    password = st.text_input("비밀번호", type="password")

    if st.button("로그인", use_container_width=True):
        if username in USERS and USERS[username]["password"] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.role = USERS[username]["role"]
            st.session_state.user_complex = USERS[username].get("complex", "전체")
            st.success("로그인 성공")
            st.rerun()
        else:
            st.error("아이디 또는 비밀번호가 올바르지 않습니다.")


def logout():
    for key, value in DEFAULT_SESSION.items():
        st.session_state[key] = value
    st.rerun()


def get_user_folder(username):
    folder = DATA_FOLDER / username
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def save_result(df, username):
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"result_{now}.csv"
    filepath = get_user_folder(username) / filename
    df.to_csv(filepath, index=False, encoding="utf-8-sig")
    return filename, filepath


def get_saved_files(username):
    folder = get_user_folder(username)
    files = sorted(folder.glob("*.csv"), key=lambda x: x.stat().st_mtime, reverse=True)
    return files


def load_saved_file(file_path):
    try:
        return pd.read_csv(file_path)
    except Exception:
        return None


def get_all_saved_files():
    results = []
    for username in USERS.keys():
        folder = get_user_folder(username)
        for file_path in folder.glob("*.csv"):
            results.append(
                {
                    "username": username,
                    "name": file_path.name,
                    "path": file_path,
                    "mtime": file_path.stat().st_mtime,
                }
            )
    results.sort(key=lambda x: x["mtime"], reverse=True)
    return results


def calculate_risk(row):
    risk = 0

    최고온도 = row.get("최고 온도", 0)
    차량감지 = row.get("차량 감지", 0)
    이벤트종류 = row.get("이벤트 종류", 0)

    try:
        최고온도 = float(최고온도)
    except Exception:
        최고온도 = 0

    try:
        차량감지 = float(차량감지)
    except Exception:
        차량감지 = 0

    try:
        이벤트종류 = float(이벤트종류)
    except Exception:
        이벤트종류 = 0

    if 최고온도 >= 12:
        risk += 30
    if 차량감지 == 1:
        risk += 20
    if 이벤트종류 >= 20:
        risk += 50

    return risk


def classify_risk(risk):
    if risk >= 70:
        return "위험"
    elif risk >= 30:
        return "주의"
    return "정상"


def normalize_complex_column(df):
    df = df.copy()
    if "단지명" not in df.columns:
        df["단지명"] = "전체"
    df["단지명"] = df["단지명"].fillna("전체").astype(str)
    return df


def process_dataframe(df, uploaded_file_name=""):
    df = df.copy()
    required_columns = ["최고 온도", "차량 감지", "이벤트 종류"]
    missing_cols = [col for col in required_columns if col not in df.columns]

    if missing_cols:
        raise ValueError(f"CSV에 필요한 컬럼이 없습니다: {', '.join(missing_cols)}")

    df = normalize_complex_column(df)
    df["위험도"] = df.apply(calculate_risk, axis=1)
    df["판정"] = df["위험도"].apply(classify_risk)

    if uploaded_file_name:
        df["업로드파일명"] = uploaded_file_name

    return df


def apply_user_complex_scope(df):
    if df is None or df.empty:
        return df

    df = normalize_complex_column(df)

    if st.session_state.role == "client" and st.session_state.user_complex != "전체":
        df = df[df["단지명"].astype(str) == st.session_state.user_complex]

    return df


def apply_complex_filter(df, key_prefix="main"):
    if df is None or df.empty:
        return df

    filtered_df = df.copy()
    filtered_df = normalize_complex_column(filtered_df)
    filtered_df = apply_user_complex_scope(filtered_df)

    st.markdown("### 🏢 단지 / 위치 필터")

    col0, col1, col2, col3 = st.columns(4)

    complex_name = "전체"
    dong = "전체"
    floor = "전체"
    area = "전체"

    if st.session_state.role == "client" and st.session_state.user_complex != "전체":
        st.info(f"현재 계정은 **{st.session_state.user_complex}** 단지만 조회 가능합니다.")
    else:
        complex_list = ["전체"] + sorted([str(x) for x in filtered_df["단지명"].dropna().unique()])
        with col0:
            complex_name = st.selectbox("단지 선택", complex_list, key=f"{key_prefix}_complex")

    if "동" in filtered_df.columns:
        dong_list = ["전체"] + sorted([str(x) for x in filtered_df["동"].dropna().unique()])
        with col1:
            dong = st.selectbox("동 선택", dong_list, key=f"{key_prefix}_dong")

    if "층" in filtered_df.columns:
        floor_list = ["전체"] + sorted([str(x) for x in filtered_df["층"].dropna().unique()])
        with col2:
            floor = st.selectbox("층 선택", floor_list, key=f"{key_prefix}_floor")

    if "구역" in filtered_df.columns:
        area_list = ["전체"] + sorted([str(x) for x in filtered_df["구역"].dropna().unique()])
        with col3:
            area = st.selectbox("구역 선택", area_list, key=f"{key_prefix}_area")

    if st.session_state.role != "client" and complex_name != "전체":
        filtered_df = filtered_df[filtered_df["단지명"].astype(str) == complex_name]

    if "동" in filtered_df.columns and dong != "전체":
        filtered_df = filtered_df[filtered_df["동"].astype(str) == dong]

    if "층" in filtered_df.columns and floor != "전체":
        filtered_df = filtered_df[filtered_df["층"].astype(str) == floor]

    if "구역" in filtered_df.columns and area != "전체":
        filtered_df = filtered_df[filtered_df["구역"].astype(str) == area]

    return filtered_df


def save_risk_log(df, username):
    if df is None or df.empty:
        return

    df = df.copy()
    df = normalize_complex_column(df)
    df["기록시간"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df["기록사용자"] = username

    log_file = LOG_FOLDER / "risk_log.csv"

    if log_file.exists():
        try:
            old_df = pd.read_csv(log_file)
            df = pd.concat([old_df, df], ignore_index=True)
        except Exception:
            pass

    df.to_csv(log_file, index=False, encoding="utf-8-sig")


def load_risk_log():
    log_file = LOG_FOLDER / "risk_log.csv"
    if log_file.exists():
        try:
            return pd.read_csv(log_file)
        except Exception:
            return None
    return None


def show_danger_alert(df):
    if df is None or df.empty or "판정" not in df.columns:
        return

    df = apply_user_complex_scope(df)
    danger_df = df[df["판정"] == "위험"].copy()

    if danger_df.empty:
        return

    danger_count = len(danger_df)
    first_row = danger_df.iloc[0]

    complex_text = first_row["단지명"] if "단지명" in danger_df.columns else "-"
    dong_text = first_row["동"] if "동" in danger_df.columns else "-"
    floor_text = first_row["층"] if "층" in danger_df.columns else "-"
    area_text = first_row["구역"] if "구역" in danger_df.columns else "-"

    message = f"🚨 위험 데이터 {danger_count}건 발생 / {complex_text} / {dong_text} / {floor_text} / {area_text}"

    st.markdown(
        f"""
        <div style="
            background:#ffebee;
            border:2px solid #d62828;
            border-radius:16px;
            padding:18px;
            margin-bottom:20px;
            text-align:center;
            font-size:24px;
            font-weight:800;
            color:#b00020;
            box-shadow:0 0 16px rgba(214,40,40,0.25);
        ">
            {message}
        </div>
        """,
        unsafe_allow_html=True
    )


def create_test_danger_data():
    test_df = pd.DataFrame([
        {"단지명": "무등산자이", "최고 온도": 15, "차량 감지": 1, "이벤트 종류": 25, "동": "101동", "층": "지하 2층", "구역": "A구역"},
        {"단지명": "무등산자이", "최고 온도": 13, "차량 감지": 1, "이벤트 종류": 21, "동": "101동", "층": "지하 1층", "구역": "B구역"},
        {"단지명": "무등산자이", "최고 온도": 11, "차량 감지": 1, "이벤트 종류": 22, "동": "102동", "층": "지하 2층", "구역": "A구역"},
        {"단지명": "센트럴파크", "최고 온도": 8, "차량 감지": 0, "이벤트 종류": 5, "동": "102동", "층": "지상 1층", "구역": "C구역"},
        {"단지명": "센트럴파크", "최고 온도": 9, "차량 감지": 1, "이벤트 종류": 10, "동": "103동", "층": "지하 3층", "구역": "D구역"},
        {"단지명": "센트럴파크", "최고 온도": 14, "차량 감지": 1, "이벤트 종류": 30, "동": "103동", "층": "지하 2층", "구역": "충전구역 3번"},
    ])
    return process_dataframe(test_df, "test_data.csv")


def render_result_section(df, key_prefix="result", show_download_name="분석결과_통합.csv"):
    if df is None or df.empty:
        st.warning("표시할 데이터가 없습니다.")
        return

    scoped_df = apply_user_complex_scope(df)

    show_danger_alert(scoped_df)

    st.subheader("분석 결과")
    filtered_df = apply_complex_filter(scoped_df, key_prefix=key_prefix)

    if filtered_df.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
        return

    st.dataframe(filtered_df, use_container_width=True, height=500)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("위험", len(scoped_df[scoped_df["판정"] == "위험"]))
    with c2:
        st.metric("주의", len(scoped_df[scoped_df["판정"] == "주의"]))
    with c3:
        st.metric("정상", len(scoped_df[scoped_df["판정"] == "정상"]))

    chart_df = scoped_df["판정"].value_counts().rename_axis("판정").reset_index(name="건수")
    st.subheader("판정별 건수")
    st.bar_chart(chart_df.set_index("판정"))

    if "단지명" in scoped_df.columns:
        complex_chart_df = scoped_df.groupby(["단지명", "판정"]).size().unstack(fill_value=0)
        st.subheader("단지별 현황")
        st.bar_chart(complex_chart_df)

    csv_data = filtered_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label="현재 화면 결과 다운로드",
        data=csv_data,
        file_name=show_download_name,
        mime="text/csv",
        use_container_width=True,
        key=f"download_{key_prefix}",
    )


def admin_dashboard(df):
    st.title("관리자 대시보드")

    if df is None or df.empty:
        st.warning("대시보드에 표시할 데이터가 없습니다.")
        return

    df = normalize_complex_column(df)
    df = apply_user_complex_scope(df)

    danger_df = df[df["판정"] == "위험"].copy()
    warning_df = df[df["판정"] == "주의"].copy()
    normal_df = df[df["판정"] == "정상"].copy()

    danger_count = len(danger_df)
    warning_count = len(warning_df)
    normal_count = len(normal_df)
    total_rows = len(df)
    total_users = len(USERS)

    if danger_count > 0:
        first_row = danger_df.iloc[0]
        st.markdown(
            f"""
            <div style="
                background:#ffebee;
                border:2px solid #d62828;
                border-radius:16px;
                padding:18px;
                margin-bottom:20px;
                text-align:center;
                font-size:24px;
                font-weight:800;
                color:#b00020;
                box-shadow:0 0 16px rgba(214,40,40,0.25);
                line-height:1.8;
            ">
                🚨 위험 데이터 {danger_count}건 발생<br>
                📍 단지: {first_row.get("단지명", "-")} / 위치: {first_row.get("동", "-")} / {first_row.get("층", "-")} / {first_row.get("구역", "-")}
            </div>
            """,
            unsafe_allow_html=True
        )

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.metric("전체 사용자", total_users)
    with c2:
        st.metric("전체 데이터", total_rows)
    with c3:
        st.metric("위험", danger_count)
    with c4:
        st.metric("주의", warning_count)
    with c5:
        st.metric("정상", normal_count)

    b1, b2, b3, b4 = st.columns(4)
    with b1:
        if st.button("🔴 위험 보기", use_container_width=True):
            st.session_state.dashboard_filter = "위험"
            st.rerun()
    with b2:
        if st.button("🟡 주의 보기", use_container_width=True):
            st.session_state.dashboard_filter = "주의"
            st.rerun()
    with b3:
        if st.button("🟢 정상 보기", use_container_width=True):
            st.session_state.dashboard_filter = "정상"
            st.rerun()
    with b4:
        if st.button("📋 전체 보기", use_container_width=True):
            st.session_state.dashboard_filter = "전체"
            st.rerun()

    st.info(f"현재 필터: {st.session_state.dashboard_filter}")

    current_filter = st.session_state.dashboard_filter
    if current_filter == "위험":
        filtered_df = danger_df
    elif current_filter == "주의":
        filtered_df = warning_df
    elif current_filter == "정상":
        filtered_df = normal_df
    else:
        filtered_df = df

    st.subheader("필터 결과")
    filtered_df = apply_complex_filter(filtered_df, key_prefix="dashboard")

    if filtered_df.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        st.dataframe(filtered_df, use_container_width=True, height=450)

    st.markdown("---")
    st.subheader("📜 위험 발생 이력")

    log_df = load_risk_log()
    if log_df is not None and not log_df.empty:
        log_df = apply_user_complex_scope(log_df)
        log_df = apply_complex_filter(log_df, key_prefix="dashboard_log")
        st.dataframe(log_df, use_container_width=True, height=300)
    else:
        st.info("저장된 위험 이력이 없습니다.")


def admin_all_users_section():
    if st.session_state.role != "admin":
        return

    st.subheader("관리자 전체 사용자 통합 조회")

    all_user_names = sorted(USERS.keys())
    admin_user_options = ["전체"] + all_user_names

    selected_admin_user = st.selectbox(
        "조회할 사용자 선택",
        options=admin_user_options,
        key="selected_admin_user"
    )

    all_files = get_all_saved_files()

    if selected_admin_user != "전체":
        all_files = [f for f in all_files if f["username"] == selected_admin_user]

    admin_search_text = st.text_input(
        "전체 사용자 파일명 검색",
        placeholder="예: 무등산자이 또는 20260320",
        key="admin_search_text"
    )

    if admin_search_text.strip():
        search_lower = admin_search_text.lower()
        all_files = [
            f for f in all_files
            if search_lower in f["name"].lower() or search_lower in f["username"].lower()
        ]

    filtered_files = all_files[:50]

    if not filtered_files:
        st.info("조건에 맞는 사용자 저장 파일이 없습니다.")
        return

    st.caption(f"검색 결과 {len(all_files)}개 중 최근 50개 표시")

    display_options = ["선택 안 함"]
    option_map = {}

    for item in filtered_files:
        display_name = f"[{item['username']}] {item['name']}"
        display_options.append(display_name)
        option_map[display_name] = item

    selected_admin_file = st.selectbox(
        "전체 사용자 CSV 선택",
        options=display_options,
        key="selected_admin_file"
    )

    if st.button("관리자 불러오기", use_container_width=True):
        if selected_admin_file == "선택 안 함":
            st.warning("불러올 파일을 먼저 선택해주세요.")
        else:
            selected_item = option_map[selected_admin_file]
            saved_df = load_saved_file(selected_item["path"])

            if saved_df is not None:
                st.success(f"관리자 불러오기 완료: {selected_admin_file}")
                render_result_section(
                    saved_df,
                    key_prefix="admin_all",
                    show_download_name=selected_item["name"]
                )
            else:
                st.error("파일을 불러오지 못했습니다.")

# -----------------------------
# 4. 메인 화면
# -----------------------------
def main():
    if not st.session_state.logged_in:
        login()
        return

    show_logo()

    col1, col2 = st.columns([6, 1])
    with col1:
        st.title("화재 위험도 자동 분석 프로그램 - (주)윤우테크")
        st.caption(
            f"로그인: {st.session_state.username} | 권한: {get_role_name(st.session_state.role)}"
        )
    with col2:
        st.write("")
        if st.button("로그아웃", use_container_width=True):
            logout()

    st.write("CSV 파일을 업로드하면 위험도와 판정을 자동 계산합니다.")

    menu_options = ["데이터 분석", "저장 파일 보기", "관리자 대시보드"]
    if st.session_state.role == "admin":
        menu_options.append("관리자 통합 조회")

    menu = st.sidebar.radio("메뉴 선택", menu_options)

    # -----------------------------
    # 데이터 분석
    # -----------------------------
    if menu == "데이터 분석":
        st.subheader("CSV 업로드 분석")

        if st.session_state.role == "client" and st.session_state.user_complex != "전체":
            st.info(f"현재 계정은 **{st.session_state.user_complex}** 단지 전용 계정입니다.")

        uploaded_files = st.file_uploader(
            "CSV 파일 여러 개 업로드",
            type=["csv"],
            accept_multiple_files=True
        )

        col_btn1, col_btn2 = st.columns(2)

        with col_btn1:
            if st.button("🚨 위험 테스트 데이터 자동 생성", use_container_width=True):
                test_df = create_test_danger_data()
                test_df = apply_user_complex_scope(test_df)

                danger_df = test_df[test_df["판정"] == "위험"].copy()
                if not danger_df.empty:
                    save_risk_log(danger_df, st.session_state.username)

                filename, filepath = save_result(test_df, st.session_state.username)
                st.success(f"테스트 위험 데이터 저장 완료: {filename}")
                render_result_section(test_df, key_prefix="test_data", show_download_name=filename)

        with col_btn2:
            if st.button("📋 테스트 데이터 설명 보기", use_container_width=True):
                st.info("이 버튼은 위험/주의/정상 데이터가 섞인 샘플 데이터를 자동 생성하여 저장합니다.")

        if uploaded_files:
            all_results = []

            for uploaded_file in uploaded_files:
                try:
                    df = pd.read_csv(uploaded_file)
                    result_df = process_dataframe(df, uploaded_file.name)

                    if st.session_state.role == "client" and st.session_state.user_complex != "전체":
                        if "단지명" in result_df.columns:
                            result_df = result_df[result_df["단지명"].astype(str) == st.session_state.user_complex]

                    if not result_df.empty:
                        all_results.append(result_df)
                except Exception as e:
                    st.error(f"{uploaded_file.name} 처리 중 오류 발생: {e}")

            if all_results:
                final_df = pd.concat(all_results, ignore_index=True)

                danger_df = final_df[final_df["판정"] == "위험"].copy()
                if not danger_df.empty:
                    save_risk_log(danger_df, st.session_state.username)

                render_result_section(
                    final_df,
                    key_prefix="analysis_result",
                    show_download_name="분석결과_통합.csv"
                )

                if st.button("결과 저장", use_container_width=True):
                    filename, filepath = save_result(final_df, st.session_state.username)
                    st.success(f"저장 완료: {filename}")
            else:
                st.warning("처리 가능한 데이터가 없습니다.")

    # -----------------------------
    # 저장 파일 보기
    # -----------------------------
    elif menu == "저장 파일 보기":
        st.subheader("저장된 결과 파일")

        saved_files = get_saved_files(st.session_state.username)

        if not saved_files:
            st.info("저장된 파일이 없습니다.")
        else:
            file_names = [f.name for f in saved_files]
            selected_file_name = st.selectbox("파일 선택", file_names)

            selected_path = next((f for f in saved_files if f.name == selected_file_name), None)

            if selected_path is not None:
                loaded_df = load_saved_file(selected_path)

                if loaded_df is not None:
                    st.success(f"불러온 파일: {selected_file_name}")
                    render_result_section(
                        loaded_df,
                        key_prefix="saved_file",
                        show_download_name=selected_file_name
                    )
                else:
                    st.error("파일을 불러오지 못했습니다.")

    # -----------------------------
    # 관리자 대시보드
    # -----------------------------
    elif menu == "관리자 대시보드":
        all_files = get_all_saved_files() if st.session_state.role == "admin" else [
            {
                "username": st.session_state.username,
                "name": f.name,
                "path": f,
                "mtime": f.stat().st_mtime
            }
            for f in get_saved_files(st.session_state.username)
        ]

        if not all_files:
            st.warning("저장된 파일이 없어 대시보드에 표시할 데이터가 없습니다.")
        else:
            latest_path = all_files[0]["path"]
            dashboard_df = load_saved_file(latest_path)

            if dashboard_df is not None:
                st.caption(f"기준 파일: {all_files[0]['name']}")
                admin_dashboard(process_dataframe(dashboard_df))
            else:
                st.error("대시보드용 데이터를 불러오지 못했습니다.")

    # -----------------------------
    # 관리자 통합 조회
    # -----------------------------
    elif menu == "관리자 통합 조회":
        admin_all_users_section()


# -----------------------------
# 실행
# -----------------------------
if __name__ == "__main__":
    main()
