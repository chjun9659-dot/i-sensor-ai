import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.set_page_config(page_title="위험도 자동 분석기", layout="wide")
st.write("버전 확인: 2026-03-25 16:10")

# -----------------------------
# 1. 기본 설정
# -----------------------------
USERS = {
    "admin": "1234",
    "행정": "1234"
}

SAVE_FOLDER = "saved_data"
os.makedirs(SAVE_FOLDER, exist_ok=True)

LOGO_PATH = "logo.png"   # 같은 폴더에 logo.png 넣어두면 자동 표시

# -----------------------------
# 2. 세션 상태 초기화
# -----------------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "username" not in st.session_state:
    st.session_state.username = ""

if "dashboard_filter" not in st.session_state:
    st.session_state.dashboard_filter = "전체"

if "selected_saved_file" not in st.session_state:
    st.session_state.selected_saved_file = "선택 안 함"


# -----------------------------
# 3. 공통 함수
# -----------------------------
def show_logo():
    if os.path.exists(LOGO_PATH):
        col1, col2, col3 = st.columns([2, 3, 2])
        with col2:
            st.image(LOGO_PATH, width=160)


def login():
    show_logo()
    st.title("🔐 로그인")
    st.write("아이디와 비밀번호를 입력하세요.")

    username = st.text_input("아이디")
    password = st.text_input("비밀번호", type="password")

    if st.button("로그인", use_container_width=True):
        if username in USERS and USERS[username] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.success("로그인 성공")
            st.rerun()
        else:
            st.error("아이디 또는 비밀번호가 올바르지 않습니다.")


def logout():
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.rerun()


def calculate_risk(row):
    risk = 0

    최고온도 = row.get("최고 온도", 0)
    차량감지 = row.get("차량 감지", 0)
    이벤트종류 = row.get("이벤트 종류", 0)

    try:
        최고온도 = float(최고온도)
    except:
        최고온도 = 0

    try:
        차량감지 = float(차량감지)
    except:
        차량감지 = 0

    try:
        이벤트종류 = float(이벤트종류)
    except:
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
    else:
        return "정상"


def process_dataframe(df):
    df = df.copy()

    required_columns = ["최고 온도", "차량 감지", "이벤트 종류"]
    missing_cols = [col for col in required_columns if col not in df.columns]

    if missing_cols:
        st.error(f"CSV에 필요한 컬럼이 없습니다: {', '.join(missing_cols)}")
        st.stop()

    df["위험도"] = df.apply(calculate_risk, axis=1)
    df["판정"] = df["위험도"].apply(classify_risk)

    return df


def save_result(df):
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"result_{now}.csv"
    filepath = os.path.join(SAVE_FOLDER, filename)
    df.to_csv(filepath, index=False, encoding="utf-8-sig")
    return filename, filepath


def load_saved_file(file_name):
    filepath = os.path.join(SAVE_FOLDER, file_name)
    if os.path.exists(filepath):
        return pd.read_csv(filepath)
    return None


def get_saved_files():
    files = [f for f in os.listdir(SAVE_FOLDER) if f.endswith(".csv")]
    files.sort(reverse=True)
    return files


def create_test_danger_data():
    test_df = pd.DataFrame([
        {"최고 온도": 15, "차량 감지": 1, "이벤트 종류": 25},
        {"최고 온도": 13, "차량 감지": 1, "이벤트 종류": 21},
        {"최고 온도": 11, "차량 감지": 1, "이벤트 종류": 22},
        {"최고 온도": 8, "차량 감지": 0, "이벤트 종류": 5},
        {"최고 온도": 9, "차량 감지": 1, "이벤트 종류": 10},
        {"최고 온도": 14, "차량 감지": 1, "이벤트 종류": 30}
    ])
    return process_dataframe(test_df)


def admin_dashboard(df, users):
    st.title("관리자 대시보드")

    total_users = len(users)

    if df is None or df.empty:
        st.warning("대시보드에 표시할 데이터가 없습니다.")
        return

    if "판정" not in df.columns:
        st.error("대시보드용 데이터에 '판정' 컬럼이 없습니다.")
        return

    danger_df = df[df["판정"] == "위험"].copy()
    warning_df = df[df["판정"] == "주의"].copy()
    normal_df = df[df["판정"] == "정상"].copy()

    danger_count = len(danger_df)
    warning_count = len(warning_df)
    normal_count = len(normal_df)
    total_rows = len(df)

    if "dashboard_filter" not in st.session_state:
        st.session_state.dashboard_filter = "전체"

    # 위험 경고 박스
    if danger_count > 0:
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
                🚨 위험 데이터 {danger_count}건 발생
            </div>
            """,
            unsafe_allow_html=True
        )

    # 상단 요약 카드
    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.markdown(
            f"""
            <div style="
                background:#f8f9fa;
                border:1px solid #d9dee7;
                border-radius:16px;
                padding:22px;
                text-align:center;
                min-height:120px;
            ">
                <div style="font-size:18px;font-weight:700;">👥 전체 사용자</div>
                <div style="font-size:34px;font-weight:800;margin-top:10px;">{total_users}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:
        st.markdown(
            f"""
            <div style="
                background:#f8f9fa;
                border:1px solid #d9dee7;
                border-radius:16px;
                padding:22px;
                text-align:center;
                min-height:120px;
            ">
                <div style="font-size:18px;font-weight:700;">📁 전체 데이터</div>
                <div style="font-size:34px;font-weight:800;margin-top:10px;">{total_rows}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c3:
        danger_bg = "#ffe5e5" if danger_count > 0 else "#fff5f5"
        danger_border = "#d62828" if danger_count > 0 else "#ffcccc"
        danger_shadow = "0 0 18px rgba(214,40,40,0.35)" if danger_count > 0 else "none"
        st.markdown(
            f"""
            <div style="
                background:{danger_bg};
                border:2px solid {danger_border};
                border-radius:16px;
                padding:22px;
                text-align:center;
                min-height:120px;
                box-shadow:{danger_shadow};
            ">
                <div style="font-size:18px;font-weight:800;color:#d62828;">🔴 위험</div>
                <div style="font-size:42px;font-weight:900;margin-top:10px;color:#d62828;">{danger_count}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c4:
        warning_bg = "#fff8dc" if warning_count > 0 else "#fffdf0"
        warning_border = "#d4a000" if warning_count > 0 else "#f4d35e"
        warning_shadow = "0 0 14px rgba(212,160,0,0.22)" if warning_count > 0 else "none"
        st.markdown(
            f"""
            <div style="
                background:{warning_bg};
                border:2px solid {warning_border};
                border-radius:16px;
                padding:22px;
                text-align:center;
                min-height:120px;
                box-shadow:{warning_shadow};
            ">
                <div style="font-size:18px;font-weight:800;color:#c99700;">🟡 주의</div>
                <div style="font-size:42px;font-weight:900;margin-top:10px;color:#c99700;">{warning_count}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c5:
        normal_bg = "#ecfff0" if normal_count > 0 else "#f3fff3"
        normal_border = "#52b788" if normal_count > 0 else "#b7efc5"
        normal_shadow = "0 0 12px rgba(43,147,72,0.18)" if normal_count > 0 else "none"
        st.markdown(
            f"""
            <div style="
                background:{normal_bg};
                border:2px solid {normal_border};
                border-radius:16px;
                padding:22px;
                text-align:center;
                min-height:120px;
                box-shadow:{normal_shadow};
            ">
                <div style="font-size:18px;font-weight:800;color:#2b9348;">🟢 정상</div>
                <div style="font-size:42px;font-weight:900;margin-top:10px;color:#2b9348;">{normal_count}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.write("")

    # 실제 클릭 버튼
    b1, b2, b3, b4 = st.columns(4)

    with b1:
        if st.button("🔴 위험 보기", key="dash_danger", use_container_width=True):
            st.session_state.dashboard_filter = "위험"
            st.rerun()

    with b2:
        if st.button("🟡 주의 보기", key="dash_warning", use_container_width=True):
            st.session_state.dashboard_filter = "주의"
            st.rerun()

    with b3:
        if st.button("🟢 정상 보기", key="dash_normal", use_container_width=True):
            st.session_state.dashboard_filter = "정상"
            st.rerun()

    with b4:
        if st.button("📋 전체 보기", key="dash_all", use_container_width=True):
            st.session_state.dashboard_filter = "전체"
            st.rerun()

    st.markdown("---")
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

    if filtered_df.empty:
        st.warning(f"{current_filter} 데이터가 없습니다.")
    else:
        st.dataframe(filtered_df, use_container_width=True, height=500)


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
    with col2:
        st.write("")
        st.write(f"로그인: {st.session_state.username}")
        if st.button("로그아웃", use_container_width=True):
            logout()

    st.write("CSV 파일을 업로드하면 위험도와 판정을 자동 계산합니다.")

    menu = st.sidebar.radio(
        "메뉴 선택",
        ["데이터 분석", "저장 파일 보기", "관리자 대시보드"]
    )

    # -----------------------------
    # 데이터 분석
    # -----------------------------
    if menu == "데이터 분석":
        st.subheader("CSV 업로드 분석")

        uploaded_files = st.file_uploader(
            "CSV 파일 여러 개 업로드",
            type=["csv"],
            accept_multiple_files=True
        )

        col_btn1, col_btn2 = st.columns(2)

        with col_btn1:
            if st.button("🚨 위험 테스트 데이터 자동 생성", use_container_width=True):
                test_df = create_test_danger_data()
                filename, filepath = save_result(test_df)
                st.success(f"테스트 위험 데이터 저장 완료: {filename}")
                st.dataframe(test_df, use_container_width=True)

        with col_btn2:
            if st.button("📋 테스트 데이터 설명 보기", use_container_width=True):
                st.info("이 버튼은 위험/주의/정상 데이터가 섞인 샘플 데이터를 자동 생성하여 저장합니다.")

        if uploaded_files:
            all_results = []

            for uploaded_file in uploaded_files:
                try:
                    df = pd.read_csv(uploaded_file)
                    result_df = process_dataframe(df)
                    result_df["업로드파일명"] = uploaded_file.name
                    all_results.append(result_df)
                except Exception as e:
                    st.error(f"{uploaded_file.name} 처리 중 오류 발생: {e}")

            if all_results:
                final_df = pd.concat(all_results, ignore_index=True)

                st.subheader("분석 결과")
                st.dataframe(final_df, use_container_width=True)

                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("위험", len(final_df[final_df["판정"] == "위험"]))
                with c2:
                    st.metric("주의", len(final_df[final_df["판정"] == "주의"]))
                with c3:
                    st.metric("정상", len(final_df[final_df["판정"] == "정상"]))

                csv_data = final_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button(
                    label="분석 결과 다운로드",
                    data=csv_data,
                    file_name="분석결과_통합.csv",
                    mime="text/csv",
                    use_container_width=True
                )

                if st.button("결과 저장", use_container_width=True):
                    filename, filepath = save_result(final_df)
                    st.success(f"저장 완료: {filename}")

    # -----------------------------
    # 저장 파일 보기
    # -----------------------------
    elif menu == "저장 파일 보기":
        st.subheader("저장된 결과 파일")

        saved_files = get_saved_files()

        if not saved_files:
            st.info("저장된 파일이 없습니다.")
        else:
            selected_file = st.selectbox("파일 선택", saved_files)

            if selected_file:
                loaded_df = load_saved_file(selected_file)

                if loaded_df is not None:
                    st.success(f"불러온 파일: {selected_file}")
                    st.dataframe(loaded_df, use_container_width=True)

                    csv_data = loaded_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                    st.download_button(
                        label="이 파일 다운로드",
                        data=csv_data,
                        file_name=selected_file,
                        mime="text/csv",
                        use_container_width=True
                    )
                else:
                    st.error("파일을 불러오지 못했습니다.")

    # -----------------------------
    # 관리자 대시보드
    # -----------------------------
    elif menu == "관리자 대시보드":
        saved_files = get_saved_files()

        if not saved_files:
            st.warning("저장된 파일이 없어 대시보드에 표시할 데이터가 없습니다.")
        else:
            latest_file = saved_files[0]
            dashboard_df = load_saved_file(latest_file)

            if dashboard_df is not None:
                st.caption(f"기준 파일: {latest_file}")
                admin_dashboard(dashboard_df, USERS)
            else:
                st.error("대시보드용 데이터를 불러오지 못했습니다.")


# -----------------------------
# 실행
# -----------------------------
if __name__ == "__main__":
    main()
