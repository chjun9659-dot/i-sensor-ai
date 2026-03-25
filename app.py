import streamlit as st
import pandas as pd
import altair as alt
import os
from datetime import datetime
from pathlib import Path

st.set_page_config(page_title="위험도 자동 분석기", layout="wide")

# -----------------------------
# 1. 로그인용 사용자 정보
# -----------------------------
USERS = {
    "admin": {"password": "1234", "role": "admin"},
    "staff1": {"password": "1111", "role": "staff"},
    "client1": {"password": "2222", "role": "client"}
}

# -----------------------------
# 2. 세션 상태 초기화
# -----------------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "username" not in st.session_state:
    st.session_state.username = ""

if "role" not in st.session_state:
    st.session_state.role = ""

if "selected_saved_file" not in st.session_state:
    st.session_state.selected_saved_file = "선택 안 함"

if "selected_admin_user" not in st.session_state:
    st.session_state.selected_admin_user = "전체"


# -----------------------------
# 3. 로그인 / 로그아웃
# -----------------------------
def login():
    st.title("🔐 로그인")
    st.write("아이디와 비밀번호를 입력하세요.")

    username = st.text_input("아이디")
    password = st.text_input("비밀번호", type="password")

    if st.button("로그인"):
        if username in USERS and USERS[username]["password"] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.role = USERS[username]["role"]
            st.success(f"{username} 로그인 성공!")
            st.rerun()
        else:
            st.error("아이디 또는 비밀번호가 틀렸습니다.")


def logout():
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.session_state.selected_saved_file = "선택 안 함"
    st.session_state.selected_admin_user = "전체"
    st.rerun()


# -----------------------------
# 4. 업로드 파일 자동 저장
# -----------------------------
def save_uploaded_file(uploaded_file, username):
    folder = os.path.join("data", username)
    os.makedirs(folder, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    safe_filename = uploaded_file.name.replace(" ", "_")
    file_path = os.path.join(folder, f"{timestamp}_{safe_filename}")

    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    return file_path


# -----------------------------
# 5. 저장 파일 목록 / 불러오기
# -----------------------------
def get_saved_files(username):
    folder = Path("data") / username
    if not folder.exists():
        return []

    files = sorted(folder.glob("*.csv"), key=lambda x: x.stat().st_mtime, reverse=True)
    return files


def get_all_saved_files():
    all_files = []

    base_folder = Path("data")
    if not base_folder.exists():
        return all_files

    for user_folder in base_folder.iterdir():
        if user_folder.is_dir():
            user_files = sorted(
                user_folder.glob("*.csv"),
                key=lambda x: x.stat().st_mtime,
                reverse=True
            )
            for file in user_files:
                all_files.append({
                    "username": user_folder.name,
                    "path": file,
                    "name": file.name,
                    "mtime": file.stat().st_mtime
                })

    all_files = sorted(all_files, key=lambda x: x["mtime"], reverse=True)
    return all_files


def load_saved_csv(file_path):
    try:
        return pd.read_csv(file_path, encoding="utf-8")
    except:
        try:
            return pd.read_csv(file_path, encoding="cp949")
        except Exception as e:
            st.error(f"저장 파일을 읽는 중 오류가 발생했습니다: {e}")
            return None


# -----------------------------
# 6. 보조 함수
# -----------------------------
def load_csv_safely(file_path):
    try:
        return pd.read_csv(file_path, encoding="utf-8")
    except:
        try:
            return pd.read_csv(file_path, encoding="cp949")
        except:
            return None


def to_number(value):
    return pd.to_numeric(value, errors="coerce")


def calculate_risk(row):
    risk = 0

    최고온도 = to_number(row.get("최고 온도"))
    차량감지 = to_number(row.get("차량 감지"))
    이벤트종류 = to_number(row.get("이벤트 종류"))

    if pd.notna(최고온도) and 최고온도 >= 12:
        risk += 30
    if pd.notna(차량감지) and 차량감지 == 1:
        risk += 20
    if pd.notna(이벤트종류) and 이벤트종류 >= 20:
        risk += 50

    return risk


def classify_risk(risk):
    if risk >= 80:
        return "위험"
    elif risk >= 50:
        return "주의"
    else:
        return "정상"
def show_alert(df):
    if "판정" in df.columns:
        danger_count = (df["판정"] == "위험").sum()

        if danger_count > 0:
            st.error(f"🚨 위험 발생 {danger_count}건! 즉시 확인하세요!")

            st.markdown("""
                <audio autoplay>
                <source src="https://actions.google.com/sounds/v1/alarms/alarm_clock.ogg" type="audio/ogg">
                </audio>
            """, unsafe_allow_html=True)

        else:
            st.success("✅ 현재 상태 정상")

def highlight_risk(row):
    if row["판정"] == "위험":
        return ["background-color: #ffb3b3"] * len(row)
    elif row["판정"] == "주의":
        return ["background-color: #ffe699"] * len(row)
    else:
        return [""] * len(row)


def shorten_filename(filename, front=14, back=18):
    if len(filename) <= front + back + 3:
        return filename
    return filename[:front] + "..." + filename[-back:]


def parse_category_date(value):
    try:
        return pd.to_datetime(value, errors="coerce")
    except:
        return pd.NaT


def get_role_name(role):
    if role == "admin":
        return "관리자"
    elif role == "staff":
        return "직원"
    elif role == "client":
        return "고객"
    return role


def summarize_all_files():
    all_files = get_all_saved_files()

    total_users = len(list(USERS.keys()))
    total_files = len(all_files)

    total_danger = 0
    total_warning = 0
    total_normal = 0

    recent_files = []

    for item in all_files:
        df = load_csv_safely(item["path"])
        if df is None:
            continue

        required_columns = ["최고 온도", "차량 감지", "이벤트 종류"]
        if not all(col in df.columns for col in required_columns):
            continue

        df = df.copy()
        df["위험도"] = df.apply(calculate_risk, axis=1)
        df["판정"] = df["위험도"].apply(classify_risk)

        total_danger += (df["판정"] == "위험").sum()
        total_warning += (df["판정"] == "주의").sum()
        total_normal += (df["판정"] == "정상").sum()

        recent_files.append({
            "사용자": item["username"],
            "파일명": item["name"],
            "업로드시간": datetime.fromtimestamp(item["mtime"]).strftime("%Y-%m-%d %H:%M:%S")
        })

    recent_files_df = pd.DataFrame(recent_files).head(10)

    return {
        "total_users": total_users,
        "total_files": total_files,
        "total_danger": int(total_danger),
        "total_warning": int(total_warning),
        "total_normal": int(total_normal),
        "recent_files_df": recent_files_df
    }


def admin_dashboard():
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">관리자 대시보드</div>', unsafe_allow_html=True)

    summary = summarize_all_files()

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.markdown(
            f'<div class="kpi-card">👥 전체 사용자<br><span style="font-size:34px;">{summary["total_users"]}</span></div>',
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            f'<div class="kpi-card">📁 전체 파일<br><span style="font-size:34px;">{summary["total_files"]}</span></div>',
            unsafe_allow_html=True
        )

    with col3:
        st.markdown(
            f'<div class="kpi-card kpi-red">🔴 위험<br><span style="font-size:34px;">{summary["total_danger"]}</span></div>',
            unsafe_allow_html=True
        )

    with col4:
        st.markdown(
            f'<div class="kpi-card kpi-yellow">🟡 주의<br><span style="font-size:34px;">{summary["total_warning"]}</span></div>',
            unsafe_allow_html=True
        )

    with col5:
        st.markdown(
            f'<div class="kpi-card kpi-green">🟢 정상<br><span style="font-size:34px;">{summary["total_normal"]}</span></div>',
            unsafe_allow_html=True
        )

    st.markdown('<div class="section-title">최근 업로드 파일</div>', unsafe_allow_html=True)

    if not summary["recent_files_df"].empty:
        st.dataframe(summary["recent_files_df"], use_container_width=True)
    else:
        st.info("최근 업로드된 파일이 없습니다.")


# -----------------------------
# 7. 분석 공통 함수
# -----------------------------
def analyze_dataframes(dataframes_with_names, key_suffix="main"):
    all_results = []
    danger_rows = []

    required_columns = ["최고 온도", "차량 감지", "이벤트 종류"]

    for file_name, df in dataframes_with_names:
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            st.error(f"{file_name} 파일에 필요한 열이 없습니다: {missing_columns}")
            continue

        df = df.copy()
        df["위험도"] = df.apply(calculate_risk, axis=1)
        df["판정"] = df["위험도"].apply(classify_risk)
        df["파일명"] = file_name

        all_results.append(df)

        danger_df = df[df["판정"].isin(["주의", "위험"])]
        if not danger_df.empty:
            danger_rows.append(danger_df)

        st.markdown(f'<div class="section-title">파일 결과: {file_name}</div>', unsafe_allow_html=True)
        st.dataframe(df.style.apply(highlight_risk, axis=1), use_container_width=True)
        show_alert(df)

    if not all_results:
        return

    final_result_df = pd.concat(all_results, ignore_index=True)

    위험개수 = (final_result_df["판정"] == "위험").sum()
    주의개수 = (final_result_df["판정"] == "주의").sum()
    정상개수 = (final_result_df["판정"] == "정상").sum()

    st.markdown("<hr>", unsafe_allow_html=True)

    if 위험개수 > 0:
        st.markdown(
            f"""
            <div style="
                background-color:#ff4d4d;
                color:white;
                padding:20px;
                border-radius:12px;
                font-size:22px;
                font-weight:800;
                text-align:center;
            ">
            🚨 긴급 경고 🚨<br>
            위험 데이터 {위험개수}건 발생! 즉시 확인하세요!
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.session_state.role == "admin":
            st.warning("관리자 알림: 즉시 조치 필요")

    elif 주의개수 > 0:
        st.markdown(
            f'<div class="alert-medium-risk">⚠️ 주의 데이터가 {주의개수}건 발견되었습니다. 점검을 권장합니다.</div>',
            unsafe_allow_html=True
        )
    else:
        st.success("✅ 현재 모든 데이터가 정상 상태입니다.")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(
            f'<div class="kpi-card kpi-red">🔴 위험<br><span style="font-size:36px;">{위험개수}</span></div>',
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            f'<div class="kpi-card kpi-yellow">🟡 주의<br><span style="font-size:36px;">{주의개수}</span></div>',
            unsafe_allow_html=True
        )

    with col3:
        st.markdown(
            f'<div class="kpi-card kpi-green">🟢 정상<br><span style="font-size:36px;">{정상개수}</span></div>',
            unsafe_allow_html=True
        )

    risk_color_scale = alt.Scale(
        domain=["정상", "주의", "위험"],
        range=["#2a9d2f", "#d4a000", "#d62828"]
    )

    st.markdown('<div class="section-title">위험도 요약 그래프</div>', unsafe_allow_html=True)

    risk_ratio = final_result_df["판정"].value_counts(normalize=True).reindex(
        ["정상", "주의", "위험"],
        fill_value=0
    ) * 100

    risk_chart_df = pd.DataFrame({
        "위험등급": ["정상", "주의", "위험"],
        "비율": [
            risk_ratio["정상"],
            risk_ratio["주의"],
            risk_ratio["위험"]
        ]
    })

    st.write("위험도 비율 (%)")

    chart = alt.Chart(risk_chart_df).mark_bar(size=80).encode(
        x=alt.X(
            "위험등급:N",
            sort=["정상", "주의", "위험"],
            axis=alt.Axis(labelAngle=0, title="위험등급")
        ),
        y=alt.Y("비율:Q", title="비율 (%)"),
        color=alt.Color("위험등급:N", scale=risk_color_scale, legend=None),
        tooltip=[
            alt.Tooltip("위험등급:N", title="위험등급"),
            alt.Tooltip("비율:Q", title="비율", format=".2f")
        ]
    ).properties(height=400)

    st.altair_chart(chart, use_container_width=True)

    file_risk_count = pd.crosstab(final_result_df["파일명"], final_result_df["판정"]).reindex(
        columns=["정상", "주의", "위험"],
        fill_value=0
    )

    file_risk_count = file_risk_count.reset_index()
    file_risk_count["파일표시명"] = file_risk_count["파일명"].apply(shorten_filename)

    file_risk_chart_df = file_risk_count.melt(
        id_vars=["파일명", "파일표시명"],
        value_vars=["정상", "주의", "위험"],
        var_name="위험등급",
        value_name="건수"
    )

    st.markdown('<div class="section-title">파일별 위험도 분포</div>', unsafe_allow_html=True)

    file_chart = alt.Chart(file_risk_chart_df).mark_bar().encode(
        y=alt.Y("파일표시명:N", sort="-x", title="파일명"),
        x=alt.X("건수:Q", title="건수"),
        color=alt.Color(
            "위험등급:N",
            scale=risk_color_scale,
            legend=alt.Legend(title="위험등급")
        ),
        tooltip=[
            alt.Tooltip("파일명:N", title="전체 파일명"),
            alt.Tooltip("위험등급:N", title="위험등급"),
            alt.Tooltip("건수:Q", title="건수")
        ]
    ).properties(height=max(400, len(file_risk_count) * 28))

    st.altair_chart(file_chart, use_container_width=True)

    st.markdown('<div class="section-title">위험 파일 집중 분석</div>', unsafe_allow_html=True)

    filter_option = st.radio(
        "표시 기준 선택",
        ["주의+위험 모두 보기", "위험만 보기"],
        horizontal=True,
        key=f"filter_option_{key_suffix}"
    )

    if filter_option == "위험만 보기":
        danger_file_df = file_risk_chart_df[
            (file_risk_chart_df["위험등급"] == "위험") &
            (file_risk_chart_df["건수"] > 0)
        ]
    else:
        danger_file_df = file_risk_chart_df[
            file_risk_chart_df["위험등급"].isin(["주의", "위험"])
        ]
        danger_file_df = danger_file_df[danger_file_df["건수"] > 0]

    if not danger_file_df.empty:
        danger_chart = alt.Chart(danger_file_df).mark_bar().encode(
            y=alt.Y("파일표시명:N", sort="-x", title="파일명"),
            x=alt.X("건수:Q", title="위험 건수"),
            color=alt.Color(
                "위험등급:N",
                scale=alt.Scale(domain=["주의", "위험"], range=["#d4a000", "#d62828"]),
                legend=alt.Legend(title="위험등급")
            ),
            tooltip=[
                alt.Tooltip("파일명:N", title="전체 파일명"),
                alt.Tooltip("위험등급:N", title="위험등급"),
                alt.Tooltip("건수:Q", title="건수")
            ]
        ).properties(height=max(300, len(danger_file_df["파일표시명"].unique()) * 30))

        st.altair_chart(danger_chart, use_container_width=True)

        filtered_file_names = danger_file_df["파일명"].unique().tolist()
        filtered_result_df = final_result_df[final_result_df["파일명"].isin(filtered_file_names)]

        st.markdown('<div class="section-title">필터 적용 파일 목록</div>', unsafe_allow_html=True)
        st.dataframe(
            filtered_result_df.style.apply(highlight_risk, axis=1),
            use_container_width=True
        )
    else:
        if filter_option == "위험만 보기":
            st.success("🎉 위험 등급 파일이 없습니다.")
        else:
            st.success("🎉 주의 또는 위험 파일이 없습니다. 모든 파일이 정상 상태입니다.")

    if "category" in final_result_df.columns:
        trend_df = final_result_df.copy()
        trend_df["일자"] = trend_df["category"].apply(parse_category_date)
        trend_df = trend_df.dropna(subset=["일자"])

        if not trend_df.empty:
            trend_group = trend_df.groupby(["일자", "판정"]).size().reset_index(name="건수")

            st.markdown('<div class="section-title">일자별 위험 추이 그래프</div>', unsafe_allow_html=True)

            trend_chart = alt.Chart(trend_group).mark_line(point=True).encode(
                x=alt.X("일자:T", title="일자"),
                y=alt.Y("건수:Q", title="건수"),
                color=alt.Color(
                    "판정:N",
                    scale=risk_color_scale,
                    legend=alt.Legend(title="위험등급")
                ),
                tooltip=[
                    alt.Tooltip("일자:T", title="일자"),
                    alt.Tooltip("판정:N", title="위험등급"),
                    alt.Tooltip("건수:Q", title="건수")
                ]
            ).properties(height=400)

            st.altair_chart(trend_chart, use_container_width=True)
        else:
            st.info("일자별 위험 추이 그래프를 표시할 날짜 데이터가 없습니다.")
    else:
        st.info("category 열이 없어 일자별 위험 추이 그래프를 표시할 수 없습니다.")

    st.markdown('<div class="section-title">온도 vs 위험도 상관 분석</div>', unsafe_allow_html=True)

    scatter_df = final_result_df.copy()
    scatter_df["최고 온도"] = pd.to_numeric(scatter_df["최고 온도"], errors="coerce")
    scatter_df["위험도"] = pd.to_numeric(scatter_df["위험도"], errors="coerce")
    scatter_df = scatter_df.dropna(subset=["최고 온도", "위험도"])

    if not scatter_df.empty:
        scatter_chart = alt.Chart(scatter_df).mark_circle(size=90).encode(
            x=alt.X("최고 온도:Q", title="최고 온도"),
            y=alt.Y("위험도:Q", title="위험도"),
            color=alt.Color(
                "판정:N",
                scale=risk_color_scale,
                legend=alt.Legend(title="위험등급")
            ),
            tooltip=[
                alt.Tooltip("파일명:N", title="파일명"),
                alt.Tooltip("최고 온도:Q", title="최고 온도"),
                alt.Tooltip("차량 감지:N", title="차량 감지"),
                alt.Tooltip("이벤트 종류:N", title="이벤트 종류"),
                alt.Tooltip("위험도:Q", title="위험도"),
                alt.Tooltip("판정:N", title="판정")
            ]
        ).properties(height=450)

        st.altair_chart(scatter_chart, use_container_width=True)

        try:
            corr_value = scatter_df["최고 온도"].corr(scatter_df["위험도"])
            if pd.notna(corr_value):
                st.write(f"최고 온도와 위험도의 상관계수: **{corr_value:.3f}**")
            else:
                st.write("최고 온도와 위험도의 상관계수를 계산할 수 없습니다.")
        except:
            st.write("최고 온도와 위험도의 상관계수를 계산할 수 없습니다.")
    else:
        st.info("상관 분석에 사용할 숫자 데이터가 부족합니다.")

    if danger_rows:
        danger_summary_df = pd.concat(danger_rows, ignore_index=True)

        st.markdown('<div class="section-title">위험 데이터 요약</div>', unsafe_allow_html=True)
        st.dataframe(danger_summary_df.style.apply(highlight_risk, axis=1), use_container_width=True)

        danger_csv = danger_summary_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            label="danger_summary.csv 다운로드",
            data=danger_csv,
            file_name="danger_summary.csv",
            mime="text/csv",
            key=f"danger_download_{key_suffix}"
        )
    else:
        st.info("위험 데이터가 없습니다.")

    st.markdown('<div class="section-title">전체 파일 통합 결과</div>', unsafe_allow_html=True)
    st.dataframe(final_result_df.style.apply(highlight_risk, axis=1), use_container_width=True)

    result_csv = final_result_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label="all_results.csv 다운로드",
        data=result_csv,
        file_name="all_results.csv",
        mime="text/csv",
        key=f"all_results_download_{key_suffix}"
    )


# -----------------------------
# 8. 관리자 통합 조회 화면
# -----------------------------
def admin_all_users_section():
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">관리자 전체 사용자 통합 조회</div>', unsafe_allow_html=True)

    all_user_names = sorted([
        name for name, info in USERS.items()
        if info["role"] in ["staff", "client", "admin"]
    ])

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
        placeholder="예: client1 또는 20260320 또는 7기둥",
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

    col_admin1, col_admin2 = st.columns([1, 5])

    with col_admin1:
        admin_load_clicked = st.button("관리자 불러오기", use_container_width=True)

    if admin_load_clicked:
        if selected_admin_file == "선택 안 함":
            st.warning("불러올 파일을 먼저 선택해주세요.")
        else:
            selected_item = option_map[selected_admin_file]
            selected_path = selected_item["path"]
            selected_user = selected_item["username"]
            selected_name = selected_item["name"]

            saved_df = load_saved_csv(selected_path)

            if saved_df is not None:
                st.success(f"관리자 불러오기 완료: [{selected_user}] {selected_name}")
                analyze_dataframes(
                    [(f"[{selected_user}] {selected_name}", saved_df)],
                    key_suffix="admin_all"
                )


# -----------------------------
# 9. 메인 앱
# -----------------------------
def main_app():
    st.markdown("""
        <style>
        .main-title {
            font-size: 48px;
            font-weight: 800;
            color: #1f2c44;
            margin-bottom: 5px;
        }
        .sub-text {
            font-size: 18px;
            color: #555;
            margin-bottom: 25px;
        }
        .section-title {
            font-size: 28px;
            font-weight: 700;
            color: #1f2c44;
            margin-top: 30px;
            margin-bottom: 10px;
        }
        .kpi-card {
            padding: 20px;
            border-radius: 16px;
            background-color: #f7f9fc;
            border: 1px solid #e6eaf0;
            text-align: center;
            font-size: 20px;
            font-weight: 700;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }
        .kpi-red {
            color: #d62828;
        }
        .kpi-yellow {
            color: #d4a000;
        }
        .kpi-green {
            color: #2a9d2f;
        }
        .alert-high-risk {
            background-color: #ffe5e5;
            border: 2px solid #d62828;
            color: #a61b1b;
            padding: 18px;
            border-radius: 14px;
            font-size: 20px;
            font-weight: 800;
            margin-top: 20px;
            margin-bottom: 20px;
        }
        .alert-medium-risk {
            background-color: #fff6db;
            border: 2px solid #d4a000;
            color: #9a7500;
            padding: 16px;
            border-radius: 14px;
            font-size: 18px;
            font-weight: 700;
            margin-top: 15px;
            margin-bottom: 20px;
        }
        div[data-testid="stButton"] > button[kind="secondary"] {
            min-height: 90px;
            font-size: 46px;
            font-weight: 800;
            border-radius: 18px;
        }
        hr {
            margin-top: 20px;
            margin-bottom: 20px;
        }
        </style>
    """, unsafe_allow_html=True)

    col_logo, col_title, col_user = st.columns([1, 4, 1.2])

    with col_logo:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=120)
            if st.button("홈", key="home_button_under_logo", use_container_width=True):
                st.session_state.selected_saved_file = "선택 안 함"
                st.session_state.selected_admin_user = "전체"
                st.rerun()
        else:
            if st.button("YW", key="home_button", use_container_width=True):
                st.session_state.selected_saved_file = "선택 안 함"
                st.session_state.selected_admin_user = "전체"
                st.rerun()

    with col_title:
        st.markdown('<div class="main-title">화재 위험도 자동 분석 프로그램</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-text">(주)윤우테크 | CSV 파일을 업로드하면 위험도와 판정을 자동 계산합니다.</div>', unsafe_allow_html=True)

    with col_user:
        st.write(f"환영합니다, **{st.session_state.username}** 님")
        st.write(f"권한: **{get_role_name(st.session_state.role)}**")
        if st.button("로그아웃"):
            logout()

    if st.session_state.role == "admin":
        st.info("관리자 계정입니다. 사용자별 데이터 관리 및 전체 운영에 사용할 수 있습니다.")
        admin_dashboard()

    st.markdown('<div class="section-title">저장된 파일 불러오기</div>', unsafe_allow_html=True)

    all_saved_files = get_saved_files(st.session_state.username)

    search_text = st.text_input("파일명 검색", placeholder="예: 7기둥 또는 20260320")

    if search_text.strip():
        searched_files = [f for f in all_saved_files if search_text.lower() in f.name.lower()]
    else:
        searched_files = all_saved_files

    filtered_files = searched_files[:20]
    file_options = ["선택 안 함"] + [file.name for file in filtered_files]

    if len(filtered_files) == 0:
        st.info("검색 조건에 맞는 저장 파일이 없습니다.")
    else:
        st.caption(f"검색 결과 {len(searched_files)}개 중 최근 20개 표시")

        selected_saved_file = st.selectbox(
            "저장된 CSV 선택",
            options=file_options,
            key="selected_saved_file"
        )

        col_load1, col_load2 = st.columns([1, 5])

        with col_load1:
            load_clicked = st.button("불러오기", use_container_width=True)

        if load_clicked:
            if selected_saved_file == "선택 안 함":
                st.warning("불러올 파일을 먼저 선택해주세요.")
            else:
                selected_path = next(
                    (file for file in filtered_files if file.name == selected_saved_file),
                    None
                )

                if selected_path:
                    saved_df = load_saved_csv(selected_path)

                    if saved_df is not None:
                        st.success(f"저장 파일 불러오기 완료: {selected_saved_file}")
                        analyze_dataframes([(selected_saved_file, saved_df)], key_suffix="manual")

    if st.session_state.role == "admin":
        admin_all_users_section()

    st.markdown("<hr>", unsafe_allow_html=True)

    st.markdown('<div class="section-title">새 CSV 파일 업로드</div>', unsafe_allow_html=True)

    if st.session_state.role != "client":
        uploaded_files = st.file_uploader(
            "CSV 파일 여러 개 선택",
            type=["csv"],
            accept_multiple_files=True
        )

        if uploaded_files:
            all_results = []
            saved_file_paths = []

            for uploaded_file in uploaded_files:
                saved_path = save_uploaded_file(uploaded_file, st.session_state.username)
                saved_file_paths.append(saved_path)

                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, encoding="utf-8")
                except:
                    try:
                        uploaded_file.seek(0)
                        df = pd.read_csv(uploaded_file, encoding="cp949")
                    except Exception as e:
                        st.error(f"{uploaded_file.name} 파일을 읽는 중 오류가 발생했습니다: {e}")
                        continue

                all_results.append((uploaded_file.name, df))

            if saved_file_paths:
                st.success("업로드 파일이 자동 저장되었습니다.")
                with st.expander("저장된 파일 위치 보기"):
                    for path in saved_file_paths:
                        st.write(path)

            if all_results:
                analyze_dataframes(all_results, key_suffix="upload")
        else:
            st.info("업로드할 CSV 파일을 선택해주세요.")
    else:
        st.info("고객 계정은 업로드 기능이 제한됩니다. 저장된 파일 조회만 가능합니다.")


# -----------------------------
# 10. 실행
# -----------------------------
if st.session_state.logged_in:
    main_app()
else:
    login()