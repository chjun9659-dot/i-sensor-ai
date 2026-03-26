import os
from io import BytesIO
from datetime import datetime
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# -----------------------------
# 기본 설정
# -----------------------------
st.set_page_config(page_title="화재 위험도 자동 분석 프로그램", layout="wide")

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
# 세션 상태 초기화
# -----------------------------
DEFAULT_SESSION = {
    "logged_in": False,
    "username": "",
    "role": "",
    "user_complex": "전체",
    "dashboard_filter": "전체",
}

for key, value in DEFAULT_SESSION.items():
    if key not in st.session_state:
        st.session_state[key] = value

# -----------------------------
# 색상 테마
# -----------------------------
COLOR_MAP = {
    "위험": "#E74C3C",
    "주의": "#F4B400",
    "정상": "#2ECC71",
    "전체": "#5B6C8F",
}

CARD_STYLE = {
    "bg": "#F8FAFC",
    "border": "#E2E8F0",
    "title": "#1F2A44",
    "sub": "#64748B"
}

# -----------------------------
# 공통 UI 함수
# -----------------------------
def show_logo():
    if os.path.exists(LOGO_PATH):
        col1, col2, col3 = st.columns([3, 2, 3])
        with col2:
            st.image(LOGO_PATH, width=170)

def show_top_banner():
    st.markdown(
        """
        <div style="
            background: linear-gradient(135deg, #1F2A44 0%, #2E5BBA 100%);
            padding: 22px 26px;
            border-radius: 20px;
            margin-bottom: 16px;
            color: white;
            box-shadow: 0 8px 24px rgba(31,42,68,0.18);
        ">
            <div style="font-size: 30px; font-weight: 800; margin-bottom: 6px;">
                화재 위험도 자동 분석 프로그램
            </div>
            <div style="font-size: 15px; opacity: 0.92;">
                윤우테크 / 아이센서 분석 대시보드
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

def metric_card(title, value, icon="", color="#5B6C8F", subtitle=""):
    st.markdown(
        f"""
        <div style="
            background: {CARD_STYLE['bg']};
            border: 1px solid {CARD_STYLE['border']};
            border-left: 8px solid {color};
            border-radius: 18px;
            padding: 20px 22px;
            min-height: 125px;
            box-shadow: 0 4px 14px rgba(15,23,42,0.05);
        ">
            <div style="font-size: 17px; font-weight: 700; color: {CARD_STYLE['title']};">
                {icon} {title}
            </div>
            <div style="font-size: 38px; font-weight: 900; margin-top: 10px; color: {color};">
                {value}
            </div>
            <div style="font-size: 13px; color: {CARD_STYLE['sub']}; margin-top: 4px;">
                {subtitle}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

def section_title(text):
    st.markdown(
        f"""
        <div style="
            margin-top: 10px;
            margin-bottom: 10px;
            font-size: 22px;
            font-weight: 800;
            color: #1F2A44;
        ">
            {text}
        </div>
        """,
        unsafe_allow_html=True
    )

def info_box(text):
    st.markdown(
        f"""
        <div style="
            background: #EFF6FF;
            border: 1px solid #BFDBFE;
            color: #1D4ED8;
            border-radius: 14px;
            padding: 14px 16px;
            margin-bottom: 14px;
            font-weight: 600;
        ">
            {text}
        </div>
        """,
        unsafe_allow_html=True
    )

# -----------------------------
# 공통 기능 함수
# -----------------------------
def get_role_name(role):
    if role == "admin":
        return "관리자"
    elif role == "staff":
        return "행정"
    elif role == "client":
        return "고객"
    return role

def login():
    show_logo()
    show_top_banner()

    st.markdown("### 🔐 로그인")
    st.write("아이디와 비밀번호를 입력하세요.")

    username = st.text_input("아이디")
    password = st.text_input("비밀번호", type="password")

    if st.button("로그인", use_container_width=True):
        if username in USERS and USERS[username]["password"] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.role = USERS[username]["role"]
            st.session_state.user_complex = USERS[username]["complex"]
            st.success("로그인 성공")
            st.rerun()
        else:
            st.error("아이디 또는 비밀번호가 올바르지 않습니다.")

def logout():
    for key, value in DEFAULT_SESSION.items():
        st.session_state[key] = value
    st.rerun()

def get_user_folder(username):
    user_folder = DATA_FOLDER / username
    user_folder.mkdir(parents=True, exist_ok=True)
    return user_folder

def normalize_complex_column(df):
    df = df.copy()
    if "단지명" not in df.columns:
        df["단지명"] = "전체"
    df["단지명"] = df["단지명"].fillna("전체").astype(str)
    return df

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

    selected_complex = "전체"
    selected_dong = "전체"
    selected_floor = "전체"
    selected_area = "전체"

    if st.session_state.role == "client" and st.session_state.user_complex != "전체":
        info_box(f"현재 계정은 <b>{st.session_state.user_complex}</b> 단지만 조회 가능합니다.")
    else:
        complex_list = ["전체"] + sorted([str(x) for x in filtered_df["단지명"].dropna().unique()])
        with col0:
            selected_complex = st.selectbox("단지 선택", complex_list, key=f"{key_prefix}_complex")

    if "동" in filtered_df.columns:
        dong_list = ["전체"] + sorted([str(x) for x in filtered_df["동"].dropna().unique()])
        with col1:
            selected_dong = st.selectbox("동 선택", dong_list, key=f"{key_prefix}_dong")

    if "층" in filtered_df.columns:
        floor_list = ["전체"] + sorted([str(x) for x in filtered_df["층"].dropna().unique()])
        with col2:
            selected_floor = st.selectbox("층 선택", floor_list, key=f"{key_prefix}_floor")

    if "구역" in filtered_df.columns:
        area_list = ["전체"] + sorted([str(x) for x in filtered_df["구역"].dropna().unique()])
        with col3:
            selected_area = st.selectbox("구역 선택", area_list, key=f"{key_prefix}_area")

    if st.session_state.role != "client" and selected_complex != "전체":
        filtered_df = filtered_df[filtered_df["단지명"].astype(str) == selected_complex]

    if "동" in filtered_df.columns and selected_dong != "전체":
        filtered_df = filtered_df[filtered_df["동"].astype(str) == selected_dong]

    if "층" in filtered_df.columns and selected_floor != "전체":
        filtered_df = filtered_df[filtered_df["층"].astype(str) == selected_floor]

    if "구역" in filtered_df.columns and selected_area != "전체":
        filtered_df = filtered_df[filtered_df["구역"].astype(str) == selected_area]

    return filtered_df

def save_result(df, username):
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"result_{now}.csv"
    filepath = get_user_folder(username) / filename
    df.to_csv(filepath, index=False, encoding="utf-8-sig")
    return filename, filepath

def get_saved_files(username):
    user_folder = get_user_folder(username)
    files = sorted(user_folder.glob("*.csv"), key=lambda x: x.stat().st_mtime, reverse=True)
    return files

def load_saved_file(file_path):
    try:
        return pd.read_csv(file_path)
    except:
        return None

def get_all_saved_files():
    results = []
    for username in USERS.keys():
        user_folder = get_user_folder(username)
        for file_path in user_folder.glob("*.csv"):
            results.append({
                "username": username,
                "name": file_path.name,
                "path": file_path,
                "mtime": file_path.stat().st_mtime
            })
    results.sort(key=lambda x: x["mtime"], reverse=True)
    return results

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
        except:
            pass

    df.to_csv(log_file, index=False, encoding="utf-8-sig")

def load_risk_log():
    log_file = LOG_FOLDER / "risk_log.csv"
    if log_file.exists():
        try:
            return pd.read_csv(log_file)
        except:
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
    first = danger_df.iloc[0]

    dong_text = first.get("동", "")
    floor_text = first.get("층", "")
    area_text = first.get("구역", "")

    location_parts = []
    for x in [dong_text, floor_text, area_text]:
        value = str(x).strip()
        if value not in ["", "-", "nan", "None"]:
            location_parts.append(value)

    location_html = ""
    if location_parts:
        location_html = f"<br>📍 위치: {' / '.join(location_parts)}"

    st.markdown(
        f"""
        <div style="
            background: linear-gradient(135deg, #FDECEC 0%, #FFF2F2 100%);
            border: 2px solid #E74C3C;
            border-radius: 18px;
            padding: 20px;
            margin-bottom: 18px;
            text-align: center;
            font-size: 22px;
            font-weight: 800;
            color: #B42318;
            box-shadow: 0 8px 20px rgba(231,76,60,0.10);
        ">
            🚨 위험 데이터 {danger_count}건 발생
            {location_html}
        </div>
        """,
        unsafe_allow_html=True
    )

def create_test_danger_data():
    test_df = pd.DataFrame([
        {"단지명": "무등산자이", "최고 온도": 15, "차량 감지": 1, "이벤트 종류": 25, "동": "101동", "층": "지하2층", "구역": "A구역"},
        {"단지명": "무등산자이", "최고 온도": 13, "차량 감지": 1, "이벤트 종류": 21, "동": "101동", "층": "지하1층", "구역": "B구역"},
        {"단지명": "무등산자이", "최고 온도": 11, "차량 감지": 1, "이벤트 종류": 22, "동": "102동", "층": "지하2층", "구역": "A구역"},
        {"단지명": "센트럴파크", "최고 온도": 8, "차량 감지": 0, "이벤트 종류": 5, "동": "102동", "층": "지상1층", "구역": "C구역"},
        {"단지명": "센트럴파크", "최고 온도": 9, "차량 감지": 1, "이벤트 종류": 10, "동": "103동", "층": "지하3층", "구역": "D구역"},
        {"단지명": "센트럴파크", "최고 온도": 14, "차량 감지": 1, "이벤트 종류": 30, "동": "103동", "층": "지하2층", "구역": "충전구역3번"}
    ])
    return process_dataframe(test_df, "test_data.csv")

def register_korean_font():
    font_candidates = [
        "NanumGothic.ttf",
        "NotoSansKR-Regular.ttf",
        "Malgun.ttf",
        "malgun.ttf",
    ]

    for font_path in font_candidates:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont("KoreanFont", font_path))
                return "KoreanFont"
            except:
                pass

    return "Helvetica"

def generate_pdf_bytes(df, apt_name="단지"):
    if df is None or df.empty:
        return None

    pdf_df = df.copy()

    if "판정" not in pdf_df.columns and "위험도" in pdf_df.columns:
        pdf_df["판정"] = pdf_df["위험도"].apply(classify_risk)

    font_name = register_korean_font()

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=20,
        leftMargin=20,
        topMargin=30,
        bottomMargin=20
    )

    styles = getSampleStyleSheet()
    styles["Title"].fontName = font_name
    styles["Normal"].fontName = font_name
    styles["Title"].fontSize = 16
    styles["Normal"].fontSize = 10

    elements = []

    elements.append(Paragraph(f"{apt_name} 위험도 분석 리포트", styles["Title"]))
    elements.append(Spacer(1, 16))

    total = len(pdf_df)
    danger = len(pdf_df[pdf_df["판정"] == "위험"]) if "판정" in pdf_df.columns else 0
    warning = len(pdf_df[pdf_df["판정"] == "주의"]) if "판정" in pdf_df.columns else 0
    normal = len(pdf_df[pdf_df["판정"] == "정상"]) if "판정" in pdf_df.columns else 0

    elements.append(Paragraph(f"전체: {total}건", styles["Normal"]))
    elements.append(Paragraph(f"위험: {danger}건", styles["Normal"]))
    elements.append(Paragraph(f"주의: {warning}건", styles["Normal"]))
    elements.append(Paragraph(f"정상: {normal}건", styles["Normal"]))
    elements.append(Spacer(1, 16))

    display_df = pdf_df.copy().fillna("").astype(str)

    max_rows = 30
    if len(display_df) > max_rows:
        display_df = display_df.head(max_rows)

    table_data = [display_df.columns.tolist()] + display_df.values.tolist()

    col_count = len(display_df.columns)
    available_width = 555
    col_width = max(45, int(available_width / max(col_count, 1)))
    col_widths = [col_width] * col_count

    table = Table(table_data, repeatRows=1, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#34495E")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("LEADING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))

    elements.append(table)

    if len(pdf_df) > max_rows:
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"* 표에는 상위 {max_rows}건만 표시되었습니다.", styles["Normal"]))

    doc.build(elements)
    pdf_data = buffer.getvalue()
    buffer.close()
    return pdf_data

# -----------------------------
# 차트 / 표 스타일 함수
# -----------------------------
def make_status_chart(df):
    if df is None or df.empty or "판정" not in df.columns:
        return None

    order = ["정상", "주의", "위험"]
    chart_df = df["판정"].value_counts().reindex(order, fill_value=0).reset_index()
    chart_df.columns = ["판정", "건수"]

    chart = (
        alt.Chart(chart_df)
        .mark_bar(cornerRadiusTopLeft=10, cornerRadiusTopRight=10)
        .encode(
            x=alt.X("판정:N", sort=order, title="판정"),
            y=alt.Y("건수:Q", title="건수"),
            color=alt.Color(
                "판정:N",
                scale=alt.Scale(
                    domain=["정상", "주의", "위험"],
                    range=[COLOR_MAP["정상"], COLOR_MAP["주의"], COLOR_MAP["위험"]]
                ),
                legend=None
            ),
            tooltip=["판정", "건수"]
        )
        .properties(height=340)
    )
    return chart

def make_complex_chart(df):
    if df is None or df.empty or "판정" not in df.columns or "단지명" not in df.columns:
        return None

    chart_df = (
        df.groupby(["단지명", "판정"])
        .size()
        .reset_index(name="건수")
    )

    chart = (
        alt.Chart(chart_df)
        .mark_bar(cornerRadiusTopLeft=8, cornerRadiusTopRight=8)
        .encode(
            x=alt.X("단지명:N", title="단지명"),
            y=alt.Y("건수:Q", title="건수"),
            color=alt.Color(
                "판정:N",
                scale=alt.Scale(
                    domain=["정상", "주의", "위험"],
                    range=[COLOR_MAP["정상"], COLOR_MAP["주의"], COLOR_MAP["위험"]]
                ),
                title="판정"
            ),
            tooltip=["단지명", "판정", "건수"]
        )
        .properties(height=340)
    )
    return chart

def style_dataframe(df):
    def highlight_row(row):
        result = row.get("판정", "")
        if result == "위험":
            return ["background-color: #FDECEC"] * len(row)
        elif result == "주의":
            return ["background-color: #FFF8E1"] * len(row)
        elif result == "정상":
            return ["background-color: #ECFDF3"] * len(row)
        return [""] * len(row)

    return df.style.apply(highlight_row, axis=1)

# -----------------------------
# 결과 표시 섹션
# -----------------------------
def render_result_section(df, key_prefix="result", file_name="분석결과.csv"):
    if df is None or df.empty:
        st.warning("표시할 데이터가 없습니다.")
        return

    scoped_df = apply_user_complex_scope(df)
    show_danger_alert(scoped_df)

    section_title("분석 결과")
    filtered_df = apply_complex_filter(scoped_df, key_prefix=key_prefix)

    if filtered_df.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
        return

    total_count = len(scoped_df)
    danger_count = len(scoped_df[scoped_df["판정"] == "위험"])
    warning_count = len(scoped_df[scoped_df["판정"] == "주의"])
    normal_count = len(scoped_df[scoped_df["판정"] == "정상"])

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("전체 데이터", total_count, icon="📁", color=COLOR_MAP["전체"], subtitle="현재 범위 기준")
    with c2:
        metric_card("위험", danger_count, icon="🔴", color=COLOR_MAP["위험"], subtitle="즉시 확인 권장")
    with c3:
        metric_card("주의", warning_count, icon="🟡", color=COLOR_MAP["주의"], subtitle="관찰 필요")
    with c4:
        metric_card("정상", normal_count, icon="🟢", color=COLOR_MAP["정상"], subtitle="안정 상태")

    st.markdown("")

    ch1, ch2 = st.columns(2)

    with ch1:
        section_title("판정별 건수")
        status_chart = make_status_chart(scoped_df)
        if status_chart is not None:
            st.altair_chart(status_chart, use_container_width=True)

    with ch2:
        if "단지명" in scoped_df.columns:
            section_title("단지별 현황")
            complex_chart = make_complex_chart(scoped_df)
            if complex_chart is not None:
                st.altair_chart(complex_chart, use_container_width=True)

    section_title("상세 데이터")
    st.dataframe(style_dataframe(filtered_df), use_container_width=True, height=500)

    download_name_base = Path(file_name).stem
    pdf_complex_name = st.session_state.user_complex if st.session_state.user_complex != "전체" else "전체"

    btn1, btn2 = st.columns(2)

    with btn1:
        csv_data = filtered_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            label="📥 현재 화면 결과 CSV 다운로드",
            data=csv_data,
            file_name=file_name,
            mime="text/csv",
            use_container_width=True,
            key=f"download_csv_{key_prefix}"
        )

    with btn2:
        pdf_data = generate_pdf_bytes(filtered_df, pdf_complex_name)
        if pdf_data:
            st.download_button(
                label="📄 현재 화면 PDF 리포트 다운로드",
                data=pdf_data,
                file_name=f"{download_name_base}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key=f"download_pdf_{key_prefix}"
            )

# -----------------------------
# 관리자 대시보드
# -----------------------------
def admin_dashboard(df):
    st.title("관리자 대시보드")

    if df is None or df.empty:
        st.warning("대시보드에 표시할 데이터가 없습니다.")
        return

    df = normalize_complex_column(df)
    df = apply_user_complex_scope(df)

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
    total_users = len(USERS)

    if danger_count > 0:
        first = danger_df.iloc[0]

        dong_text = first.get("동", "")
        floor_text = first.get("층", "")
        area_text = first.get("구역", "")

        location_parts = []
        for x in [dong_text, floor_text, area_text]:
            value = str(x).strip()
            if value not in ["", "-", "nan", "None"]:
                location_parts.append(value)

        location_html = ""
        if location_parts:
            location_html = f"<br>📍 위치: {' / '.join(location_parts)}"

        st.markdown(
            f"""
            <div style="
                background: linear-gradient(135deg, #FDECEC 0%, #FFF3F2 100%);
                border: 2px solid #E74C3C;
                border-radius: 18px;
                padding: 20px;
                margin-bottom: 20px;
                text-align: center;
                font-size: 22px;
                font-weight: 800;
                color: #B42318;
                box-shadow: 0 8px 20px rgba(231,76,60,0.10);
            ">
                🚨 위험 데이터 {danger_count}건 발생
                {location_html}
            </div>
            """,
            unsafe_allow_html=True
        )

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        metric_card("전체 사용자", total_users, icon="👥", color="#4F46E5", subtitle="등록 계정 수")
    with c2:
        metric_card("전체 데이터", total_rows, icon="📁", color=COLOR_MAP["전체"], subtitle="현재 파일 기준")
    with c3:
        metric_card("위험", danger_count, icon="🔴", color=COLOR_MAP["위험"], subtitle="즉시 확인")
    with c4:
        metric_card("주의", warning_count, icon="🟡", color=COLOR_MAP["주의"], subtitle="관찰 필요")
    with c5:
        metric_card("정상", normal_count, icon="🟢", color=COLOR_MAP["정상"], subtitle="안정 상태")

    st.write("")

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
    info_box(f"현재 필터: <b>{st.session_state.dashboard_filter}</b>")

    current_filter = st.session_state.dashboard_filter

    if current_filter == "위험":
        filtered_df = danger_df
    elif current_filter == "주의":
        filtered_df = warning_df
    elif current_filter == "정상":
        filtered_df = normal_df
    else:
        filtered_df = df

    chart_left, chart_right = st.columns(2)

    with chart_left:
        section_title("대시보드 판정별 건수")
        status_chart = make_status_chart(df)
        if status_chart is not None:
            st.altair_chart(status_chart, use_container_width=True)

    with chart_right:
        if "단지명" in df.columns:
            section_title("대시보드 단지별 현황")
            complex_chart = make_complex_chart(df)
            if complex_chart is not None:
                st.altair_chart(complex_chart, use_container_width=True)

    st.markdown("---")
    st.subheader("필터 결과")
    filtered_df = apply_complex_filter(filtered_df, key_prefix="dashboard")

    if filtered_df.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
    else:
        st.dataframe(style_dataframe(filtered_df), use_container_width=True, height=450)

        pdf_data = generate_pdf_bytes(filtered_df, "관리자대시보드")
        if pdf_data:
            st.download_button(
                label="📄 대시보드 PDF 리포트 다운로드",
                data=pdf_data,
                file_name="관리자대시보드_리포트.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="dashboard_pdf_download"
            )

    st.markdown("---")
    st.subheader("📜 위험 발생 이력")

    log_df = load_risk_log()
    if log_df is not None and not log_df.empty:
        log_df = apply_user_complex_scope(log_df)
        log_df = apply_complex_filter(log_df, key_prefix="dashboard_log")
        st.dataframe(style_dataframe(log_df), use_container_width=True, height=300)
    else:
        st.info("저장된 위험 이력이 없습니다.")

# -----------------------------
# 관리자 전체 사용자 통합 조회
# -----------------------------
def admin_all_users_section():
    if st.session_state.role != "admin":
        return

    st.subheader("관리자 전체 사용자 통합 조회")

    all_user_names = sorted(USERS.keys())
    selected_admin_user = st.selectbox("조회할 사용자 선택", ["전체"] + all_user_names)

    all_files = get_all_saved_files()

    if selected_admin_user != "전체":
        all_files = [f for f in all_files if f["username"] == selected_admin_user]

    search_text = st.text_input("파일명 검색", placeholder="예: 무등산자이 또는 20260325")

    if search_text.strip():
        search_lower = search_text.lower()
        all_files = [
            f for f in all_files
            if search_lower in f["name"].lower() or search_lower in f["username"].lower()
        ]

    if not all_files:
        st.info("조건에 맞는 저장 파일이 없습니다.")
        return

    display_options = ["선택 안 함"]
    option_map = {}

    for item in all_files[:50]:
        display_name = f"[{item['username']}] {item['name']}"
        display_options.append(display_name)
        option_map[display_name] = item

    selected_file = st.selectbox("전체 사용자 CSV 선택", display_options, key="admin_all_file")

    if st.button("관리자 불러오기", use_container_width=True):
        if selected_file == "선택 안 함":
            st.warning("불러올 파일을 먼저 선택해주세요.")
        else:
            selected_item = option_map[selected_file]
            saved_df = load_saved_file(selected_item["path"])

            if saved_df is not None:
                st.success(f"불러오기 완료: {selected_file}")
                render_result_section(
                    saved_df,
                    key_prefix="admin_all",
                    file_name=selected_item["name"]
                )
            else:
                st.error("파일을 불러오지 못했습니다.")

# -----------------------------
# 메인 화면
# -----------------------------
def main():
    if not st.session_state.logged_in:
        login()
        return

    show_logo()
    show_top_banner()

    col1, col2 = st.columns([6, 1])
    with col1:
        st.caption(
            f"로그인: {st.session_state.username} | 권한: {get_role_name(st.session_state.role)}"
        )

    with col2:
        if st.button("로그아웃", use_container_width=True):
            logout()

    st.write("CSV 파일을 업로드하면 위험도와 판정을 자동 계산합니다.")

    menu_options = ["데이터 분석", "저장 파일 보기", "관리자 대시보드"]
    if st.session_state.role == "admin":
        menu_options.append("관리자 통합 조회")

    menu = st.sidebar.radio("메뉴 선택", menu_options)

    if menu == "데이터 분석":
        st.subheader("CSV 업로드 분석")

        if st.session_state.role == "client" and st.session_state.user_complex != "전체":
            info_box(f"현재 계정은 <b>{st.session_state.user_complex}</b> 단지 전용 계정입니다.")

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

                danger_only = test_df[test_df["판정"] == "위험"].copy()
                if not danger_only.empty:
                    save_risk_log(danger_only, st.session_state.username)

                filename, _ = save_result(test_df, st.session_state.username)
                st.success(f"테스트 데이터 저장 완료: {filename}")
                render_result_section(test_df, key_prefix="test_data", file_name=filename)

        with col_btn2:
            if st.button("📋 테스트 데이터 설명 보기", use_container_width=True):
                st.info("위험/주의/정상 데이터가 섞인 샘플 데이터를 자동 생성합니다.")

        if uploaded_files:
            all_results = []

            for uploaded_file in uploaded_files:
                try:
                    df = pd.read_csv(uploaded_file)
                    result_df = process_dataframe(df, uploaded_file.name)

                    if st.session_state.role == "client" and st.session_state.user_complex != "전체":
                        result_df = result_df[result_df["단지명"].astype(str) == st.session_state.user_complex]

                    if not result_df.empty:
                        all_results.append(result_df)

                except Exception as e:
                    st.error(f"{uploaded_file.name} 처리 중 오류 발생: {e}")

            if all_results:
                final_df = pd.concat(all_results, ignore_index=True)

                danger_only = final_df[final_df["판정"] == "위험"].copy()
                if not danger_only.empty:
                    save_risk_log(danger_only, st.session_state.username)

                render_result_section(
                    final_df,
                    key_prefix="analysis_result",
                    file_name="분석결과_통합.csv"
                )

                if st.button("결과 저장", use_container_width=True):
                    filename, _ = save_result(final_df, st.session_state.username)
                    st.success(f"저장 완료: {filename}")
            else:
                st.warning("처리 가능한 데이터가 없습니다.")

    elif menu == "저장 파일 보기":
        st.subheader("저장된 결과 파일")

        saved_files = get_saved_files(st.session_state.username)

        if not saved_files:
            st.info("저장된 파일이 없습니다.")
        else:
            file_names = [f.name for f in saved_files]
            selected_file_name = st.selectbox("파일 선택", file_names)

            selected_path = None
            for f in saved_files:
                if f.name == selected_file_name:
                    selected_path = f
                    break

            if selected_path is not None:
                loaded_df = load_saved_file(selected_path)

                if loaded_df is not None:
                    st.success(f"불러온 파일: {selected_file_name}")
                    render_result_section(
                        loaded_df,
                        key_prefix="saved_file",
                        file_name=selected_file_name
                    )
                else:
                    st.error("파일을 불러오지 못했습니다.")

    elif menu == "관리자 대시보드":
        if st.session_state.role == "admin":
            all_files = get_all_saved_files()
        else:
            all_files = []
            for f in get_saved_files(st.session_state.username):
                all_files.append({
                    "username": st.session_state.username,
                    "name": f.name,
                    "path": f,
                    "mtime": f.stat().st_mtime
                })

        if not all_files:
            st.warning("저장된 파일이 없어 대시보드에 표시할 데이터가 없습니다.")
        else:
            latest_path = all_files[0]["path"]
            dashboard_df = load_saved_file(latest_path)

            if dashboard_df is not None:
                dashboard_df = process_dataframe(dashboard_df)
                st.caption(f"기준 파일: {all_files[0]['name']}")
                admin_dashboard(dashboard_df)
            else:
                st.error("대시보드용 데이터를 불러오지 못했습니다.")

    elif menu == "관리자 통합 조회":
        admin_all_users_section()

# -----------------------------
# 실행
# -----------------------------
if __name__ == "__main__":
    main()


