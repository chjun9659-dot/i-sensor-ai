# 연차관리 전용 모듈: app.py를 import하지 않고 의존성을 전달받습니다.
import os
import re
import csv
import uuid
from datetime import datetime, date, timedelta
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

def _set_dependencies(dependencies):
    required = ['apply_role_filter', 'clean_text', 'get_current_sheet_urls', 'get_gsheet_client', 'load_df', 'load_google_sheet_data', 'load_users_from_gsheet', 'to_excel_bytes', 'ui_card']
    missing = [name for name in required if name not in dependencies]
    if missing:
        raise RuntimeError(f"연차관리 의존성 누락: {missing}")
    for name in required:
        globals()[name] = dependencies[name]

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

        url = get_current_sheet_urls().get("연차관리", "")
        if not url:
            raise ValueError("연차관리 구글시트 URL이 없습니다.")
        sh = get_gsheet_client().open_by_url(url)
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
    df = df.astype(object)
    
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

    # ✅ 연차 관련 컬럼을 모두 object 타입으로 강제
    safe_cols = [
        "이름",
        "입사일",
        "기산시작일",
        "기산종료일",
        "근속년수",
        "발생 연차",
        "사용 연차",
        "잔여 연차",
    ]

    for col in safe_cols:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(object)

    today = date.today()
    changed = False
    changed_names = []

    for idx in df.index:
        end_date = pd.to_datetime(
            df.loc[idx, "기산종료일"],
            errors="coerce"
        )

        if pd.isna(end_date):
            continue

        if today > end_date.date():
            hire_date = pd.to_datetime(
                df.loc[idx, "입사일"],
                errors="coerce"
            )

            if pd.isna(hire_date):
                continue

            start_date, new_end_date, service_years, leave_days = \
                calculate_auto_leave_days(hire_date.date())

            df.loc[idx, "기산시작일"] = str(start_date)
            df.loc[idx, "기산종료일"] = str(new_end_date)

            # ✅ 숫자도 저장 시 안전하게 문자열로 처리
            df.loc[idx, "근속년수"] = str(int(service_years))
            df.loc[idx, "발생 연차"] = format_leave_number(leave_days)

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
    st.markdown(
        '<div class="erp-section-title">직원별 연차 요약 카드</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="erp-section-desc">전체 직원 기준</div>',
        unsafe_allow_html=True
    )

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

def has_registered_leave_on_date(row, use_date):
    """같은 날짜의 연차/반차 중복 등록을 방지합니다."""
    target = pd.Timestamp(use_date).date()
    for col in USE_COLS:
        parsed_date, _ = parse_use_entry(row.get(col, None))
        if parsed_date is not None and parsed_date.date() == target:
            return True
    return False


def backup_sheet_before_write(values):
    """쓰기 직전 구글시트 원본을 로컬 CSV로 보존합니다."""
    os.makedirs(VACATION_BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    path = os.path.join(VACATION_BACKUP_DIR, f"연차관리_저장전_{stamp}_{uuid.uuid4().hex[:6]}.csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerows(values)
    return path


def save_vacation_data(df):
    sheet_urls = get_current_sheet_urls()
    url = sheet_urls.get("연차관리", "")

    if not url:
        raise Exception("연차관리 구글시트 URL이 없습니다.")

    client = get_gsheet_client()
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    if not match:
        raise ValueError("연차관리 구글시트 URL 형식이 올바르지 않습니다.")
    sheet_id = match.group(1)
    spreadsheet = client.open_by_key(sheet_id)
    worksheet = spreadsheet.get_worksheet(0)

    save_df = df.copy()

    if "row_id" in save_df.columns:
        save_df = save_df.drop(columns=["row_id"])

    values = worksheet.get_all_values()
    if not values:
        raise Exception("연차관리 시트에 헤더가 없습니다.")

    raw_headers = [str(x).strip() for x in values[0]]
    if not raw_headers or any(not h for h in raw_headers):
        raise ValueError("연차관리 시트 헤더에 빈 칸이 있어 저장을 중단했습니다.")
    if len(set(raw_headers)) != len(raw_headers):
        raise ValueError("연차관리 시트 헤더가 중복되어 저장을 중단했습니다.")
    headers = raw_headers

    # 필요한 열이 누락되면 빈 값으로 덮어쓰지 않고 중단합니다.
    missing_columns = [col for col in headers if col not in save_df.columns]
    if missing_columns:
        raise ValueError(f"저장 데이터에 누락된 열이 있습니다: {missing_columns}")

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

    # 백업에 실패하면 원본 시트도 수정하지 않습니다.
    backup_path = backup_sheet_before_write(values)
    try:
        if rows:
            worksheet.update("A2", rows, value_input_option="USER_ENTERED")

        old_rows = max(0, len(values) - 1)
        new_rows = len(rows)
        if old_rows > new_rows:
            blank_rows = old_rows - new_rows
            clear_start = new_rows + 2
            empty_rows = [[""] * len(headers) for _ in range(blank_rows)]
            worksheet.update(f"A{clear_start}", empty_rows, value_input_option="RAW")
    except Exception as exc:
        raise RuntimeError(
            f"구글시트 저장에 실패했습니다. 저장 전 원본 백업: {backup_path}. "
            "시트가 일부 변경되었을 수 있으니 확인 후 다시 시도하세요."
        ) from exc

    load_google_sheet_data.clear()

def vacation_page(dependencies):
    _set_dependencies(dependencies)

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
        df = load_df("연차관리").copy().astype(object)

        # ✅ 기산종료일 지난 직원만 자동 갱신
        df, vacation_changed, changed_names = refresh_expired_vacation_rows(df)

        if vacation_changed:
            backup_file = create_backup()
            save_vacation_data(df)
            load_google_sheet_data.clear()
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
        st.markdown(
            '<div class="erp-section-title">🛠 관리 도구</div>',
            unsafe_allow_html=True
        )

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
                    load_google_sheet_data.clear()
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
    st.markdown(
        '<div class="erp-section-title">👥 직원 선택</div>',
        unsafe_allow_html=True
    )

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
    st.markdown(
        '<div class="erp-section-title">📌 현재 연차 현황</div>',
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)

    total = to_number(employee["발생 연차"])
    used = to_number(employee["사용 연차"])
    remain = to_number(employee["잔여 연차"])

    with col1:
        ui_card("총 연차", format_leave_number(total), "발생 연차")

    with col2:
        ui_card("사용 연차", format_leave_number(used), "사용 완료")

    with col3:
        ui_card("잔여 연차", format_leave_number(remain), "현재 잔여")

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

            # ✅ 발생연차/기산기간 재계산
            hire_date = pd.to_datetime(df.iloc[row_pos]["입사일"], errors="coerce")

            if pd.notna(hire_date):
                start_date, end_date, service_years, auto_leave_days = calculate_auto_leave_days(
                    hire_date.date()
                )

                df.iloc[row_pos, start_col_pos] = str(start_date)
                df.iloc[row_pos, end_col_pos] = str(end_date)

                service_col_pos = list(df.columns).index("근속년수")
                df.iloc[row_pos, service_col_pos] = int(service_years)

                df.iloc[row_pos, total_col_pos] = float(auto_leave_days)

                total_leave = float(auto_leave_days)

            remain_leave = total_leave - used_leave

            # ✅ 배포앱 안전 처리: df 전체를 object 타입으로
            df = df.astype(object)

            # ✅ 선택 직원 값 저장
            df.iloc[row_pos, used_col_pos] = format_leave_number(used_leave)
            df.iloc[row_pos, remain_col_pos] = format_leave_number(remain_leave)

            # ✅ 구글시트 저장
            save_vacation_data(df)

            save_vacation_log(
                action="재계산",
                target_name=str(selected_name),
                use_date="",
                used_days="",
                reason="",
                note="선택 직원 연차 다시 계산"
            )

            load_google_sheet_data.clear()

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

                if has_registered_leave_on_date(df.loc[idx], use_date):
                    st.error("해당 날짜에 이미 연차 또는 반차가 등록되어 있습니다. 중복 등록할 수 없습니다.")
                elif current_remain < leave_amount:
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

                        load_google_sheet_data.clear()
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

                    if selected_col not in USE_COLS or selected_col not in df.columns:
                        st.error("취소할 사용일 열을 확인할 수 없습니다.")
                        st.stop()

                    parsed_date, cancel_amount = parse_use_entry(selected_value)
                    if parsed_date is None:
                        st.error("선택한 사용일의 날짜 형식이 올바르지 않아 취소를 중단했습니다.")
                        st.stop()

                    df = df.astype(object)
                    df.loc[idx, selected_col] = ""

                    # 기존 합계를 단순 차감하지 않고 남은 사용일 전체로 재계산합니다.
                    target_one = recalculate_vacation_summary(df.loc[[idx]].copy())
                    df.loc[idx, "사용 연차"] = target_one.loc[idx, "사용 연차"]
                    df.loc[idx, "잔여 연차"] = target_one.loc[idx, "잔여 연차"]

                    save_vacation_data(df)

                    save_vacation_log(
                        action="취소",
                        target_name=str(selected_name),
                        use_date=str(parsed_date.date()),
                        used_days=format_leave_number(cancel_amount),
                        reason="연차 사용 취소",
                        note=f"{selected_col} 취소"
                    )

                    load_google_sheet_data.clear()
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
                        load_google_sheet_data.clear()
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
                                load_google_sheet_data.clear()
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
                        load_google_sheet_data.clear()
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

        with metric_col1:
            ui_card("해당 월 사용 건수", monthly_count, "사용 횟수")

        with metric_col2:
            ui_card("해당 월 총 사용일수", format_leave_number(monthly_amount), "사용 연차")

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

        st.markdown(
            '<div class="erp-section-title">📊 기본 연차 정보</div>',
            unsafe_allow_html=True
        )

        basic_df = display_df[basic_cols].copy()

        if "잔여 연차" in basic_df.columns:
            styled_basic_df = basic_df.style.map(
                style_remaining_leave,
                subset=["잔여 연차"]
            )
            st.dataframe(styled_basic_df, use_container_width=True, height=400)
        else:
            st.dataframe(basic_df, use_container_width=True, height=400)

        st.markdown(
            '<div class="erp-section-title">🧾 연차 사용 이력</div>',
            unsafe_allow_html=True
        )

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
