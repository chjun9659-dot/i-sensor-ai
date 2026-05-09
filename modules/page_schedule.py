import streamlit as st
import pandas as pd
from datetime import date

def schedule_page():
    from __main__ import (
        render_common_style,
        load_schedule_data,
        save_schedule_data,
        append_schedule_data,
        find_original_schedule_index,
        save_schedule_log,
        ui_card,
        EXPECTED_COLUMNS,
    )

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

            if "상품구분" in df.columns:
                product_series = df["상품구분"].astype(str).str.strip()

                if st.session_state.business == "아이센서":
                    df = df[product_series.str.contains("아이센서", na=False)].copy()

                elif st.session_state.business == "전기차 충전기":
                    df = df[product_series.str.contains("전기차", na=False)].copy()

            login_role = str(st.session_state.get("role", "")).strip()
            login_name = str(st.session_state.get("display_name", "")).strip()

            if login_role != "관리자" and "시공담당" in df.columns:
                df = df[df["시공담당"].astype(str).str.strip() == login_name].copy()

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
                ["아이센서", "전기차 충전기"],
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

                e4, e5, e6, e7 = st.columns(4)

                edit_product = e4.selectbox(
                    "상품구분 수정",
                    ["아이센서", "전기차 충전기"],
                    index=0 if str(edit_row["상품구분"]).strip() == "아이센서" else 1
                )

                edit_qty = e5.number_input(
                    "수량 수정",
                    min_value=0,
                    step=1,
                    value=int(edit_row["수량"])
                )

                edit_note = e6.text_input(
                    "비고 수정",
                    value=str(edit_row["비고"])
                )

                edit_status = e7.selectbox(
                    "상태 수정",
                    ["진행중", "완료"],
                    index=0 if str(edit_row["상태"]).strip() == "진행중" else 1
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
                    full_df.loc[original_idx, "상품구분"] = edit_product

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
    # 여기부터 기존 UI 코드 조금씩 이동
