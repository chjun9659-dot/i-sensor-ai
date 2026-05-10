import streamlit as st
import pandas as pd
from datetime import date, datetime

def page_dashboard():
    from __main__ import (
        render_common_style,
        page_title,
        load_notice,
        ui_card,

        apply_author_filter,
        load_tasks_df,
        load_schedule_df,
        apply_role_filter,
        load_df,
        load_schedule_data,
        apply_product_filter,
        load_inspection_data,
        normalize_inspection_df,
        load_maintenance_data,
        load_maintenance_payment_data,
        get_contract_expiring_soon,
        get_gsheet_client,
        NOTICE_SHEET_URL,
    )

    render_common_style()

    page_title("📊 통합 대시보드")
    st.caption("영업 / 계약 / 시공 / 실사 / 유지보수 / 연차 / 오늘 할 일을 한 화면에서 확인합니다.")
    notice_df = load_notice() 

    if not notice_df.empty:
        notice_df = notice_df.fillna("")
        notice_df = notice_df[notice_df["내용"].astype(str).str.strip() != ""].copy()

        if not notice_df.empty:
            recent_notice_df = notice_df.tail(3).iloc[::-1].reset_index(drop=True)

            for _, row in recent_notice_df.iterrows():
                st.markdown(f"""
                <div style="
                    background:#fff7ed;
                    border:1px solid #fed7aa;
                    border-left:6px solid #f97316;
                    border-radius:12px;
                    padding:14px 18px;
                    margin:10px 0 12px 0;
                    font-size:15px;
                    font-weight:600;
                    color:#7c2d12;
                ">
                📢 공지사항: {row['내용']}
                <div style="font-size:12px; font-weight:400; color:#9a3412; margin-top:6px;">
                    {row.get('작성일', '')} / {row.get('작성자', '')}
                </div>
                </div>
                """, unsafe_allow_html=True)

    # 👇 관리자만 공지 등록 가능
    if st.session_state.get("role") == "관리자":
        with st.expander("📢 공지 등록", expanded=False):
            notice_text = st.text_input("공지 내용 입력", key="notice_input")

            if st.button("공지 등록", key="notice_save_btn"):
                if notice_text.strip() == "":
                    st.warning("공지 내용을 입력하세요")
                else:
                    try:
                        client = get_gsheet_client()
                        spreadsheet = client.open_by_url(NOTICE_SHEET_URL)
                        ws = spreadsheet.worksheet("공지사항")

                        ws.append_row([
                            datetime.now().strftime("%Y-%m-%d"),
                            notice_text,
                            st.session_state.get("display_name", st.session_state.get("username"))
                        ])

                        st.success("공지 등록 완료!")
                        st.rerun()

                    except Exception as e:
                        st.error(f"공지 등록 실패: {e}")

            st.divider()
            st.markdown('<div class="erp-section-title">🗑️ 공지 삭제</div>', unsafe_allow_html=True)

            notice_delete_df = load_notice().fillna("")

            if notice_delete_df.empty:
                st.info("삭제할 공지가 없습니다.")
            else:
                notice_delete_df = notice_delete_df.reset_index(drop=True)

                delete_options = [
                    f"{idx} | {row.get('작성일', '')} | {row.get('내용', '')}"
                    for idx, row in notice_delete_df.iterrows()
                    if str(row.get("내용", "")).strip() != ""
                ]

                if delete_options:
                    selected_delete_notice = st.selectbox(
                        "삭제할 공지 선택",
                        delete_options,
                        key="notice_delete_select"
                    )

                    confirm_notice_delete = st.checkbox(
                        "정말 이 공지를 삭제합니다.",
                        key="notice_delete_confirm"
                    )

                    if st.button("선택 공지 삭제", key="notice_delete_btn"):
                        if not confirm_notice_delete:
                            st.warning("삭제 확인 체크를 먼저 해주세요.")
                        else:
                            delete_idx = int(selected_delete_notice.split("|")[0].strip())

                            try:
                                client = get_gsheet_client()
                                spreadsheet = client.open_by_url(NOTICE_SHEET_URL)
                                ws = spreadsheet.worksheet("공지사항")

                                save_df = notice_delete_df.drop(index=delete_idx).reset_index(drop=True)

                                for col in ["작성일", "내용", "작성자"]:
                                    if col not in save_df.columns:
                                        save_df[col] = ""

                                save_df = save_df[["작성일", "내용", "작성자"]].fillna("")

                                ws.clear()
                                ws.update(
                                    [["작성일", "내용", "작성자"]] + save_df.values.tolist()
                                )

                                st.success("공지 삭제 완료!")
                                st.rerun()

                            except Exception as e:
                                st.error(f"공지 삭제 실패: {e}")
                else:
                    st.info("삭제할 공지가 없습니다.")              

    today = date.today()
    today_str = str(today)

    # =========================
    # 1. 기본 데이터 불러오기
    # =========================
    try:
        sales_df = pd.DataFrame()
        possible_df = pd.DataFrame()
        bid_df = pd.DataFrame()
        contract_df = pd.DataFrame()
        task_df = apply_author_filter(load_tasks_df())
        schedule_common_df = apply_author_filter(load_schedule_df())

    except Exception as e:
        st.error(f"대시보드 기본 데이터를 불러오지 못했습니다: {e}")
        return

    # KPI용 최소 데이터만 로딩
    try:
        sales_df = apply_role_filter(load_df("영업현황")) if st.session_state.business == "아이센서" else pd.DataFrame()
        contract_df = apply_role_filter(load_df("계약단지")) if st.session_state.business == "아이센서" else apply_role_filter(load_df("계약접수현황"))
    except:
        sales_df = pd.DataFrame()
        contract_df = pd.DataFrame()

    # =========================
    # 2. 상단 핵심 KPI
    # =========================
    st.markdown(
        '<div class="erp-section-title">📌 핵심 현황</div>',
        unsafe_allow_html=True
    )

    k1, k2, k3 = st.columns(3)

    with k1:
        ui_card("영업현황", len(sales_df) if not sales_df.empty else 0, "전체 영업 데이터", "info")

    with k2:
        ui_card("계약/접수", len(contract_df) if not contract_df.empty else 0, "계약·접수 현황", "success")

    with k3:
        ui_card("오늘 할 일", len(task_df) if not task_df.empty else 0, "등록된 할 일", "danger")

    st.divider()

    # =========================
    # 3. 시공 일정 요약
    # =========================
    st.markdown(
        '<div class="erp-section-title">🛠 시공 일정 요약</div>',
        unsafe_allow_html=True
    )

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
                st.markdown('<div class="erp-section-title">📅 오늘 시공 일정</div>', unsafe_allow_html=True)
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
    st.markdown(
        '<div class="erp-section-title">🔎 실사 관리 요약</div>',
        unsafe_allow_html=True
    )

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
            st.markdown('<div class="erp-section-title">최근 실사 요청</div>', unsafe_allow_html=True)
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

        st.markdown(
            '<div class="erp-section-title">📡 유지보수 / 수금 요약</div>',
            unsafe_allow_html=True
        )

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
            with m1:
                ui_card("전체 유지보수 계약", len(maintenance_df) if not maintenance_df.empty else 0, "전체 계약")

            with m2:
                ui_card("진행중 계약", active_count, "현재 진행중")

            with m3:
                ui_card("전체 미수금", f"{total_unpaid:,} 원", "미입금 합계")

            with m4:
                ui_card("60일 내 종료예정", expiring_count, "계약 종료 임박")

            if total_unpaid > 0:
                st.warning(f"현재 유지보수 미수금이 {total_unpaid:,}원 있습니다.")

            if not unpaid_df.empty:
                st.markdown(
                    '<div class="erp-section-title">💰 미입금 현황</div>',
                    unsafe_allow_html=True
                )
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

        st.markdown(
            '<div class="erp-section-title">👥 연차 요약</div>',
            unsafe_allow_html=True
        )

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

                with v1:
                    ui_card("직원 수", 0, "연차 대상 직원")

                with v2:
                    ui_card("잔여 5일 이하", 0, "연차 소진 임박")

                with v3:
                    ui_card("잔여 0일 이하", 0, "잔여 연차 없음")
                st.info("연차 데이터가 없습니다.")
                
            else:
                for col in ["발생 연차", "사용 연차", "잔여 연차"]:
                    if col in vacation_df.columns:
                        vacation_df[col] = pd.to_numeric(vacation_df[col], errors="coerce").fillna(0)

                low_leave_df = vacation_df[vacation_df["잔여 연차"] <= 5].copy() if "잔여 연차" in vacation_df.columns else pd.DataFrame()
                zero_leave_df = vacation_df[vacation_df["잔여 연차"] <= 0].copy() if "잔여 연차" in vacation_df.columns else pd.DataFrame()

                v1, v2, v3 = st.columns(3)
                with v1:
                    ui_card("직원 수", len(vacation_df), "연차 대상 직원")

                with v2:
                    ui_card("잔여 5일 이하", len(low_leave_df), "연차 소진 임박")

                with v3:
                    ui_card("잔여 0일 이하", len(zero_leave_df), "잔여 연차 없음")

                if login_role == "관리자":
                    st.markdown('<div class="erp-section-title">📊 연차 사용 현황 그래프</div>', unsafe_allow_html=True)

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
    st.markdown(
        '<div class="erp-section-title">✅ 오늘 할 일 / 일정</div>',
        unsafe_allow_html=True
    )

    t1, t2 = st.columns(2)

    with t1:
        st.markdown('<div class="erp-section-title">오늘 할 일</div>', unsafe_allow_html=True)

        if task_df.empty:
            st.info("등록된 할 일이 없습니다.")
        else:
            st.dataframe(
                task_df[["등록일시", "작성자", "사업", "할일"]].tail(10),
                use_container_width=True,
                hide_index=True
            )

    with t2:
        st.markdown('<div class="erp-section-title">일정 관리</div>', unsafe_allow_html=True)

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
