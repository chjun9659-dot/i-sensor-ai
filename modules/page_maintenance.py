import streamlit as st
import pandas as pd
from datetime import date


def maintenance_page():
    from __main__ import (
        render_common_style,
        page_title,
        ui_card,
        load_maintenance_data,
        load_maintenance_payment_data,
        get_contract_expiring_soon,
        format_currency,
        calculate_total_contract_amount,
        save_maintenance_data,
        save_maintenance_payment_data,
        maintenance_safe_float,
        style_unpaid_amount,
        MAINTENANCE_COLUMNS,
        MAINTENANCE_PAYMENT_COLUMNS,
    )

    render_common_style()

    page_title("📡 아이센서 유지보수관리 프로그램")
    st.caption("유지보수 계약등록, 월별 청구/수금, 미수금 관리")
    st.markdown("""
    <style>
    .maintenance-alert-box {
        border-radius: 14px;
        padding: 14px 18px;
        margin: 8px 0 18px 0;
        font-size: 15px;
        font-weight: 600;
    }
    .maintenance-alert-danger {
        background: #fff1f2;
        border: 1px solid #fecdd3;
        color: #9f1239;
    }
    .maintenance-alert-warning {
        background: #fffbeb;
        border: 1px solid #fde68a;
        color: #92400e;
    }
    .maintenance-guide-box {
        border: 1px solid #e5e7eb;
        background: #ffffff;
        border-radius: 14px;
        padding: 14px 18px;
        margin: 10px 0 18px 0;
        color: #334155;
        font-size: 14px;
        line-height: 1.7;
    }
    .section-gap {
        margin-top: 18px;
        margin-bottom: 18px;
    }
    </style>
    """, unsafe_allow_html=True)

    try:
        df = load_maintenance_data()
        pay_df = load_maintenance_payment_data()
    except Exception as e:
        st.error(f"유지보수 데이터를 불러오지 못했습니다: {e}")
        return

    df = df.reset_index(drop=True)
    df["row_id"] = df.index

    pay_df = pay_df.reset_index(drop=True)
    pay_df["row_id"] = pay_df.index

    current_ym = make_year_month(date.today().year, date.today().month)

    total_count = len(df)
    active_count = len(df[df["계약상태"].astype(str).str.strip() == "진행중"]) if not df.empty else 0
    total_qty = int(df["수량"].sum()) if not df.empty else 0
    total_amount = int(pd.to_numeric(df["총계약금액"], errors="coerce").fillna(0).sum()) if not df.empty else 0

    unpaid_df = pay_df[pay_df["입금여부"].astype(str).str.strip() != "입금완료"].copy() if not pay_df.empty else pd.DataFrame()
    total_unpaid = int(pd.to_numeric(unpaid_df["미수금"], errors="coerce").fillna(0).sum()) if not unpaid_df.empty else 0

    expiring_df = get_contract_expiring_soon(df, within_days=60)
    expiring_count = len(expiring_df)

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    with c1:
        ui_card("전체 계약", total_count, "전체 유지보수 계약")

    with c2:
        ui_card("진행중 계약", active_count, "현재 진행중")

    with c3:
        ui_card("총 수량", total_qty, "설치 수량")

    with c4:
        ui_card("전체 계약금액", format_currency(total_amount), "전체 계약")

    with c5:
        ui_card("전체 미수금", format_currency(total_unpaid), "미수금 합계")

    with c6:
        ui_card("60일 내 종료예정", expiring_count, "계약 종료 임박")
    if total_unpaid > 0:
        st.markdown(
            f'<div class="maintenance-alert-box maintenance-alert-danger">⚠️ 현재 미수금 총액: {format_currency(total_unpaid)} 원</div>',
            unsafe_allow_html=True
        )

    if expiring_count > 0:
        st.markdown(
            f'<div class="maintenance-alert-box maintenance-alert-warning">⏰ 60일 이내 종료 예정 계약이 {expiring_count}건 있습니다.</div>',
            unsafe_allow_html=True
        )

    st.markdown("""
    <div class="maintenance-guide-box">
    • 유지보수 계약 등록 → 월별 청구 생성 → 청구/수금 처리 순서로 사용하시면 됩니다.<br>
    • 미수금이 있는 건은 월별 청구/수금 현황에서 빨간색으로 표시됩니다.<br>
    • 계약 종료 예정 건은 별도 섹션에서 빠르게 확인할 수 있습니다.
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    with st.expander("📝 1. 유지보수 계약 등록", expanded=False):
        with st.form("maintenance_add_form"):
            a1, a2, a3 = st.columns(3)
            code_no = a1.text_input("코드번호", key="mt_code_no")
            site_name = a2.text_input("단지명", key="mt_site_name")
            phone = a3.text_input("연락처", key="mt_phone")

            a4, a5, a6 = st.columns(3)
            region = a4.text_input("지역", key="mt_region")
            sales_manager = a5.text_input("영업담당자", key="mt_sales_manager")
            qty = a6.number_input("수량", min_value=0, step=1, value=0, key="mt_qty")

            a7, a8, a9 = st.columns(3)
            unit_price = a7.number_input("단가", min_value=0, step=1000, value=0, key="mt_unit_price")
            start_date = a8.date_input("계약시작일", value=date.today(), key="mt_start_date")
            end_date = a9.date_input("계약종료일", value=date.today(), key="mt_end_date")

            b1, b2, b3 = st.columns(3)
            contract_status = b1.selectbox("계약상태", ["진행중", "종료", "해지"], key="mt_contract_status")
            billing_cycle = b2.selectbox("청구주기", ["매월", "분기", "반기", "연간"], key="mt_billing_cycle")
            note = b3.text_input("비고", key="mt_note")

            uploaded_file = st.file_uploader("계약서 첨부", type=None, key="mt_uploaded_file")

            total_amount_preview = calculate_total_contract_amount(qty, unit_price)
            st.info(f"자동 계산 총계약금액: {format_currency(total_amount_preview)} 원")

            submitted = st.form_submit_button("등록하기")

            if submitted:
                if not site_name.strip():
                    st.warning("단지명을 입력해주세요.")
                else:
                    attachment_name = ""
                    attachment_link = ""

                    if uploaded_file is not None:
                        st.warning("첨부파일 업로드 기능은 현재 점검중입니다.")

                    new_row = pd.DataFrame([{
                        "코드번호": code_no.strip(),
                        "단지명": site_name.strip(),
                        "연락처": phone.strip(),
                        "지역": region.strip(),
                        "영업담당자": sales_manager.strip(),
                        "수량": int(qty),
                        "단가": float(unit_price),
                        "계약시작일": str(start_date),
                        "계약종료일": str(end_date),
                        "총계약금액": calculate_total_contract_amount(qty, unit_price),
                        "계약상태": contract_status,
                        "청구주기": billing_cycle,
                        "비고": note.strip(),
                        "첨부파일명": attachment_name,
                        "첨부파일링크": attachment_link
                    }])

                    # =========================
                    # 저장용 데이터
                    # =========================
                    full_df = load_maintenance_data()

                    # ✅ 안전장치
                    if full_df is None or full_df.empty:
                        st.error("원본 유지보수 데이터가 비어 있습니다. 저장을 중단합니다.")
                        st.stop()

                    # ✅ 컬럼 보정
                    for col in MAINTENANCE_COLUMNS:
                        if col not in full_df.columns:
                            full_df[col] = ""

                    full_df = full_df[MAINTENANCE_COLUMNS].copy()

                    # ✅ 데이터 추가 (딱 1번)
                    save_df = pd.concat([full_df, new_row], ignore_index=True)

                    # ✅ 저장 (딱 1번)
                    save_maintenance_data(save_df)

                    st.success("유지보수 계약 등록 완료!")
                    st.rerun()

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("🧾 2. 월별 청구 생성", expanded=False):
        g1, g2 = st.columns(2)
        gen_year = g1.number_input("청구 생성 연도", min_value=2024, max_value=2100, value=date.today().year, step=1, key="mt_gen_year")
        gen_month = g2.selectbox("청구 생성 월", list(range(1, 13)), index=date.today().month - 1, key="mt_gen_month")

        target_ym = make_year_month(gen_year, gen_month)
        st.info(f"{target_ym} 기준으로 진행중 계약의 월 청구를 생성합니다. 이미 같은 월 자료가 있으면 중복 생성되지 않습니다.")

        if st.button("해당 월 청구 일괄 생성", key="mt_generate_claim_btn"):
            new_claim_df = generate_monthly_claim_rows(df, pay_df, gen_year, gen_month)

            if new_claim_df.empty:
                st.warning("새로 생성할 청구 자료가 없습니다. 이미 생성되었거나 진행중 계약이 없습니다.")
            else:
                base_df = pay_df[MAINTENANCE_PAYMENT_COLUMNS].copy() if not pay_df.empty else pd.DataFrame(columns=MAINTENANCE_PAYMENT_COLUMNS)
                base_df = pd.concat([base_df, new_claim_df], ignore_index=True)
                save_maintenance_payment_data(base_df)
                st.success(f"{target_ym} 청구자료 {len(new_claim_df)}건 생성 완료!")
                st.rerun()

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("💰 3. 월별 청구/수금 처리", expanded=False):
        if pay_df.empty:
            st.info("생성된 청구자료가 없습니다.")
        else:
            payment_options = [
                f"{row['row_id']} | {row['기준년월']} | {row['코드번호']} | {row['단지명']}"
                for _, row in pay_df.iterrows()
            ]

            selected_payment = st.selectbox("처리할 청구 선택", payment_options, key="mt_payment_select")
            pay_idx = int(selected_payment.split("|")[0].strip())
            pay_row = pay_df.loc[pay_df["row_id"] == pay_idx].iloc[0]

            with st.form(f"mt_payment_form_{pay_idx}"):
                p1, p2, p3 = st.columns(3)
                edit_claim_amount = p1.number_input("청구금액", min_value=0, step=1000, value=int(maintenance_safe_float(pay_row["청구금액"], 0)))
                issue_status = p2.selectbox(
                    "발행여부",
                    ["미발행", "발행완료"],
                    index=0 if str(pay_row["발행여부"]) != "발행완료" else 1
                )
                deposit_status = p3.selectbox(
                    "입금여부",
                    ["미입금", "입금완료"],
                    index=0 if str(pay_row["입금여부"]) != "입금완료" else 1
                )

                p4, p5, p6 = st.columns(3)

                default_issue_date = pd.to_datetime(pay_row["발행일"], errors="coerce")
                if pd.isna(default_issue_date):
                    default_issue_date = pd.Timestamp(date.today())

                default_deposit_date = pd.to_datetime(pay_row["입금일"], errors="coerce")
                if pd.isna(default_deposit_date):
                    default_deposit_date = pd.Timestamp(date.today())

                issue_date = p4.date_input("발행일", value=default_issue_date.date(), key=f"mt_issue_date_{pay_idx}")
                deposit_date = p5.date_input("입금일", value=default_deposit_date.date(), key=f"mt_deposit_date_{pay_idx}")
                payment_note = p6.text_input("비고", value=str(pay_row["비고"]), key=f"mt_payment_note_{pay_idx}")

                unpaid_preview = calculate_unpaid_amount(edit_claim_amount, deposit_status)
                st.info(f"처리 후 미수금: {format_currency(unpaid_preview)} 원")

                payment_submit = st.form_submit_button("저장하기")

                if payment_submit:
                    save_pay_df = pay_df[MAINTENANCE_PAYMENT_COLUMNS].copy()

                    save_pay_df.loc[pay_idx, "청구금액"] = float(edit_claim_amount)
                    save_pay_df.loc[pay_idx, "발행여부"] = issue_status
                    save_pay_df.loc[pay_idx, "발행일"] = str(issue_date) if issue_status == "발행완료" else ""
                    save_pay_df.loc[pay_idx, "입금여부"] = deposit_status
                    save_pay_df.loc[pay_idx, "입금일"] = str(deposit_date) if deposit_status == "입금완료" else ""
                    save_pay_df.loc[pay_idx, "미수금"] = calculate_unpaid_amount(edit_claim_amount, deposit_status)
                    save_pay_df.loc[pay_idx, "비고"] = payment_note.strip()

                    save_maintenance_payment_data(save_pay_df)
                    st.success("청구/수금 처리 완료!")
                    st.rerun()

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("📋 4. 월별 청구/수금 현황", expanded=False):
        f1, f2, f3 = st.columns(3)
        filter_year = f1.number_input("조회 연도", min_value=2024, max_value=2100, value=date.today().year, step=1, key="mt_filter_year")
        filter_month = f2.selectbox("조회 월", list(range(1, 13)), index=date.today().month - 1, key="mt_filter_month")
        filter_deposit = f3.selectbox("입금상태", ["전체", "미입금", "입금완료"], key="mt_filter_deposit")

        target_ym = make_year_month(filter_year, filter_month)
        view_df = pay_df.copy()

        if not view_df.empty:
            view_df = view_df[view_df["기준년월"] == target_ym]

            if filter_deposit != "전체":
                view_df = view_df[view_df["입금여부"] == filter_deposit]

        month_claim_total = int(pd.to_numeric(view_df["청구금액"], errors="coerce").fillna(0).sum()) if not view_df.empty else 0
        month_unpaid_total = int(pd.to_numeric(view_df["미수금"], errors="coerce").fillna(0).sum()) if not view_df.empty else 0
        month_paid_count = len(view_df[view_df["입금여부"] == "입금완료"]) if not view_df.empty else 0

        m1, m2, m3 = st.columns(3)

        with m1:
            ui_card("월 청구금액", format_currency(month_claim_total), "이번달 청구")

        with m2:
            ui_card("월 미수금", format_currency(month_unpaid_total), "미수금 합계")

        with m3:
            ui_card("입금완료 건수", month_paid_count, "입금 완료")

        if view_df.empty:
            st.info("조건에 맞는 수금자료가 없습니다.")
        else:
            show_pay_df = view_df.copy()

            show_pay_df["청구금액"] = pd.to_numeric(show_pay_df["청구금액"], errors="coerce").fillna(0)
            show_pay_df["미수금"] = pd.to_numeric(show_pay_df["미수금"], errors="coerce").fillna(0)

            display_df = show_pay_df[
                [
                    "기준년월", "코드번호", "단지명", "청구금액",
                    "발행여부", "발행일", "입금여부", "입금일",
                    "미수금", "영업담당자", "계약상태", "비고"
                ]
            ].copy()

            # 숫자 포맷
            display_df["청구금액"] = display_df["청구금액"].apply(format_currency)
            display_df["미수금"] = display_df["미수금"].apply(format_currency)

            # 👉 스타일 적용 (핵심)
            styled_df = display_df.style.map(
                style_unpaid_amount,
                subset=["미수금"]
            )

            st.dataframe(styled_df, use_container_width=True, hide_index=True) 

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("📄 5. 계약 현황", expanded=False):
        f1, f2, f3 = st.columns(3)
        region_filter = f1.selectbox(
            "지역 선택",
            ["전체"] + sorted([x for x in df["지역"].astype(str).unique().tolist() if str(x).strip() != ""]),
            key="mt_region_filter"
        )
        status_filter = f2.selectbox(
            "계약상태 선택",
            ["전체", "진행중", "종료", "해지"],
            key="mt_status_filter"
        )
        keyword = f3.text_input("검색", placeholder="코드번호 / 단지명 / 담당자", key="mt_keyword")

        filtered_df = df.copy()

        if region_filter != "전체":
            filtered_df = filtered_df[filtered_df["지역"] == region_filter]

        if status_filter != "전체":
            filtered_df = filtered_df[filtered_df["계약상태"] == status_filter]

        if keyword.strip():
            kw = keyword.strip()
            filtered_df = filtered_df[
                filtered_df["코드번호"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["단지명"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["영업담당자"].astype(str).str.contains(kw, case=False, na=False)
            ]

        if filtered_df.empty:
            st.info("조건에 맞는 계약이 없습니다.")
        else:
            show_df = filtered_df.copy()
            show_df["단가"] = show_df["단가"].apply(format_currency)
            show_df["총계약금액"] = show_df["총계약금액"].apply(format_currency)

            st.dataframe(
                show_df[
                    [
                        "코드번호", "단지명", "연락처", "지역", "영업담당자",
                        "수량", "단가", "계약시작일", "계약종료일",
                        "총계약금액", "계약상태", "청구주기", "비고"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("📅 5-1. 계약 종료 예정", expanded=False):
        soon_df = get_contract_expiring_soon(df, within_days=60)

        if soon_df.empty:
            st.info("60일 이내 종료 예정 계약이 없습니다.")
        else:
            display_soon_df = soon_df[
                [
                    "코드번호", "단지명", "지역", "영업담당자",
                    "계약시작일", "계약종료일", "계약상태", "청구주기", "비고"
                ]
            ].copy()

            st.dataframe(display_soon_df, use_container_width=True, hide_index=True)

    st.divider()        
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("🔍 6. 상세 보기", expanded=False):
        if df.empty:
            st.info("조회할 계약이 없습니다.")
        else:
            view_options = [
                f"{row['row_id']} | {row['코드번호']} | {row['단지명']} | {row['영업담당자']}"
                for _, row in df.iterrows()
            ]
            selected_view = st.selectbox("조회할 계약 선택", view_options, key="mt_view_select")
            view_idx = int(selected_view.split("|")[0].strip())
            row = df.loc[df["row_id"] == view_idx].iloc[0]

            v1, v2 = st.columns(2)

            with v1:
                st.markdown(
                    '<div class="erp-section-title">기본 정보</div>',
                    unsafe_allow_html=True
                )
                st.write(f"**코드번호**: {row['코드번호']}")
                st.write(f"**단지명**: {row['단지명']}")
                st.write(f"**연락처**: {row['연락처']}")
                st.write(f"**지역**: {row['지역']}")
                st.write(f"**영업담당자**: {row['영업담당자']}")
                st.write(f"**수량**: {row['수량']}")
                st.write(f"**단가**: {format_currency(row['단가'])} 원")
                st.write(f"**총계약금액**: {format_currency(row['총계약금액'])} 원")

            with v2:
                st.markdown(
                    '<div class="erp-section-title">계약 정보</div>',
                    unsafe_allow_html=True
                )
                st.write(f"**계약시작일**: {row['계약시작일']}")
                st.write(f"**계약종료일**: {row['계약종료일']}")
                st.write(f"**계약상태**: {row['계약상태']}")
                st.write(f"**청구주기**: {row['청구주기']}")
                st.write(f"**비고**: {row['비고']}")

                if str(row["첨부파일링크"]).strip():
                    st.markdown(f"[📎 첨부파일 열기]({row['첨부파일링크']})")
                else:
                    st.caption("첨부파일 없음")

            st.markdown(
                '<div class="erp-section-title">월별 수금 이력</div>',
                unsafe_allow_html=True
            )
            history_df = pay_df[pay_df["코드번호"].astype(str) == str(row["코드번호"]).strip()].copy()

            if history_df.empty:
                st.info("이 계약의 수금 이력이 없습니다.")
            else:
                history_df["청구금액"] = history_df["청구금액"].apply(format_currency)
                history_df["미수금"] = history_df["미수금"].apply(format_currency)
                st.dataframe(
                    history_df[
                        ["기준년월", "청구금액", "발행여부", "발행일", "입금여부", "입금일", "미수금", "비고"]
                    ],
                    use_container_width=True,
                    hide_index=True
                )

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("✏️ 7. 계약 수정", expanded=False):
        if df.empty:
            st.info("수정할 계약이 없습니다.")
        else:
            edit_options = [
                f"{row['row_id']} | {row['코드번호']} | {row['단지명']} | {row['영업담당자']}"
                for _, row in df.iterrows()
            ]
            selected_edit = st.selectbox("수정할 계약 선택", edit_options, key="mt_edit_select")
            edit_idx = int(selected_edit.split("|")[0].strip())
            edit_row = df.loc[df["row_id"] == edit_idx].iloc[0]

            default_start = pd.to_datetime(edit_row["계약시작일"], errors="coerce")
            if pd.isna(default_start):
                default_start = pd.Timestamp(date.today())

            default_end = pd.to_datetime(edit_row["계약종료일"], errors="coerce")
            if pd.isna(default_end):
                default_end = pd.Timestamp(date.today())

            with st.form(f"maintenance_edit_form_{edit_idx}"):
                e1, e2, e3 = st.columns(3)
                edit_code = e1.text_input("코드번호 수정", value=str(edit_row["코드번호"]))
                edit_site = e2.text_input("단지명 수정", value=str(edit_row["단지명"]))
                edit_phone = e3.text_input("연락처 수정", value=str(edit_row["연락처"]))

                e4, e5, e6 = st.columns(3)
                edit_region = e4.text_input("지역 수정", value=str(edit_row["지역"]))
                edit_sales = e5.text_input("영업담당자 수정", value=str(edit_row["영업담당자"]))
                edit_qty = e6.number_input("수량 수정", min_value=0, step=1, value=int(edit_row["수량"]))

                e7, e8, e9 = st.columns(3)
                edit_unit_price = e7.number_input("단가 수정", min_value=0, step=1000, value=int(maintenance_safe_float(edit_row["단가"], 0)))
                edit_start = e8.date_input("계약시작일 수정", value=default_start.date())
                edit_end = e9.date_input("계약종료일 수정", value=default_end.date())

                e10, e11, e12 = st.columns(3)
                edit_status = e10.selectbox(
                    "계약상태 수정",
                    ["진행중", "종료", "해지"],
                    index=["진행중", "종료", "해지"].index(edit_row["계약상태"]) if str(edit_row["계약상태"]) in ["진행중", "종료", "해지"] else 0
                )
                edit_cycle = e11.selectbox(
                    "청구주기 수정",
                    ["매월", "분기", "반기", "연간"],
                    index=["매월", "분기", "반기", "연간"].index(edit_row["청구주기"]) if str(edit_row["청구주기"]) in ["매월", "분기", "반기", "연간"] else 0
                )
                edit_note = e12.text_input("비고 수정", value=str(edit_row["비고"]))

                new_file = st.file_uploader("새 첨부파일 업로드", type=None, key=f"mt_new_file_{edit_idx}")

                total_amount_preview = calculate_total_contract_amount(edit_qty, edit_unit_price)
                st.info(f"수정 후 총계약금액: {format_currency(total_amount_preview)} 원")

                edit_submit = st.form_submit_button("수정 저장")

                if edit_submit:
                    save_df = df[MAINTENANCE_COLUMNS].copy()

                    old_code = str(save_df.loc[edit_idx, "코드번호"]).strip()
                    new_code = edit_code.strip()

                    save_df.loc[edit_idx, "코드번호"] = new_code
                    save_df.loc[edit_idx, "단지명"] = edit_site.strip()
                    save_df.loc[edit_idx, "연락처"] = edit_phone.strip()
                    save_df.loc[edit_idx, "지역"] = edit_region.strip()
                    save_df.loc[edit_idx, "영업담당자"] = edit_sales.strip()
                    save_df.loc[edit_idx, "수량"] = int(edit_qty)
                    save_df.loc[edit_idx, "단가"] = float(edit_unit_price)
                    save_df.loc[edit_idx, "계약시작일"] = str(edit_start)
                    save_df.loc[edit_idx, "계약종료일"] = str(edit_end)
                    save_df.loc[edit_idx, "총계약금액"] = calculate_total_contract_amount(edit_qty, edit_unit_price)
                    save_df.loc[edit_idx, "계약상태"] = edit_status
                    save_df.loc[edit_idx, "청구주기"] = edit_cycle
                    save_df.loc[edit_idx, "비고"] = edit_note.strip()

                    if new_file is not None:
                        try:
                            new_attachment_name, new_attachment_link = upload_file_to_drive(new_file)
                            save_df.loc[edit_idx, "첨부파일명"] = new_attachment_name
                            save_df.loc[edit_idx, "첨부파일링크"] = new_attachment_link
                        except Exception as e:
                            st.error(f"첨부파일 업로드 실패: {e}")
                            st.stop()

                    save_maintenance_data(save_df)

                    if not pay_df.empty:
                        save_pay_df = pay_df[MAINTENANCE_PAYMENT_COLUMNS].copy()
                        save_pay_df.loc[save_pay_df["코드번호"].astype(str) == old_code, "코드번호"] = new_code
                        save_pay_df.loc[save_pay_df["코드번호"].astype(str) == new_code, "단지명"] = edit_site.strip()
                        save_pay_df.loc[save_pay_df["코드번호"].astype(str) == new_code, "영업담당자"] = edit_sales.strip()
                        save_pay_df.loc[save_pay_df["코드번호"].astype(str) == new_code, "계약상태"] = edit_status
                        save_maintenance_payment_data(save_pay_df)

                    st.success("수정 완료!")
                    st.rerun()

    st.divider()
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    with st.expander("🗑️ 8. 계약 삭제", expanded=False):
        if df.empty:
            st.info("삭제할 계약이 없습니다.")
        else:
            delete_options = [
                f"{row['row_id']} | {row['코드번호']} | {row['단지명']} | {row['영업담당자']}"
                for _, row in df.iterrows()
            ]
            selected_delete = st.selectbox("삭제할 계약 선택", delete_options, key="mt_delete_select")
            confirm_delete = st.checkbox("정말 삭제합니다. 되돌리기 어렵습니다.", key="mt_delete_confirm")

            if st.button("선택 계약 삭제", key="mt_delete_btn"):
                if not confirm_delete:
                    st.warning("삭제 확인 체크를 먼저 해주세요.")
                else:
                    delete_idx = int(selected_delete.split("|")[0].strip())
                    delete_code = str(df.loc[delete_idx, "코드번호"]).strip()

                    save_df = df[MAINTENANCE_COLUMNS].copy()
                    save_df = save_df.drop(index=delete_idx).reset_index(drop=True)
                    save_maintenance_data(save_df)

                    if not pay_df.empty and delete_code != "":
                        save_pay_df = pay_df[MAINTENANCE_PAYMENT_COLUMNS].copy()
                        save_pay_df = save_pay_df[save_pay_df["코드번호"].astype(str) != delete_code].reset_index(drop=True)
                        save_maintenance_payment_data(save_pay_df)

                    st.success("삭제 완료!")
                    st.rerun()   


def make_year_month(year, month):
    return f"{int(year)}-{int(month):02d}"


def calculate_unpaid_amount(claim_amount, deposit_status):
    try:
        claim = float(
            str(claim_amount)
            .replace(",", "")
            .replace("원", "")
            .strip()
        )
    except:
        claim = 0

    if str(deposit_status).strip() == "입금완료":
        return 0

    return claim

def can_generate_claim_by_cycle(start_date_value, billing_cycle, target_year, target_month):
    """
    계약시작일과 청구주기를 기준으로 해당 년/월에 청구 생성 대상인지 판단
    """
    start_dt = pd.to_datetime(start_date_value, errors="coerce")

    if pd.isna(start_dt):
        return False

    start_index = start_dt.year * 12 + start_dt.month
    target_index = int(target_year) * 12 + int(target_month)

    if target_index < start_index:
        return False

    diff = target_index - start_index
    cycle = str(billing_cycle).strip()

    if cycle == "매월":
        return True
    elif cycle == "분기":
        return diff % 3 == 0
    elif cycle == "반기":
        return diff % 6 == 0
    elif cycle == "연간":
        return diff % 12 == 0

    # 값이 비어 있거나 예외면 매월로 처리
    return True


def is_contract_active_for_month(start_date_value, end_date_value, target_year, target_month):
    """
    해당 년/월이 계약기간 안에 있는지 판단
    """
    start_dt = pd.to_datetime(start_date_value, errors="coerce")
    end_dt = pd.to_datetime(end_date_value, errors="coerce")

    if pd.isna(start_dt):
        return False

    month_start = pd.Timestamp(year=int(target_year), month=int(target_month), day=1)
    month_end = month_start + pd.offsets.MonthEnd(0)

    if month_end < start_dt:
        return False

    if not pd.isna(end_dt) and month_start > end_dt:
        return False

    return True

def generate_monthly_claim_rows(contract_df, payment_df, target_year, target_month):
    target_ym = make_year_month(target_year, target_month)

    existing_keys = set()
    if not payment_df.empty:
        for _, row in payment_df.iterrows():
            key = (str(row["코드번호"]).strip(), str(row["기준년월"]).strip())
            existing_keys.add(key)

    new_rows = []

    active_df = contract_df[contract_df["계약상태"].astype(str).isin(["진행중"])].copy()

    for _, row in active_df.iterrows():
        code_no = str(row["코드번호"]).strip()
        site_name = str(row["단지명"]).strip()
        billing_cycle = str(row["청구주기"]).strip()
        start_date_value = row["계약시작일"]
        end_date_value = row["계약종료일"]

        if code_no == "" and site_name == "":
            continue

        # 1) 계약기간 안에 있어야 함
        if not is_contract_active_for_month(start_date_value, end_date_value, target_year, target_month):
            continue

        # 2) 청구주기에 맞는 월만 생성
        if not can_generate_claim_by_cycle(start_date_value, billing_cycle, target_year, target_month):
            continue

        key = (code_no, target_ym)
        if key in existing_keys:
            continue

        claim_amount = calculate_total_contract_amount(row["수량"], row["단가"])

        new_rows.append({
            "코드번호": code_no,
            "단지명": site_name,
            "기준년월": target_ym,
            "청구금액": claim_amount,
            "발행여부": "미발행",
            "발행일": "",
            "입금여부": "미입금",
            "입금일": "",
            "미수금": claim_amount,
            "영업담당자": str(row["영업담당자"]).strip(),
            "계약상태": str(row["계약상태"]).strip(),
            "비고": f"{billing_cycle} 자동생성"
        })

    return pd.DataFrame(new_rows)
