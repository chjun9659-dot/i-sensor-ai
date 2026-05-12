import streamlit as st
import pandas as pd
from datetime import date

def inspection_page():
    from __main__ import (
        render_common_style,
        load_inspection_data,
        normalize_inspection_df,
        apply_product_filter,
        detect_inspection_duplicates,
        ui_card,
        INSPECTION_STATUS_OPTIONS,
        PRODUCT_OPTIONS,
        CONTRACT_OPTIONS,
        safe_int,
        INSPECTION_COLUMNS,
        save_inspection_data,
        set_inspection_flash,
        find_original_inspection_index,
        ENV_OPTIONS,
        JATU_OPTIONS,
        upload_file_to_drive,
    )

    render_common_style()

    st.markdown(
        '<div class="erp-page-title">🔎 실사 관리 프로그램</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="erp-page-desc">실사 요청 등록 → 담당자 배정 → 일정 입력 → 결과 작성 → 계약 여부 관리</div>',
        unsafe_allow_html=True
    )

    try:
        df = load_inspection_data()
        df = normalize_inspection_df(df)
        df = apply_product_filter(df)
        dup_df = detect_inspection_duplicates(df)
        if not dup_df.empty:
            st.warning(f"⚠️ 중복 데이터 {len(dup_df)}건 발견되었습니다. (구글시트 정리 필요)")

        df = df.reset_index(drop=True)
        df["row_id"] = df.index

        total_count = len(df)
        pending_count = len(df[df["진행상태"] == "요청접수"])
        assigned_count = len(df[df["진행상태"].isin(["담당자배정", "일정확정", "실사진행"])])
        done_count = len(df[df["진행상태"] == "실사완료"])
        contract_done_count = len(df[df["계약여부"] == "계약"])

        c1, c2, c3, c4, c5 = st.columns(5)

        with c1:
            ui_card("전체 요청", total_count)

        with c2:
            ui_card("요청접수", pending_count)

        with c3:
            ui_card("진행중", assigned_count)

        with c4:
            ui_card("실사완료", done_count)

        with c5:
            ui_card("계약완료", contract_done_count)

        st.divider()

    except Exception as e:
        st.error(f"실사 데이터를 불러오지 못했습니다: {e}")
        return

    with st.expander("📝 1. 실사 요청 등록", expanded=False):
        st.markdown('<div class="erp-section-title">실사 요청 등록</div>', unsafe_allow_html=True)
        form_ver = st.session_state.inspection_form_version

        with st.form(f"inspection_request_form_new_{form_ver}"):
            c1, c2, c3 = st.columns(3)
            req_date = c1.date_input("요청일", value=date.today(), key=f"req_date_{form_ver}")
            operator_name = c2.text_input("운영사", key=f"operator_name_{form_ver}")
            site_name = c3.text_input("현장명", key=f"site_name_{form_ver}")

            c4, c5, c6 = st.columns(3)
            site_address = c4.text_input("현장주소", key=f"site_address_{form_ver}")
            site_phone = c5.text_input("현장연락처", key=f"site_phone_{form_ver}")
            product_type = c6.selectbox("상품구분", PRODUCT_OPTIONS, key=f"product_type_{form_ver}")
            env_gov = st.selectbox("환경부", ["", "대상", "비대상"], key=f"env_{form_ver}")
            jatu = st.selectbox("자투", ["", "있음", "없음"], key=f"jatu_{form_ver}")

            c7, c8, c9 = st.columns(3)
            parking_count = c7.number_input("주차면수", min_value=0, step=1, value=0, key=f"parking_count_{form_ver}")
            new_qty = c8.number_input("신규설치수량", min_value=0, step=1, value=0, key=f"new_qty_{form_ver}")
            installed_qty = c9.number_input("기설치수량", min_value=0, step=1, value=0, key=f"installed_qty_{form_ver}")

            c10, c11 = st.columns(2)
            sales_manager = c10.text_input("영업담당자", key=f"sales_manager_{form_ver}")
            sales_phone = c11.text_input("영업담당자 연락처", key=f"sales_phone_{form_ver}")

            request_content = st.text_area("요청내용", key=f"request_content_{form_ver}")
            note = st.text_input("비고", key=f"note_{form_ver}")

            st.markdown('<div class="erp-section-title">첨부파일</div>', unsafe_allow_html=True)
            uploaded_file = st.file_uploader(
                "실사 관련 파일 업로드",
                type=["pdf", "png", "jpg", "jpeg", "xlsx", "xls", "doc", "docx"],
                key=f"insp_uploaded_file_new_{form_ver}"
            )

            submit_request = st.form_submit_button("실사 요청 등록")

            if submit_request:
                if not site_name.strip():
                    st.warning("현장명을 입력해주세요.")
                elif not sales_manager.strip():
                    st.warning("영업담당자를 입력해주세요.")
                else:
                    attachment_name = ""
                    attachment_link = ""

                    if uploaded_file is not None:
                        attachment_name = uploaded_file.name
                        attachment_link = ""
                    else:
                        attachment_name = ""
                        attachment_link = ""

                    new_row = pd.DataFrame([{
                        "요청일": str(req_date),
                        "운영사": operator_name.strip(),
                        "현장명": site_name.strip(),
                        "현장주소": site_address.strip(),
                        "현장연락처": site_phone.strip(),
                        "주차면수": int(parking_count),
                        "상품구분": product_type,
                        "환경부": env_gov,
                        "자투": jatu,
                        "신규설치수량": int(new_qty),
                        "기설치수량": int(installed_qty),
                        "영업담당자": sales_manager.strip(),
                        "영업담당연락처": sales_phone.strip(),
                        "요청내용": request_content.strip(),
                        "비고": note.strip(),
                        "첨부파일명": attachment_name,
                        "첨부파일링크": attachment_link,
                        "실사담당자": "",
                        "실사예정일": "",
                        "실사완료일": "",
                        "진행상태": "요청접수",
                        "실사결과": "",
                        "특이사항": "",
                        "후속조치": "",
                        "계약여부": "대기",
                        "계약일": "",
                        "계약수량": 0,
                        "계약금액": 0,
                        "미계약사유": ""
                    }])

                    # ✅ 반드시 원본 전체 실사 데이터를 다시 불러와서 저장
                    full_df = load_inspection_data()
                    full_df = full_df[INSPECTION_COLUMNS].copy()

                    # ✅ 새 행 1번만 추가
                    save_df = pd.concat([full_df, new_row], ignore_index=True)

                    # ✅ 전체 데이터 저장
                    save_inspection_data(save_df)   

                    set_inspection_flash("실사 요청이 등록되었습니다.", "success")
                    st.session_state.inspection_form_version += 1
                    st.rerun()

    st.divider()

    with st.expander("📋 2. 전체 실사 현황", expanded=False):
        st.markdown('<div class="erp-section-title">전체 실사 현황</div>', unsafe_allow_html=True)
        st.markdown('<div class="erp-soft-box">전체 실사 데이터를 확인합니다.</div>', unsafe_allow_html=True)

        status_list = ["전체"] + INSPECTION_STATUS_OPTIONS
        product_list = ["전체"] + PRODUCT_OPTIONS
        contract_list = ["전체"] + CONTRACT_OPTIONS

        f1, f2, f3, f4 = st.columns(4)
        status_filter = f1.selectbox("진행상태", status_list, key="insp_filter_status_new")
        product_filter = f2.selectbox("상품구분", product_list, key="insp_filter_product_new")
        contract_filter = f3.selectbox("계약여부", contract_list, key="insp_filter_contract_new")
        keyword = f4.text_input("검색", placeholder="현장명 / 주소 / 담당자 / 운영사", key="insp_filter_keyword_new")

        filtered_df = df.copy()

        if status_filter != "전체":
            filtered_df = filtered_df[filtered_df["진행상태"] == status_filter]

        if product_filter != "전체":
            filtered_df = filtered_df[filtered_df["상품구분"] == product_filter]

        if contract_filter != "전체":
            filtered_df = filtered_df[filtered_df["계약여부"] == contract_filter]

        if keyword.strip():
            kw = keyword.strip()
            filtered_df = filtered_df[
                filtered_df["현장명"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["현장주소"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["영업담당자"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["실사담당자"].astype(str).str.contains(kw, case=False, na=False) |
                filtered_df["운영사"].astype(str).str.contains(kw, case=False, na=False)
            ].copy()

        show_df = filtered_df[[
            "요청일",
            "상품구분",
            "환경부",
            "자투",
            "현장명",
            "현장주소",
            "현장연락처",
            "운영사",
            "신규설치수량",
            "기설치수량",
            "주차면수",

            "영업담당자",
            "영업담당연락처",

            "실사담당자",
            "실사예정일",
            "진행상태",
            "계약여부",

            "첨부파일링크"
        ]].copy()

        def status_style(val):
            if str(val) == "요청접수":
                return "background-color: #fef3c7; color: #92400e; font-weight: 600;"
            elif str(val) in ["담당자배정", "일정확정", "실사진행"]:
                return "background-color: #dbeafe; color: #1d4ed8; font-weight: 600;"
            elif str(val) == "실사완료":
                return "background-color: #dcfce7; color: #166534; font-weight: 600;"
            elif str(val) == "계약완료":
                return "background-color: #dcfce7; color: #166534; font-weight: 700;"
            elif str(val) == "미계약종결":
                return "background-color: #fee2e2; color: #991b1b; font-weight: 600;"
            return ""

        def contract_style(val):
            if str(val) == "계약":
                return "background-color: #dcfce7; color: #166534; font-weight: 700;"
            elif str(val) == "미계약":
                return "background-color: #fee2e2; color: #991b1b; font-weight: 700;"
            elif str(val) == "대기":
                return "background-color: #f3f4f6; color: #374151; font-weight: 600;"
            return ""

        if show_df.empty:
            st.info("조건에 맞는 실사 내역이 없습니다.")
        else:
            show_df["첨부파일열기"] = show_df["첨부파일링크"]

            for col in show_df.columns:
                show_df[col] = show_df[col].apply(
                    lambda x: "" if pd.isna(x) or str(x).strip().lower() in ["none", "nan", "nat"] else x
                )
            styled_df = show_df.style.map(status_style, subset=["진행상태"]).map(contract_style, subset=["계약여부"])

            st.dataframe(
                styled_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "현장명": st.column_config.TextColumn("현장명", width="medium"),
                    "현장주소": st.column_config.TextColumn("현장주소", width="medium"),
                    "운영사": st.column_config.TextColumn("운영사", width="small"),
                    "신규설치수량": st.column_config.NumberColumn("신규설치수량", width="small"),
                    "기설치수량": st.column_config.NumberColumn("기설치수량", width="small"),
                    "주차면수": st.column_config.NumberColumn("주차면수", width="small"),
                    "영업담당자": st.column_config.TextColumn("영업담당자", width="small"),
                    "영업담당연락처": st.column_config.TextColumn("영업담당연락처", width="medium"),
                    "실사담당자": st.column_config.TextColumn("실사담당자", width="small"),
                    "실사예정일": st.column_config.TextColumn("실사예정일", width="small"),
                    "진행상태": st.column_config.TextColumn("진행상태", width="small"),
                    "계약여부": st.column_config.TextColumn("계약여부", width="small"),
                    "첨부파일링크": None,
                    "첨부파일열기": st.column_config.LinkColumn(
                        "첨부파일 열기",
                        display_text="열기",
                        width="small"
                    )
                }
            )

            st.caption("상태와 계약여부는 색상으로 구분되며, 첨부파일은 '열기'로 바로 확인할 수 있습니다.")

    st.divider()

    with st.expander("🧑‍🔧 3. 담당자 배정 / 일정 입력", expanded=False):
        if df.empty:
            st.info("배정할 실사 요청이 없습니다.")
        else:
            assign_options = [
                f"{row['row_id']} | {row['요청일']} | {row['현장명']} | {row['상품구분']} | {row['영업담당자']}"
                for _, row in df.iterrows()
            ]

            selected_assign = st.selectbox(
                "배정할 실사 선택",
                assign_options,
                key="insp_assign_select_new"
            )
            assign_idx = int(selected_assign.split("|")[0].strip())
            assign_row = df.loc[df["row_id"] == assign_idx].iloc[0]

            assign_date_raw = str(assign_row["실사예정일"]).strip()
            parsed_assign_date = pd.to_datetime(assign_date_raw, errors="coerce")
            default_assign_date = parsed_assign_date.date() if pd.notna(parsed_assign_date) else date.today()

            with st.form(f"inspection_assign_form_{assign_idx}"):
                a1, a2, a3 = st.columns(3)

                inspector = a1.text_input(
                    "실사담당자",
                    value=str(assign_row["실사담당자"])
                )

                inspect_date = a2.date_input(
                    "실사예정일",
                    value=default_assign_date
                )

                current_status = str(assign_row["진행상태"]).strip()
                default_status_index = (
                    INSPECTION_STATUS_OPTIONS.index(current_status)
                    if current_status in INSPECTION_STATUS_OPTIONS
                    else 0
                )

                inspect_status = a3.selectbox(
                    "진행상태",
                    INSPECTION_STATUS_OPTIONS,
                    index=default_status_index
                )

                b1, b2 = st.columns(2)
                assign_submit = b1.form_submit_button("배정 / 일정 저장")
                delete_submit = b2.form_submit_button("담당자 배정 삭제")

                if assign_submit:
                    target_row = df.loc[df["row_id"] == assign_idx].iloc[0]

                    full_df = load_inspection_data()
                    full_df = full_df[INSPECTION_COLUMNS].copy()

                    original_idx = find_original_inspection_index(full_df, target_row)

                    if original_idx is None:
                        st.error("원본 데이터를 찾지 못했습니다.")
                        st.stop()

                    full_df.loc[original_idx, "실사담당자"] = inspector.strip()
                    full_df.loc[original_idx, "실사예정일"] = str(inspect_date)
                    full_df.loc[original_idx, "진행상태"] = inspect_status

                    save_inspection_data(full_df)
                    load_inspection_data.clear()

                    set_inspection_flash("담당자 배정 및 일정 저장 완료!", "success")
                    st.rerun()

                if delete_submit:
                    target_row = df.loc[df["row_id"] == assign_idx].iloc[0]

                    full_df = load_inspection_data()
                    full_df = full_df[INSPECTION_COLUMNS].copy()

                    original_idx = find_original_inspection_index(full_df, target_row)

                    if original_idx is None:
                        st.error("원본 데이터를 찾지 못했습니다.")
                        st.stop()

                    full_df.loc[original_idx, "실사담당자"] = ""
                    full_df.loc[original_idx, "실사예정일"] = ""
                    full_df.loc[original_idx, "진행상태"] = "요청접수"

                    save_inspection_data(full_df)
                    load_inspection_data.clear()
                    set_inspection_flash("담당자 배정이 삭제되었습니다.", "success")
                    st.rerun()

    st.divider()

    with st.expander("📝 4. 실사 결과 입력", expanded=False):
        if df.empty:
            st.info("입력할 실사 내역이 없습니다.")
        else:
            result_options = [
                f"{row['row_id']} | {row['현장명']} | {row['실사담당자']} | {row['진행상태']}"
                for _, row in df.iterrows()
            ]

            selected_result = st.selectbox("결과 입력 대상 선택", result_options, key="insp_result_select_new")
            result_idx = int(selected_result.split("|")[0].strip())
            result_row = df.loc[df["row_id"] == result_idx].iloc[0]

            complete_date_raw = str(result_row["실사완료일"]).strip()
            parsed_complete_date = pd.to_datetime(complete_date_raw, errors="coerce")
            default_complete_date = parsed_complete_date.date() if pd.notna(parsed_complete_date) else date.today()

            with st.form(f"inspection_result_form_{result_idx}"):
                r1, r2 = st.columns(2)
                result_text = r1.text_area("실사결과", value=str(result_row["실사결과"]))
                special_note = r2.text_area("특이사항", value=str(result_row["특이사항"]))

                r3, r4 = st.columns(2)
                follow_up = r3.text_area("후속조치", value=str(result_row["후속조치"]))
                complete_date = r4.date_input("실사완료일", value=default_complete_date)

                result_status = st.selectbox(
                    "진행상태",
                    INSPECTION_STATUS_OPTIONS,
                    index=INSPECTION_STATUS_OPTIONS.index(result_row["진행상태"]) if result_row["진행상태"] in INSPECTION_STATUS_OPTIONS else 0
                )

                result_submit = st.form_submit_button("실사 결과 저장")

                if result_submit:
                    target_row = df.loc[df["row_id"] == result_idx].iloc[0]

                    full_df = load_inspection_data()
                    full_df = full_df[INSPECTION_COLUMNS].copy()

                    original_idx = find_original_inspection_index(full_df, target_row)

                    if original_idx is None:
                        st.error("원본 데이터를 찾지 못했습니다.")
                        st.stop()

                    full_df.loc[original_idx, "실사결과"] = result_text.strip()
                    full_df.loc[original_idx, "특이사항"] = special_note.strip()
                    full_df.loc[original_idx, "후속조치"] = follow_up.strip()
                    full_df.loc[original_idx, "실사완료일"] = str(complete_date)
                    full_df.loc[original_idx, "진행상태"] = result_status

                    save_inspection_data(full_df)

                    set_inspection_flash("실사 결과 저장 완료!", "success")
                    st.rerun()

    st.divider()

    with st.expander("💰 5. 계약 여부 입력", expanded=False):
        if df.empty:
            st.info("계약 처리할 내역이 없습니다.")
        else:
            contract_options = [
                f"{row['row_id']} | {row['현장명']} | {row['상품구분']} | 현재:{row['계약여부']}"
                for _, row in df.iterrows()
            ]

            selected_contract = st.selectbox("계약 처리 대상 선택", contract_options, key="insp_contract_select_new")
            contract_idx = int(selected_contract.split("|")[0].strip())
            contract_row = df.loc[df["row_id"] == contract_idx].iloc[0]

            contract_date_raw = str(contract_row["계약일"]).strip()
            parsed_contract_date = pd.to_datetime(contract_date_raw, errors="coerce")
            default_contract_date = parsed_contract_date.date() if pd.notna(parsed_contract_date) else date.today()

            contract_qty_default = safe_int(contract_row["계약수량"], 0)
            contract_amount_default = safe_int(contract_row["계약금액"], 0)

            with st.form(f"inspection_contract_form_{contract_idx}"):
                ct1, ct2, ct3 = st.columns(3)
                contract_status = ct1.selectbox(
                    "계약여부",
                    CONTRACT_OPTIONS,
                    index=CONTRACT_OPTIONS.index(contract_row["계약여부"]) if contract_row["계약여부"] in CONTRACT_OPTIONS else 0
                )
                contract_date = ct2.date_input("계약일", value=default_contract_date)
                contract_qty = ct3.number_input("계약수량", min_value=0, step=1, value=contract_qty_default)

                ct4, ct5 = st.columns(2)
                contract_amount = ct4.number_input("계약금액", min_value=0, step=10000, value=contract_amount_default)
                fail_reason = ct5.text_input("미계약사유", value=str(contract_row["미계약사유"]))

                contract_submit = st.form_submit_button("계약 정보 저장")

                if contract_submit:
                    target_row = df.loc[df["row_id"] == contract_idx].iloc[0]

                    full_df = load_inspection_data()
                    full_df = full_df[INSPECTION_COLUMNS].copy()

                    original_idx = find_original_inspection_index(full_df, target_row)

                    if original_idx is None:
                        st.error("원본 데이터를 찾지 못했습니다.")
                        st.stop()

                    full_df.loc[original_idx, "계약여부"] = contract_status
                    full_df.loc[original_idx, "계약일"] = str(contract_date) if contract_status == "계약" else ""
                    full_df.loc[original_idx, "계약수량"] = int(contract_qty) if contract_status == "계약" else 0
                    full_df.loc[original_idx, "계약금액"] = int(contract_amount) if contract_status == "계약" else 0
                    full_df.loc[original_idx, "미계약사유"] = fail_reason.strip() if contract_status == "미계약" else ""

                    if contract_status == "계약":
                        full_df.loc[original_idx, "진행상태"] = "계약완료"
                    elif contract_status == "미계약":
                        full_df.loc[original_idx, "진행상태"] = "미계약종결"

                    save_inspection_data(full_df)

    st.divider()

    with st.expander("✏️ 6. 상세 보기 / 수정", expanded=False):
        if df.empty:
            st.info("조회할 내역이 없습니다.")
        else:
            view_options = [
                f"{row['row_id']} | {row['현장명']} | {row['상품구분']} | {row['진행상태']}"
                for _, row in df.iterrows()
            ]

            selected_view = st.selectbox("조회 대상 선택", view_options, key="insp_view_select_new")
            view_idx = int(selected_view.split("|")[0].strip())
            view_row = df.loc[df["row_id"] == view_idx].iloc[0]

            is_edit_mode = (
                st.session_state.get("inspection_edit_mode", False)
                and st.session_state.get("inspection_edit_target") == view_idx
            )

            if not is_edit_mode:
                st.markdown("""
                <style>
                .detail-title {
                    font-size: 22px;
                    font-weight: 800;
                    color: #0f172a;
                    margin-bottom: 14px;
                }
                .detail-summary-wrap {
                    display: grid;
                    grid-template-columns: repeat(4, 1fr);
                    gap: 12px;
                    margin-bottom: 18px;
                }
                .detail-summary-card {
                    border: 1px solid #e5e7eb;
                    border-radius: 14px;
                    background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
                    padding: 14px 16px;
                    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.06);
                }
                .detail-summary-label {
                    font-size: 12px;
                    color: #64748b;
                    margin-bottom: 8px;
                }
                .detail-summary-value {
                    font-size: 20px;
                    font-weight: 700;
                    color: #0f172a;
                    line-height: 1.25;
                }
                .detail-section {
                    margin-top: 10px;
                    margin-bottom: 14px;
                    padding: 14px;
                    border: 1px solid #e5e7eb;
                    border-radius: 14px;
                    background: #ffffff;
                    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.05);
                }
                .detail-section-title {
                    font-size: 16px;
                    font-weight: 800;
                    color: #0f172a;
                    margin-bottom: 12px;
                }
                .detail-label {
                    font-size: 13px;
                    font-weight: 700;
                    color: #334155;
                    margin-bottom: 6px;
                }
                .detail-box {
                    border: 1px solid #dbe3ee;
                    border-radius: 12px;
                    background: #f8fbff;
                    padding: 12px 14px;
                    min-height: 46px;
                    font-size: 14px;
                    color: #0f172a;
                    display: flex;
                    align-items: center;
                }
                .detail-textarea {
                    border: 1px solid #e5e7eb;
                    border-radius: 12px;
                    background: #f8fafc;
                    padding: 14px;
                    white-space: pre-wrap;
                    line-height: 1.7;
                    font-size: 14px;
                    color: #111827;
                    min-height: 72px;
                }
                </style>
                """, unsafe_allow_html=True)

                st.markdown('<div class="detail-title">📄 실사 상세보기</div>', unsafe_allow_html=True)

                st.markdown(
                    f"""
                    <div class="detail-summary-wrap">
                        <div class="detail-summary-card">
                            <div class="detail-summary-label">진행상태</div>
                            <div class="detail-summary-value">{str(view_row["진행상태"]).strip() or "-"}</div>
                        </div>
                        <div class="detail-summary-card">
                            <div class="detail-summary-label">계약여부</div>
                            <div class="detail-summary-value">{str(view_row["계약여부"]).strip() or "-"}</div>
                        </div>
                        <div class="detail-summary-card">
                            <div class="detail-summary-label">상품구분</div>
                            <div class="detail-summary-value">{str(view_row["상품구분"]).strip() or "-"}</div>
                        </div>
                        <div class="detail-summary-card">
                            <div class="detail-summary-label">영업담당자</div>
                            <div class="detail-summary-value">{str(view_row["영업담당자"]).strip() or "-"}</div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                # 1. 기본 정보
                st.markdown('<div class="detail-section">', unsafe_allow_html=True)
                st.markdown('<div class="detail-section-title">기본 정보</div>', unsafe_allow_html=True)

                row1_col1, row1_col2, row1_col3 = st.columns(3)
                with row1_col1:
                    st.markdown('<div class="detail-label">요청일</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["요청일"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row1_col2:
                    st.markdown('<div class="detail-label">운영사</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["운영사"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row1_col3:
                    st.markdown('<div class="detail-label">현장명</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["현장명"]).strip() or "-"}</div>', unsafe_allow_html=True)

                row2_col1, row2_col2, row2_col3 = st.columns(3)
                with row2_col1:
                    st.markdown('<div class="detail-label">현장주소</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["현장주소"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row2_col2:
                    st.markdown('<div class="detail-label">현장연락처</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["현장연락처"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row2_col3:
                    st.markdown('<div class="detail-label">상품구분</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["상품구분"]).strip() or "-"}</div>', unsafe_allow_html=True)

                row3_col1, row3_col2, row3_col3 = st.columns(3)
                with row3_col1:
                    st.markdown('<div class="detail-label">환경부</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["환경부"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row3_col2:
                    st.markdown('<div class="detail-label">자투</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["자투"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row3_col3:
                    st.markdown('<div class="detail-label">주차면수</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["주차면수"]).strip() or "0"}</div>', unsafe_allow_html=True)

                row4_col1, row4_col2, row4_col3 = st.columns(3)
                with row4_col1:
                    st.markdown('<div class="detail-label">신규설치수량</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["신규설치수량"]).strip() or "0"}</div>', unsafe_allow_html=True)
                with row4_col2:
                    st.markdown('<div class="detail-label">기설치수량</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["기설치수량"]).strip() or "0"}</div>', unsafe_allow_html=True)
                with row4_col3:
                    st.markdown('<div class="detail-label">실사담당자</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["실사담당자"]).strip() or "-"}</div>', unsafe_allow_html=True)

                row5_col1, row5_col2, row5_col3 = st.columns(3)
                with row5_col1:
                    st.markdown('<div class="detail-label">영업담당자</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["영업담당자"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row5_col2:
                    st.markdown('<div class="detail-label">영업담당자 연락처</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["영업담당연락처"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row5_col3:
                    st.markdown('<div class="detail-label">실사예정일</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["실사예정일"]).strip() or "-"}</div>', unsafe_allow_html=True)

                row6_col1, row6_col2, row6_col3 = st.columns(3)
                with row6_col1:
                    st.markdown('<div class="detail-label">실사완료일</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["실사완료일"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row6_col2:
                    st.markdown('<div class="detail-label">계약일</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["계약일"]).strip() or "-"}</div>', unsafe_allow_html=True)
                with row6_col3:
                    st.markdown('<div class="detail-label">계약수량</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["계약수량"]).strip() or "0"}</div>', unsafe_allow_html=True)

                row7_col1, row7_col2 = st.columns(2)
                with row7_col1:
                    st.markdown('<div class="detail-label">계약금액</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["계약금액"]).strip() or "0"}</div>', unsafe_allow_html=True)
                with row7_col2:
                    st.markdown('<div class="detail-label">미계약사유</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="detail-box">{str(view_row["미계약사유"]).strip() or "-"}</div>', unsafe_allow_html=True)

                st.markdown('</div>', unsafe_allow_html=True)

                # 2. 내용 정보
                st.markdown('<div class="detail-section">', unsafe_allow_html=True)
                st.markdown('<div class="detail-section-title">내용 정보</div>', unsafe_allow_html=True)

                st.markdown('<div class="detail-label">요청내용</div>', unsafe_allow_html=True)
                request_text = str(view_row["요청내용"]).strip()
                st.markdown(
                    f'<div class="detail-textarea">{request_text if request_text else "요청내용이 없습니다."}</div>',
                    unsafe_allow_html=True
                )

                st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)

                st.markdown('<div class="detail-label">비고</div>', unsafe_allow_html=True)
                note_text = str(view_row["비고"]).strip()
                st.markdown(
                    f'<div class="detail-textarea">{note_text if note_text else "비고가 없습니다."}</div>',
                    unsafe_allow_html=True
                )

                st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)

                st.markdown('<div class="detail-label">실사결과</div>', unsafe_allow_html=True)
                result_text = str(view_row["실사결과"]).strip()
                st.markdown(
                    f'<div class="detail-textarea">{result_text if result_text else "실사결과가 없습니다."}</div>',
                    unsafe_allow_html=True
                )

                st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)

                st.markdown('<div class="detail-label">특이사항</div>', unsafe_allow_html=True)
                special_text = str(view_row["특이사항"]).strip()
                st.markdown(
                    f'<div class="detail-textarea">{special_text if special_text else "특이사항이 없습니다."}</div>',
                    unsafe_allow_html=True
                )

                st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)

                st.markdown('<div class="detail-label">후속조치</div>', unsafe_allow_html=True)
                follow_text = str(view_row["후속조치"]).strip()
                st.markdown(
                    f'<div class="detail-textarea">{follow_text if follow_text else "후속조치가 없습니다."}</div>',
                    unsafe_allow_html=True
                )

                st.markdown('</div>', unsafe_allow_html=True)

                # 3. 첨부파일
                st.markdown('<div class="detail-section">', unsafe_allow_html=True)
                st.markdown('<div class="detail-section-title">첨부파일</div>', unsafe_allow_html=True)

                if str(view_row["첨부파일링크"]).strip():
                    file_name = str(view_row["첨부파일명"]).strip() or "첨부파일 열기"
                    st.link_button(file_name, str(view_row["첨부파일링크"]))
                else:
                    st.markdown('<div class="detail-box">첨부파일이 없습니다.</div>', unsafe_allow_html=True)

                st.markdown('</div>', unsafe_allow_html=True)

                st.markdown("---")

                if st.button("수정", use_container_width=True, key=f"insp_edit_mode_btn_{view_idx}"):
                    st.session_state.inspection_edit_mode = True
                    st.session_state.inspection_edit_target = view_idx
                    st.rerun()   

            else:

                st.markdown('<div class="erp-section-title">📌 기본 정보 수정</div>', unsafe_allow_html=True)

                req_date_raw = str(view_row["요청일"]).strip()
                parsed_req_date = pd.to_datetime(req_date_raw, errors="coerce")
                default_req_date = parsed_req_date.date() if pd.notna(parsed_req_date) else date.today()

                with st.form(f"inspection_edit_form_{view_idx}"):

                    e1, e2, e3 = st.columns(3)
                    edit_req_date = e1.date_input("요청일 수정", value=default_req_date)
                    edit_operator = e2.text_input("운영사 수정", value=str(view_row["운영사"]))
                    edit_name = e3.text_input("현장명 수정", value=str(view_row["현장명"]))

                    e4, e5, e6 = st.columns(3)
                    edit_addr = e4.text_input("현장주소 수정", value=str(view_row["현장주소"]))
                    edit_phone = e5.text_input("현장연락처 수정", value=str(view_row["현장연락처"]))
                    edit_product = e6.selectbox(
                        "상품구분 수정",
                        PRODUCT_OPTIONS,
                        index=PRODUCT_OPTIONS.index(view_row["상품구분"]) if view_row["상품구분"] in PRODUCT_OPTIONS else 0
                    )

                    e6_1, e6_2 = st.columns(2)
                    current_env = str(view_row["환경부"]).strip() if "환경부" in view_row.index else ""
                    current_jatu = str(view_row["자투"]).strip() if "자투" in view_row.index else ""

                    env_index = ENV_OPTIONS.index(current_env) if current_env in ENV_OPTIONS else 0
                    jatu_index = JATU_OPTIONS.index(current_jatu) if current_jatu in JATU_OPTIONS else 0

                    env_gov = e6_1.selectbox("환경부 수정", ENV_OPTIONS, index=env_index, key=f"edit_env_{view_idx}")
                    jatu = e6_2.selectbox("자투 수정", JATU_OPTIONS, index=jatu_index, key=f"edit_jatu_{view_idx}")

                    e7, e8, e9 = st.columns(3)
                    edit_parking = e7.number_input("주차면수 수정", min_value=0, step=1, value=safe_int(view_row["주차면수"], 0))
                    edit_new_qty = e8.number_input("신규설치수량 수정", min_value=0, step=1, value=safe_int(view_row["신규설치수량"], 0))
                    edit_old_qty = e9.number_input("기설치수량 수정", min_value=0, step=1, value=safe_int(view_row["기설치수량"], 0))

                    e10, e11 = st.columns(2)
                    edit_sales = e10.text_input("영업담당자 수정", value=str(view_row["영업담당자"]))
                    edit_sales_phone = e11.text_input("영업담당자 연락처 수정", value=str(view_row["영업담당연락처"]))

                    edit_request = st.text_area("요청내용 수정", value=str(view_row["요청내용"]))
                    edit_note = st.text_input("비고 수정", value=str(view_row["비고"]))

                    st.markdown('<div class="erp-section-title">첨부파일 수정</div>', unsafe_allow_html=True)
                    edit_uploaded_file = st.file_uploader(
                        "새 첨부파일 업로드",
                        type=["pdf", "png", "jpg", "jpeg", "xlsx", "xls", "doc", "docx"],
                        key=f"insp_edit_uploaded_file_{view_idx}"
                    )

                    current_file_name = str(view_row["첨부파일명"]).strip()
                    current_file_link = str(view_row["첨부파일링크"]).strip()

                    delete_current_file = False

                    if current_file_link:
                        st.caption(f"현재 첨부파일: {current_file_name if current_file_name else '첨부파일'}")
                        delete_current_file = st.checkbox(
                            "기존 첨부파일 삭제",
                            key=f"delete_current_file_{view_idx}"
                        )

                    s1, s2 = st.columns(2)
                    save_submit = s1.form_submit_button("기본 정보 수정 저장", use_container_width=True)
                    cancel_submit = s2.form_submit_button("취소", use_container_width=True)

                    if save_submit:
                        target_row = df.loc[df["row_id"] == view_idx].iloc[0]

                        full_df = load_inspection_data()
                        full_df = full_df[INSPECTION_COLUMNS].copy()

                        original_idx = find_original_inspection_index(full_df, target_row)

                        if original_idx is None:
                            st.error("원본 데이터를 찾지 못했습니다.")
                            st.stop()

                        full_df.loc[original_idx, "요청일"] = str(edit_req_date)
                        full_df.loc[original_idx, "운영사"] = edit_operator.strip()
                        full_df.loc[original_idx, "현장명"] = edit_name.strip()
                        full_df.loc[original_idx, "현장주소"] = edit_addr.strip()
                        full_df.loc[original_idx, "현장연락처"] = edit_phone.strip()
                        full_df.loc[original_idx, "상품구분"] = edit_product
                        full_df.loc[original_idx, "주차면수"] = int(edit_parking)
                        full_df.loc[original_idx, "신규설치수량"] = int(edit_new_qty)
                        full_df.loc[original_idx, "기설치수량"] = int(edit_old_qty)
                        full_df.loc[original_idx, "영업담당자"] = edit_sales.strip()
                        full_df.loc[original_idx, "영업담당연락처"] = edit_sales_phone.strip()
                        full_df.loc[original_idx, "요청내용"] = edit_request.strip()
                        full_df.loc[original_idx, "비고"] = edit_note.strip()
                        full_df.loc[original_idx, "환경부"] = env_gov
                        full_df.loc[original_idx, "자투"] = jatu

                        save_inspection_data(full_df)

                        if delete_current_file:
                            full_df.loc[original_idx, "첨부파일명"] = ""
                            full_df.loc[original_idx, "첨부파일링크"] = ""

                        if edit_uploaded_file is not None:
                            full_df.loc[original_idx, "첨부파일명"] = edit_uploaded_file.name
                            full_df.loc[original_idx, "첨부파일링크"] = ""

                        save_inspection_data(full_df)

                        st.session_state.inspection_edit_mode = False
                        st.session_state.inspection_edit_target = None
                        set_inspection_flash("기본 정보 수정 완료!", "success")
                        st.rerun()

                    if cancel_submit:
                        st.session_state.inspection_edit_mode = False
                        st.session_state.inspection_edit_target = None
                        st.rerun()

    st.divider()

    with st.expander("🗑️ 7. 실사 요청 삭제", expanded=False):
        if df.empty:
            st.info("삭제할 내역이 없습니다.")
        else:
            delete_options = [
                f"{row['row_id']} | {row['현장명']} | {row['상품구분']} | {row['영업담당자']}"
                for _, row in df.iterrows()
            ]

            selected_delete = st.selectbox("삭제할 내역 선택", delete_options, key="insp_delete_select_new")
            confirm_delete = st.checkbox("정말 삭제합니다. 되돌리기 어렵습니다.", key="insp_delete_confirm_new")

            if st.button("선택 내역 삭제", key="insp_delete_btn_new"):
                if not confirm_delete:
                    st.warning("삭제 확인 체크를 먼저 해주세요.")
                else:
                    delete_idx = int(selected_delete.split("|")[0].strip())
                    target_row = df.loc[df["row_id"] == delete_idx].iloc[0]

                    full_df = load_inspection_data()
                    full_df = full_df[INSPECTION_COLUMNS].copy()

                    original_idx = find_original_inspection_index(full_df, target_row)

                    if original_idx is None:
                        st.error("원본 데이터를 찾지 못했습니다.")
                        st.stop()

                    full_df = full_df.drop(index=original_idx).reset_index(drop=True)

                    save_inspection_data(full_df)

                    set_inspection_flash("삭제가 완료되었습니다.", "success")
                    st.rerun()
