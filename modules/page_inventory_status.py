import streamlit as st
import pandas as pd

def inventory_status_page():

    from __main__ import is_admin

    if not is_admin():
        st.warning("관리자만 접근 가능합니다.")
        return
    
    from __main__ import (
        render_common_style,
        ui_card,
        load_df,
    )

    render_common_style()

    from __main__ import get_gsheet_client, get_current_sheet_urls
    import re

    client = get_gsheet_client()

    url = get_current_sheet_urls().get("입출고현황", "")

    sheet_id = re.search(
        r"/d/([a-zA-Z0-9-_]+)",
        url
    ).group(1)

    spreadsheet = client.open_by_key(sheet_id)

    ws = spreadsheet.worksheet("판매현황")

    values = ws.get_all_values()

    # 아이센서 탭 추가 읽기
    ws_stock = spreadsheet.worksheet("아이센서")

    stock_values = ws_stock.get_all_values()

    stock_headers = [str(x).strip() for x in stock_values[1]]
    stock_rows = stock_values[2:]

    stock_df = pd.DataFrame(
        stock_rows,
        columns=stock_headers
    ).fillna("")

    stock_df = stock_df.loc[
        :,
        stock_df.columns.astype(str).str.strip() != ""
    ]

    headers = [str(x).strip() for x in values[2]]
    rows = values[3:]

    df = pd.DataFrame(rows, columns=headers).fillna("")

    df = df.loc[:, df.columns.astype(str).str.strip() != ""]

    st.markdown(
        '<div class="erp-page-title">📦 아이센서 입출고 / 입금 현황</div>',
        unsafe_allow_html=True
    )

    st.info("읽기 전용 페이지입니다")

    for col in ["수량","아이센서","시공","발행 예정금액","입금 확인 금액","미수 금액"]:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.strip()
            .replace("", "0")
            .astype(float)
        )

    # 판매현황 숫자 정리
    for col in ["수량","발행 예정금액","입금 확인 금액","미수 금액"]:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .replace("", "0")
            .astype(float)
        )

    # 아이센서 숫자 정리
    for col in ["수량","합계 금액","송금액","미수금","전체 재고"]:
        stock_df[col] = (
            stock_df[col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .replace("", "0")
            .astype(float)
        )

    # 판매현황
    total_out = df["수량"].sum()
    total_sales = df["발행 예정금액"].sum()
    sales_paid = df["입금 확인 금액"].sum()
    sales_misu = df["미수 금액"].sum()

    # 아이센서
    total_in = stock_df["수량"].sum()
    buy_total = stock_df["합계 금액"].sum()
    buy_misu = stock_df["미수금"].sum()

    # 재고
    current_stock = total_in - total_out

    c1,c2,c3,c4,c5,c6=st.columns(6)

    with c1:
        ui_card("총 입고", f"{int(total_in):,}")

    with c2:
        ui_card("총 출고", f"{int(total_out):,}")

    with c3:
        ui_card("현재 재고", f"{int(current_stock):,}")

    with c4:
        ui_card("매입 미수", f"{int(buy_misu):,}원")

    with c5:
        ui_card("매출 미수", f"{int(sales_misu):,}원")

    with c6:
        ui_card("입금액", f"{int(sales_paid):,}원")

    st.divider()

    st.markdown(
        '<div class="erp-section-title">판매 / 출고 / 입금 전체 현황</div>',
        unsafe_allow_html=True
    )

    display_df = df.copy()

    for col in ["아이센서", "시공", "발행 예정금액", "입금 확인 금액", "미수 금액"]:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(
                lambda x: f"{int(float(x)):,}" if str(x).strip() not in ["", "nan"] else ""
            )

    st.dataframe(display_df, use_container_width=True, hide_index=True)
