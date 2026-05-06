import streamlit as st
import pandas as pd
from datetime import date

def page_dashboard():
    from __main__ import (
        render_inspection_common_style,
        render_common_style,
        page_title,
        load_notice,
        ui_card,
    )

    render_inspection_common_style()
    render_common_style()

    page_title("📊 통합 대시보드")
    st.caption("영업 / 계약 / 시공 / 실사 / 유지보수 / 연차 / 오늘 할 일을 한 화면에서 확인합니다.")

    st.info("대시보드 UI 이동 성공"))
