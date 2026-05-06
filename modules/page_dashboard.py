import streamlit as st

from modules.ui_common import render_common_style

def page_dashboard():

    st.title("📊 통합 대시보드")
    st.caption("영업 / 계약 / 시공 / 실사 / 유지보수 / 연차 / 오늘 할 일을 한 화면에서 확인합니다.")

    render_common_style()

    st.success("대시보드 모듈 연결 성공")
