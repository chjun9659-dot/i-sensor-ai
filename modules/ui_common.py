import streamlit as st

def page_title(text):
    st.markdown(f"""
    <div class="erp-page-title">
        {text}
    </div>
    """, unsafe_allow_html=True)


def render_global_style():
    st.markdown("""
    <style>
    .main {
        background-color: #f1f5f9 !important;
    }

    .erp-page-title {
        font-size: 22px !important;
        font-weight: 700 !important;
        margin-top: 10px !important;
        margin-bottom: 10px !important;
        display: flex !important;
        align-items: center !important;
        gap: 8px !important;
        color: #0f172a !important;
    }

    .erp-page-desc {
        font-size: 13px !important;
        color: #64748b !important;
        margin-bottom: 20px !important;
    }

    .yw-card {
        background: #ffffff;
        border-radius: 14px;
        padding: 16px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 8px 20px rgba(0,0,0,0.12);
        border-left: 5px solid transparent;
        min-height: 80px;
        margin-bottom: 8px;
        transition: all 0.2s ease;
    }

    .yw-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 22px rgba(0,0,0,0.12);
    }

    div[data-testid="stMetric"] {
        background: #ffffff;
        border-radius: 14px;
        padding: 14px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        transition: all 0.2s ease;
    }

    div[data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 18px rgba(0,0,0,0.10) !important;
    }

    .card-value {
        font-size: 24px;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.2;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Noto Sans KR", sans-serif;
        font-variant-numeric: tabular-nums;
    }

    .card-title {
        font-size: 13px;
        font-weight: 700;
        color: #64748b;
        margin-bottom: 3px;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Noto Sans KR", sans-serif;
    }

    .card-sub {
        font-size: 11px;
        color: #94a3b8;
        margin-top: 3px;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Noto Sans KR", sans-serif;
    }
    .erp-section-title,
    div.erp-section-title,
    .stMarkdown .erp-section-title {
        font-size: 17px !important;
        font-weight: 700 !important;
        color: #0f172a !important;
        margin-top: 22px !important;
        margin-bottom: 10px !important;
        padding-left: 10px !important;
        border-left: 4px solid #2563eb !important;
        line-height: 1.4 !important;
    }
    .erp-section-desc {
        font-size: 12px !important;
        color: #64748b !important;
        margin-top: -4px !important;
        margin-bottom: 12px !important;
    }            
    </style>
    """, unsafe_allow_html=True)


def render_common_style():
    render_global_style()


def render_inspection_common_style():
    render_global_style()


def ui_card(title, value, sub="", status=""):
    st.markdown(f"""
    <div class="yw-card {status}">
        <div class="card-title">{title}</div>
        <div class="card-value">{value}</div>
        <div class="card-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)
