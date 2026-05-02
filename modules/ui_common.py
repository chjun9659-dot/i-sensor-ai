import streamlit as st

def render_common_style():
    st.markdown("""
    <style>
    .main {
        background-color: #f1f5f9 !important;
    }
    .yw-card {
        background: #ffffff;
        border-radius: 14px;
        padding: 16px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 8px 20px rgba(0,0,0,0.15);
        border-left: 5px solid transparent;
        min-height: 80px;
        margin-bottom: 8px;
        transition: all 0.2s ease;
    }
    .yw-card:hover {
        transform: translateY(-8px) scale(1.03);
        box-shadow: 0 20px 40px rgba(0,0,0,0.25);
    }

    div[data-testid="stMetric"]:hover {
        transform: translateY(-4px) scale(1.01);
        box-shadow: 0 10px 20px rgba(0,0,0,0.12) !important;
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
                                
    </style>
    """, unsafe_allow_html=True)


def ui_card(title, value, sub="", status=""):
    st.markdown(f"""
    <div class="yw-card {status}">
        <div class="card-title">{title}</div>
        <div class="card-value">{value}</div>
        <div class="card-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)
