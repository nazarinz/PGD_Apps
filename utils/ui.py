import streamlit as st
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo  # Python 3.9+

def set_page(title: str, icon: str = "🤖"):
    # Try to use custom PGD logo if available
    logo_path = Path(__file__).resolve().parent.parent / "assets" / "logo.png"
    page_icon = str(logo_path) if logo_path.exists() else icon
    try:
        st.set_page_config(
            page_title=title, 
            page_icon=page_icon, 
            layout="wide",
            initial_sidebar_state="expanded",
            menu_items={
                "Get Help": "mailto:nazarudin@gsid.co.id",
                "Report a bug": "mailto:nazarudin@gsid.co.id",
                "About": "PGD Apps v1.0 — Tim PGD Automation"
            }
        )
    except Exception:
        # set_page_config can only be called once; ignore if already set
        pass
    
    # Apply custom styling
    _apply_custom_styles()

def _apply_custom_styles():
    """Terapkan custom CSS untuk UI yang lebih baik."""
    st.markdown("""
    <style>
    /* Main styling */
    * {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Header styling */
    h1 {
        color: #1f77b4;
        border-bottom: 3px solid #1f77b4;
        padding-bottom: 0.5rem;
        margin-bottom: 1rem;
    }
    
    h2 {
        color: #0d47a1;
        margin-top: 1.5rem;
        margin-bottom: 0.8rem;
    }
    
    h3 {
        color: #1565c0;
        margin-top: 1rem;
    }
    
    /* Button styling */
    .stButton > button {
        background-color: #1f77b4;
        color: white;
        border-radius: 6px;
        border: none;
        padding: 0.6rem 1.2rem;
        font-weight: 500;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        background-color: #0d47a1;
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    }
    
    /* Info/Success/Warning/Error boxes */
    .stAlert {
        border-radius: 6px;
        padding: 1rem;
        border-left: 4px solid;
    }
    
    /* Card-like styling for containers */
    .stContainer {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 1rem;
        border: 1px solid #e0e0e0;
    }
    
    /* Tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 6px 6px 0 0;
        border-bottom: 3px solid transparent;
    }
    
    .stTabs [aria-selected="true"] {
        border-bottom-color: #1f77b4;
    }
    
    /* Sidebar styling */
    .sidebar .sidebar-content {
        background-color: #f5f5f5;
    }
    
    /* Input fields */
    .stTextInput input, .stNumberInput input, .stSelectbox select, .stMultiSelect {
        border-radius: 6px;
        border: 1px solid #d0d0d0;
    }
    
    .stTextInput input:focus, .stNumberInput input:focus {
        border-color: #1f77b4;
        box-shadow: 0 0 0 2px rgba(31, 119, 180, 0.1);
    }
    
    /* Metrics */
    .stMetric {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #e0e0e0;
    }
    
    /* Dataframe styling */
    .stDataFrame {
        border-radius: 6px;
        overflow: hidden;
    }
    
    /* Divider */
    hr {
        border: none;
        height: 2px;
        background: linear-gradient(to right, #1f77b4, transparent);
        margin: 1.5rem 0;
    }
    
    /* Caption and text adjustments */
    .stCaption {
        color: #666;
        font-size: 0.9rem;
    }
    
    /* Download button */
    .stDownloadButton > button {
        background-color: #28a745;
        color: white;
    }
    
    .stDownloadButton > button:hover {
        background-color: #218838;
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    }
    </style>
    """, unsafe_allow_html=True)

def header(title: str, subtitle: str | None = None):
    st.title(title)
    if subtitle:
        st.caption(f"📌 {subtitle}")

def footer(note: str = "PGD Apps • Made by Nazarudin Zaini :D"):
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col1:
        st.caption(note)
    
    with col2:
        st.caption("📧 Email: nazarudin@gsid.co.id")
    
    with col3:
        # Waktu WIB (Jakarta) lengkap dengan detik
        now_wib = datetime.now(ZoneInfo("Asia/Jakarta"))
        st.caption(f"⏰ {now_wib.strftime('%Y-%m-%d')}")
