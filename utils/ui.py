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

    _render_theme_selector()

    # Apply custom styling
    _apply_custom_styles()


def _render_theme_selector():
    """Theme selector untuk mode Light/Dark/System di sidebar."""
    if "theme_mode" not in st.session_state:
        st.session_state.theme_mode = "System"

    with st.sidebar:
        st.markdown("### 🎨 Theme")
        selected = st.radio(
            "Pilih mode tampilan",
            options=["System", "Light", "Dark"],
            index=["System", "Light", "Dark"].index(st.session_state.theme_mode),
            horizontal=True,
            label_visibility="collapsed",
            key="theme_mode_selector",
        )

    st.session_state.theme_mode = selected

    selected_value = {"System": "system", "Light": "light", "Dark": "dark"}[selected]
    st.markdown(
        f"""
        <script>
            (() => {{
                const userTheme = '{selected_value}';
                const root = document.documentElement;
                if (userTheme === 'system') {{
                    root.removeAttribute('data-user-theme');
                }} else {{
                    root.setAttribute('data-user-theme', userTheme);
                }}
            }})();
        </script>
        """,
        unsafe_allow_html=True,
    )


def _apply_custom_styles():
    """Terapkan custom CSS untuk UI yang lebih baik, termasuk dark/light mode."""
    st.markdown("""
    <style>
    :root {
        --pgd-bg-primary: #ffffff;
        --pgd-bg-secondary: #f8f9fa;
        --pgd-text-primary: #212529;
        --pgd-text-secondary: #666;
        --pgd-border: #e0e0e0;
        --pgd-heading: #1f77b4;
        --pgd-heading-dark: #0d47a1;
    }

    @media (prefers-color-scheme: dark) {
        :root {
            --pgd-bg-primary: #0e1117;
            --pgd-bg-secondary: #1b2130;
            --pgd-text-primary: #e6edf3;
            --pgd-text-secondary: #9aa4b2;
            --pgd-border: #2f3a4d;
            --pgd-heading: #7cb9ff;
            --pgd-heading-dark: #9ecbff;
        }
    }

    :root[data-user-theme='light'] {
        --pgd-bg-primary: #ffffff;
        --pgd-bg-secondary: #f8f9fa;
        --pgd-text-primary: #212529;
        --pgd-text-secondary: #666;
        --pgd-border: #e0e0e0;
        --pgd-heading: #1f77b4;
        --pgd-heading-dark: #0d47a1;
    }

    :root[data-user-theme='dark'] {
        --pgd-bg-primary: #0e1117;
        --pgd-bg-secondary: #1b2130;
        --pgd-text-primary: #e6edf3;
        --pgd-text-secondary: #9aa4b2;
        --pgd-border: #2f3a4d;
        --pgd-heading: #7cb9ff;
        --pgd-heading-dark: #9ecbff;
    }

    html, body, .stApp,
    section[data-testid="stAppViewContainer"],
    section[data-testid="stAppViewContainer"] > .main {
        background-color: var(--pgd-bg-primary) !important;
    }

    * {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }

    h1 {
        color: var(--pgd-heading);
        border-bottom: 3px solid var(--pgd-heading);
        padding-bottom: 0.5rem;
        margin-bottom: 1rem;
    }

    h2 {
        color: var(--pgd-heading-dark);
        margin-top: 1.5rem;
        margin-bottom: 0.8rem;
    }

    h3 {
        color: var(--pgd-heading);
        margin-top: 1rem;
    }

    p, li, label, .stMarkdown, [data-testid="stMarkdownContainer"] {
        color: var(--pgd-text-primary) !important;
    }

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

    .stAlert {
        border-radius: 6px;
        padding: 1rem;
        border-left: 4px solid;
    }

    .stContainer {
        background-color: var(--pgd-bg-secondary);
        border-radius: 8px;
        padding: 1rem;
        border: 1px solid var(--pgd-border);
    }

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

    .sidebar .sidebar-content {
        background-color: var(--pgd-bg-secondary);
    }

    .stTextInput input, .stNumberInput input, .stSelectbox select, .stMultiSelect {
        border-radius: 6px;
        border: 1px solid var(--pgd-border);
        background-color: var(--pgd-bg-primary);
        color: var(--pgd-text-primary);
    }

    .stTextInput input:focus, .stNumberInput input:focus {
        border-color: #1f77b4;
        box-shadow: 0 0 0 2px rgba(31, 119, 180, 0.1);
    }

    .stMetric {
        background-color: var(--pgd-bg-secondary);
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid var(--pgd-border);
    }

    .stDataFrame {
        border-radius: 6px;
        overflow: hidden;
    }

    hr {
        border: none;
        height: 2px;
        background: linear-gradient(to right, var(--pgd-heading), transparent);
        margin: 1.5rem 0;
    }

    .stCaption {
        color: var(--pgd-text-secondary);
        font-size: 0.9rem;
    }

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
