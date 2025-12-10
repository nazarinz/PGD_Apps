"""
Reusable UI components untuk konsistensi visual di seluruh aplikasi.
"""

import streamlit as st
from typing import List, Optional, Dict, Any
from utils.theme import PRIMARY_COLOR, PRIMARY_DARK, SUCCESS_COLOR, WARNING_COLOR, ERROR_COLOR, INFO_COLOR


def render_card(title: str, content: str = "", footer: str = "", icon: str = "📌"):
    """
    Render card container dengan styling yang konsisten.
    
    Args:
        title: Judul card
        content: Konten utama (markdown supported)
        footer: Teks footer (opsional)
        icon: Icon di depan title
    """
    st.markdown(f"""
    <div style="
        background: linear-gradient(135deg, #f8f9fa 0%, #ffffff 100%);
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.08);
        transition: all 0.3s ease;
    " onmouseover="this.style.transform='translateY(-4px)'; this.style.boxShadow='0 8px 20px rgba(31, 119, 180, 0.15)'; this.style.borderColor='{PRIMARY_COLOR}';"
       onmouseout="this.style.transform='none'; this.style.boxShadow='0 2px 6px rgba(0, 0, 0, 0.08)'; this.style.borderColor='#e0e0e0';">
        <h3 style="color: {PRIMARY_DARK}; margin: 0 0 0.5rem 0;">{icon} {title}</h3>
        <div style="color: #555; font-size: 0.95rem; line-height: 1.5; margin-bottom: 0.5rem;">
            {content}
        </div>
        {f'<div style="color: #999; font-size: 0.85rem; margin-top: 1rem; padding-top: 1rem; border-top: 1px solid #e0e0e0;">{footer}</div>' if footer else ''}
    </div>
    """, unsafe_allow_html=True)


def render_stats(stats: Dict[str, str], cols: int = 4):
    """
    Render statistics cards dalam grid layout.
    
    Args:
        stats: Dictionary dengan {label: value}
        cols: Jumlah kolom (2-4)
    """
    columns = st.columns(cols)
    for idx, (label, value) in enumerate(stats.items()):
        with columns[idx % cols]:
            st.metric(label, value)


def render_section_divider():
    """Render section divider yang stylish."""
    st.markdown(f"""
    <hr style="
        border: none;
        height: 2px;
        background: linear-gradient(to right, {PRIMARY_COLOR}, transparent);
        margin: 1.5rem 0;
    ">
    """, unsafe_allow_html=True)


def render_progress_bar(progress: float, label: str = "", color: str = "primary"):
    """
    Render progress bar custom dengan style yang bagus.
    
    Args:
        progress: Nilai 0-1
        label: Label progress (opsional)
        color: 'primary', 'success', 'warning', 'error', 'info'
    """
    color_map = {
        "primary": PRIMARY_COLOR,
        "success": SUCCESS_COLOR,
        "warning": WARNING_COLOR,
        "error": ERROR_COLOR,
        "info": INFO_COLOR,
    }
    
    bar_color = color_map.get(color, PRIMARY_COLOR)
    percentage = int(progress * 100)
    
    st.markdown(f"""
    <div style="margin: 1rem 0;">
        {f'<p style="margin-bottom: 0.5rem; color: #666; font-size: 0.9rem;"><strong>{label}</strong> {percentage}%</p>' if label else f'<p style="margin-bottom: 0.5rem; color: #666; font-size: 0.9rem;">{percentage}%</p>'}
        <div style="
            width: 100%;
            height: 8px;
            background-color: #e0e0e0;
            border-radius: 4px;
            overflow: hidden;
        ">
            <div style="
                width: {percentage}%;
                height: 100%;
                background: linear-gradient(to right, {bar_color}, {PRIMARY_COLOR});
                transition: width 0.5s ease;
            "></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_alert(message: str, alert_type: str = "info", dismissible: bool = False):
    """
    Render custom alert dengan berbagai tipe.
    
    Args:
        message: Pesan alert
        alert_type: 'info', 'success', 'warning', 'error'
        dismissible: Bisa ditutup atau tidak
    """
    icon_map = {
        "info": "ℹ️",
        "success": "✅",
        "warning": "⚠️",
        "error": "❌",
    }
    
    color_map = {
        "info": ("#d1ecf1", "#17a2b8", "#0c5460"),
        "success": ("#d4edda", "#28a745", "#155724"),
        "warning": ("#fff3cd", "#ffc107", "#856404"),
        "error": ("#f8d7da", "#dc3545", "#721c24"),
    }
    
    bg_color, border_color, text_color = color_map.get(alert_type, color_map["info"])
    icon = icon_map.get(alert_type, "ℹ️")
    
    st.markdown(f"""
    <div style="
        background-color: {bg_color};
        border-left: 4px solid {border_color};
        padding: 1rem;
        border-radius: 6px;
        margin: 1rem 0;
        color: {text_color};
    ">
        <strong>{icon} {alert_type.upper()}</strong><br>
        {message}
    </div>
    """, unsafe_allow_html=True)


def render_tabs(tabs: Dict[str, callable]):
    """
    Render tabs dengan callback function untuk setiap tab.
    
    Args:
        tabs: Dictionary dengan {tab_name: content_function}
    
    Contoh:
        def tab1_content():
            st.write("Konten tab 1")
        
        def tab2_content():
            st.write("Konten tab 2")
        
        render_tabs({
            "Tab 1": tab1_content,
            "Tab 2": tab2_content,
        })
    """
    tab_list = st.tabs(list(tabs.keys()))
    for tab, (tab_name, content_func) in zip(tab_list, tabs.items()):
        with tab:
            content_func()


def render_help_box(title: str, content: str):
    """
    Render help/tip box dengan styling khusus.
    
    Args:
        title: Judul help box
        content: Isi help box (markdown supported)
    """
    st.markdown(f"""
    <div style="
        background: linear-gradient(135deg, #e3f2fd 0%, #f3e5f5 100%);
        border-left: 4px solid {PRIMARY_COLOR};
        padding: 1rem;
        border-radius: 6px;
        margin: 1rem 0;
    ">
        <strong style="color: {PRIMARY_DARK};">💡 {title}</strong><br>
        <small style="color: #555;">{content}</small>
    </div>
    """, unsafe_allow_html=True)


def render_code_block(code: str, language: str = "python"):
    """
    Render code block dengan syntax highlighting.
    
    Args:
        code: Kode yang ingin ditampilkan
        language: Bahasa pemrograman (python, sql, javascript, etc)
    """
    st.code(code, language=language)


def render_data_quality_indicator(quality_score: float, label: str = "Data Quality"):
    """
    Render indikator kualitas data dengan visual feedback.
    
    Args:
        quality_score: Nilai 0-100
        label: Label indikator
    """
    if quality_score >= 80:
        color = SUCCESS_COLOR
        status = "Excellent"
    elif quality_score >= 60:
        color = INFO_COLOR
        status = "Good"
    elif quality_score >= 40:
        color = WARNING_COLOR
        status = "Fair"
    else:
        color = ERROR_COLOR
        status = "Poor"
    
    st.markdown(f"""
    <div style="margin: 1rem 0;">
        <div style="display: flex; justify-content: space-between; margin-bottom: 0.5rem;">
            <span style="color: #666; font-weight: 500;">{label}</span>
            <span style="color: {color}; font-weight: bold;">{status} ({quality_score:.0f}%)</span>
        </div>
        <div style="
            width: 100%;
            height: 12px;
            background-color: #e0e0e0;
            border-radius: 6px;
            overflow: hidden;
        ">
            <div style="
                width: {quality_score}%;
                height: 100%;
                background: linear-gradient(to right, {color}, {PRIMARY_COLOR});
                transition: width 0.5s ease;
            "></div>
        </div>
    </div>
    """, unsafe_allow_html=True)
