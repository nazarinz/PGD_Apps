"""
Theme configuration untuk konsistensi visual di seluruh PGD Apps.
"""

# Color palette
PRIMARY_COLOR = "#1f77b4"
PRIMARY_DARK = "#0d47a1"
PRIMARY_LIGHT = "#42a5f5"
SUCCESS_COLOR = "#28a745"
WARNING_COLOR = "#ffc107"
ERROR_COLOR = "#dc3545"
INFO_COLOR = "#17a2b8"

# Neutral colors
LIGHT_BG = "#f8f9fa"
BORDER_COLOR = "#e0e0e0"
TEXT_DARK = "#212529"
TEXT_GRAY = "#666"
TEXT_LIGHT = "#999"

# Typography
FONT_FAMILY = "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"

# Spacing (in rem)
SPACING_XS = "0.25rem"
SPACING_SM = "0.5rem"
SPACING_MD = "1rem"
SPACING_LG = "1.5rem"
SPACING_XL = "2rem"
SPACING_XXL = "3rem"

# Border radius
RADIUS_SMALL = "4px"
RADIUS_MEDIUM = "6px"
RADIUS_LARGE = "8px"

# Shadows
SHADOW_SMALL = "0 2px 6px rgba(0, 0, 0, 0.08)"
SHADOW_MEDIUM = "0 4px 12px rgba(0, 0, 0, 0.12)"
SHADOW_LARGE = "0 8px 20px rgba(31, 119, 180, 0.15)"
SHADOW_HOVER = "0 8px 20px rgba(31, 119, 180, 0.15)"

# Transitions
TRANSITION_FAST = "all 0.2s ease"
TRANSITION_NORMAL = "all 0.3s ease"
TRANSITION_SLOW = "all 0.5s ease"

# Icon sizes (em)
ICON_SMALL = "1rem"
ICON_MEDIUM = "1.5rem"
ICON_LARGE = "2rem"
ICON_XL = "2.5rem"

def get_color_scheme(theme_type: str = "light") -> dict:
    """
    Dapatkan color scheme berdasarkan tema.
    
    Args:
        theme_type: 'light' atau 'dark'
    
    Returns:
        Dictionary dengan warna untuk berbagai komponen
    """
    if theme_type == "dark":
        return {
            "bg_primary": "#1e1e1e",
            "bg_secondary": "#2d2d2d",
            "text_primary": "#ffffff",
            "text_secondary": "#b0b0b0",
            "border": "#404040",
        }
    else:  # light
        return {
            "bg_primary": "#ffffff",
            "bg_secondary": "#f8f9fa",
            "text_primary": "#212529",
            "text_secondary": "#666",
            "border": "#e0e0e0",
        }


def create_styled_button(label: str, color: str = "primary") -> str:
    """
    Buat button HTML yang styled.
    """
    color_map = {
        "primary": PRIMARY_COLOR,
        "success": SUCCESS_COLOR,
        "warning": WARNING_COLOR,
        "error": ERROR_COLOR,
        "info": INFO_COLOR,
    }
    
    bg_color = color_map.get(color, PRIMARY_COLOR)
    
    return f"""
    <button style="
        background-color: {bg_color};
        color: white;
        border: none;
        padding: 0.6rem 1.2rem;
        border-radius: {RADIUS_MEDIUM};
        font-weight: 500;
        cursor: pointer;
        transition: {TRANSITION_NORMAL};
    " onmouseover="this.style.transform='translateY(-2px)'; this.style.boxShadow='{SHADOW_HOVER}';"
       onmouseout="this.style.transform='none'; this.style.boxShadow='none';">
        {label}
    </button>
    """
