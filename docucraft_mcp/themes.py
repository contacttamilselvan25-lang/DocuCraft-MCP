"""Universal Theme & Styling Engine for DocuCraft MCP.

Provides curated color palettes, typography mappings, and conversion helpers
for PPTX, DOCX, XLSX, and PDF document generation.
"""

from typing import Dict, Any, Tuple, Optional


PRESET_THEMES: Dict[str, Dict[str, Any]] = {
    "corporate_blue": {
        "name": "Corporate Blue",
        "primary": "#1E3A8A",      # Deep Navy
        "secondary": "#0284C7",    # Bright Ocean Blue
        "accent": "#06B6D4",       # Cyan highlight
        "bg_light": "#F8FAFC",     # Crisp Light Slate
        "text_dark": "#0F172A",    # Dark Charcoal
        "text_light": "#FFFFFF",   # Pure White
        "font_heading": "Calibri",
        "font_body": "Calibri",
    },
    "emerald_growth": {
        "name": "Emerald Growth",
        "primary": "#065F46",      # Deep Forest Green
        "secondary": "#10B981",    # Emerald Mint
        "accent": "#34D399",       # Light Green Accent
        "bg_light": "#F0FDF4",     # Soft Mint Tint
        "text_dark": "#064E3B",    # Dark Forest Slate
        "text_light": "#FFFFFF",
        "font_heading": "Segoe UI",
        "font_body": "Segoe UI",
    },
    "modern_dark": {
        "name": "Modern Dark",
        "primary": "#0F172A",      # Midnight Slate
        "secondary": "#8B5CF6",    # Electric Violet
        "accent": "#EC4899",       # Neon Pink Accent
        "bg_light": "#1E293B",     # Dark Surface Slate
        "text_dark": "#F8FAFC",    # Off-white text on dark
        "text_light": "#FFFFFF",
        "font_heading": "Arial",
        "font_body": "Arial",
    },
    "warm_executive": {
        "name": "Warm Executive",
        "primary": "#831843",      # Deep Burgundy / Maroon
        "secondary": "#D97706",    # Amber Gold
        "accent": "#F59E0B",       # Goldenrod
        "bg_light": "#FFFBEB",     # Warm Cream
        "text_dark": "#4C0519",    # Dark Wine
        "text_light": "#FFFFFF",
        "font_heading": "Georgia",
        "font_body": "Garamond",
    },
    "minimal_slate": {
        "name": "Minimal Slate",
        "primary": "#334155",      # Slate Gray
        "secondary": "#64748B",    # Muted Gray
        "accent": "#0EA5E9",       # Sky Blue Accent
        "bg_light": "#F1F5F9",     # Light Cool Gray
        "text_dark": "#0F172A",    # Charcoal Black
        "text_light": "#FFFFFF",
        "font_heading": "Helvetica",
        "font_body": "Arial",
    },
    "crimson_bold": {
        "name": "Crimson Bold",
        "primary": "#991B1B",      # Ruby Crimson
        "secondary": "#DC2626",    # Vivid Red
        "accent": "#F87171",       # Coral Accent
        "bg_light": "#FEF2F2",     # Rose Tint
        "text_dark": "#450A0A",    # Dark Maroon
        "text_light": "#FFFFFF",
        "font_heading": "Trebuchet MS",
        "font_body": "Calibri",
    },
    "sunset_orange": {
        "name": "Sunset Orange",
        "primary": "#C2410C",      # Deep Rust Orange
        "secondary": "#EA580C",    # Bright Tangerine
        "accent": "#F59E0B",       # Sunburst Gold
        "bg_light": "#FFF7ED",     # Peach Cream
        "text_dark": "#431407",    # Dark Umber
        "text_light": "#FFFFFF",
        "font_heading": "Segoe UI",
        "font_body": "Calibri",
    },
    "royal_indigo": {
        "name": "Royal Indigo",
        "primary": "#3730A3",      # Deep Indigo
        "secondary": "#6366F1",    # Bright Lavender Indigo
        "accent": "#A855F7",       # Purple Accent
        "bg_light": "#EEF2FF",     # Soft Indigo Tint
        "text_dark": "#1E1B4B",    # Night Sky
        "text_light": "#FFFFFF",
        "font_heading": "Arial",
        "font_body": "Calibri",
    }
}


def hex_to_rgb(hex_str: str) -> Tuple[int, int, int]:
    """Convert hex color string (e.g. '#1E3A8A' or '1E3A8A') to (R, G, B) tuple."""
    clean = hex_str.strip().lstrip("#")
    if len(clean) == 3:
        clean = "".join([c * 2 for c in clean])
    if len(clean) != 6:
        clean = "1E3A8A"  # Fallback default navy
    try:
        r = int(clean[0:2], 16)
        g = int(clean[2:4], 16)
        b = int(clean[4:6], 16)
        return (r, g, b)
    except Exception:
        return (30, 58, 138)


def get_theme_config(
    theme_name: str = "corporate_blue",
    primary_color: Optional[str] = None,
    secondary_color: Optional[str] = None,
    accent_color: Optional[str] = None,
    font_family: Optional[str] = None
) -> Dict[str, Any]:
    """Resolve theme configuration blending presets with user overrides.

    Args:
        theme_name: One of the curated presets or 'custom'.
        primary_color: Hex override for primary brand color (e.g. '#10B981').
        secondary_color: Hex override for secondary/accent color.
        accent_color: Hex override for badge/callout highlights.
        font_family: Font override (e.g. 'Calibri', 'Arial', 'Georgia').

    Returns:
        Dict containing full color palettes, RGB tuples, and font choices.
    """
    key = theme_name.lower().strip().replace("-", "_").replace(" ", "_")
    base = PRESET_THEMES.get(key, PRESET_THEMES["corporate_blue"]).copy()

    if primary_color:
        base["primary"] = primary_color.strip()
    if secondary_color:
        base["secondary"] = secondary_color.strip()
    if accent_color:
        base["accent"] = accent_color.strip()
    if font_family:
        base["font_heading"] = font_family.strip()
        base["font_body"] = font_family.strip()

    # Precalculate RGB tuples
    base["primary_rgb"] = hex_to_rgb(base["primary"])
    base["secondary_rgb"] = hex_to_rgb(base["secondary"])
    base["accent_rgb"] = hex_to_rgb(base["accent"])
    base["bg_light_rgb"] = hex_to_rgb(base["bg_light"])
    base["text_dark_rgb"] = hex_to_rgb(base["text_dark"])
    base["text_light_rgb"] = hex_to_rgb(base["text_light"])

    return base
