"""
Design system: colors, fonts, and styling constants.
Clean, modern aesthetic inspired by minimal web interfaces.
"""

COLORS = {
    "dark": {
        "bg_primary": "#0f0f0f",
        "bg_secondary": "#1a1a1a",
        "bg_surface": "#242424",
        "bg_hover": "#2e2e2e",
        "bg_input": "#1a1a1a",
        "text_primary": "#f5f5f5",
        "text_secondary": "#8b8b8b",
        "text_muted": "#5a5a5a",
        "accent": "#c97c3a",
        "accent_hover": "#d99a5b",
        "success": "#4ade80",
        "warning": "#fbbf24",
        "error": "#f87171",
        "border": "#2e2e2e",
        "border_subtle": "#1f1f1f",
    },
    "light": {
        "bg_primary": "#ffffff",
        "bg_secondary": "#f7f7f7",
        "bg_surface": "#ffffff",
        "bg_hover": "#f0f0f0",
        "bg_input": "#f5f5f5",
        "text_primary": "#1a1a1a",
        "text_secondary": "#6b6b6b",
        "text_muted": "#999999",
        "accent": "#b5651d",
        "accent_hover": "#c97c3a",
        "success": "#16a34a",
        "warning": "#d97706",
        "error": "#dc2626",
        "border": "#e5e5e5",
        "border_subtle": "#f0f0f0",
    },
}

# Font families with fallbacks
FONT_FAMILY = "Helvetica"
FONT_MONO = "Courier"

FONTS = {
    "heading": (FONT_FAMILY, 18, "bold"),
    "subheading": (FONT_FAMILY, 13, "bold"),
    "body": (FONT_FAMILY, 13),
    "small": (FONT_FAMILY, 11),
    "tiny": (FONT_FAMILY, 10),
    "mono": (FONT_MONO, 12),
}

# Spacing constants
PAD_XS = 4
PAD_SM = 8
PAD_MD = 12
PAD_LG = 16
PAD_XL = 24

# Corner radius
RADIUS_SM = 6
RADIUS_MD = 8
RADIUS_LG = 12
