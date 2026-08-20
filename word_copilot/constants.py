# ── ICON CACHE ─────────────────────────────────────────────────────────────────
ICON_CACHE_DIR = "icons"
FA_BASE_URL = (
    "https://raw.githubusercontent.com/FortAwesome/Font-Awesome/"
    "6.x/svgs/{style}/{icon}.svg"
)
VALID_ICON_STYLES = ["solid", "regular", "brands"]

# ── COLOR SYSTEM ───────────────────────────────────────────────────────────────
DEFAULT_DSL_COLORS = {
    "a1": "#A4D65E",
    "a2": "#6BAF5B",
    "a3": "#4A7C59",
    "a4": "#2D4A2D",
    "a5": "#7FFF00",
    "a6": "#F0F5E6",
    "bg1": "#FFFFFF",
    "bg2": "#F2F2F2",
    "t1": "#1A1A1A",
    "t2": "#4A4A4A",
}

DSL_COLOR_ALIASES = {
    "a1": "theme_accent1",
    "a2": "theme_accent2",
    "a3": "theme_accent3",
    "a4": "theme_accent4",
    "a5": "theme_accent5",
    "a6": "theme_accent6",
    "bg1": "theme_bg1",
    "bg2": "theme_bg2",
    "t1": "theme_text1",
    "t2": "theme_text2",
}

THEME_COLORS = {
    "theme_bg1": DEFAULT_DSL_COLORS["bg1"],
    "theme_text1": DEFAULT_DSL_COLORS["t1"],
    "theme_bg2": DEFAULT_DSL_COLORS["bg2"],
    "theme_text2": DEFAULT_DSL_COLORS["t2"],
    "theme_accent1": DEFAULT_DSL_COLORS["a1"],
    "theme_accent2": DEFAULT_DSL_COLORS["a2"],
    "theme_accent3": DEFAULT_DSL_COLORS["a3"],
    "theme_accent4": DEFAULT_DSL_COLORS["a4"],
    "theme_accent5": DEFAULT_DSL_COLORS["a5"],
    "theme_accent6": DEFAULT_DSL_COLORS["a6"],
}

WORD_THEME_MAP = {
    "theme_accent1": 5,
    "theme_accent2": 6,
    "theme_accent3": 7,
    "theme_accent4": 8,
    "theme_accent5": 9,
    "theme_accent6": 10,
    "theme_text1": 13,
    "theme_text2": 14,
    "theme_bg1": 12,
    "theme_bg2": 15,
}

# ── UNIT CONSTANTS ────────────────────────────────────────────────────────────
PX_TO_EMU = 9525
PT_TO_EMU = 12700
INCH_TO_EMU = 914400

PAGE_W_PX = 816
PAGE_H_PX = 1056
MARGIN_PX = 72
USABLE_W_PX = PAGE_W_PX - 2 * MARGIN_PX  # 672 px

# ── ELEMENT & DIRECTIVE TYPES ─────────────────────────────────────────────────
FLOW_TYPES = {
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "p",
    "ul",
    "ol",
    "item",
    "hr",
    "br",
    "pagebreak",
    "table",
    "image",
    "icon",
    "svg",
    "textbox",
    "page",
}

EDIT_TYPES = {
    "edit",
    "delete",
    "replace",
    "insert_before",
    "insert_after",
    "insert_at",
    "endblock",
}

HEADING_TYPES = {"h1", "h2", "h3", "h4", "h5", "h6"}

HEADING_SIZES = {
    "h1": 28,
    "h2": 22,
    "h3": 18,
    "h4": 14,
    "h5": 12,
    "h6": 11,
}
