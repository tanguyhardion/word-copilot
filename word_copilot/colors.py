import re
from word_copilot.constants import (
    DEFAULT_DSL_COLORS,
    DSL_COLOR_ALIASES,
    THEME_COLORS,
    WORD_THEME_MAP,
)


def clamp(v, lo=0, hi=255):
    return max(lo, min(hi, int(round(v))))


def hex_to_rgb(hex_color):
    try:
        h = re.sub(r"[^0-9A-Fa-f]", "", hex_color.strip().lstrip("#"))
        if len(h) == 3:
            h = "".join(c * 2 for c in h)
        if len(h) != 6:
            raise ValueError(f"bad hex: '{h}'")
        return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))
    except Exception as e:
        print(f"hex_to_rgb warning: {e}")
        return (128, 128, 128)


def rgb_to_hex(r, g, b):
    return f"#{clamp(r):02X}{clamp(g):02X}{clamp(b):02X}"


def bgr_int_to_hex(rgb_int):
    try:
        v = int(rgb_int)
        r = v & 0xFF
        g = (v >> 8) & 0xFF
        b = (v >> 16) & 0xFF
        return f"#{r:02X}{g:02X}{b:02X}"
    except Exception as e:
        print(f"bgr_int_to_hex warning: {e}")
        return None


def lighten_hex(hex_color, amount):
    r, g, b = hex_to_rgb(hex_color)
    return rgb_to_hex(
        r + (255 - r) * amount, g + (255 - g) * amount, b + (255 - b) * amount
    )


def darken_hex(hex_color, amount):
    r, g, b = hex_to_rgb(hex_color)
    return rgb_to_hex(r * (1 - amount), g * (1 - amount), b * (1 - amount))


def dsl_resolve_color(c):
    """Resolve a DSL color token to a hex string."""
    if not c:
        return None
    c = c.strip().lower()
    if c in DSL_COLOR_ALIASES:
        token = DSL_COLOR_ALIASES[c]
        return THEME_COLORS.get(token, DEFAULT_DSL_COLORS.get(c))
    m = re.match(r"^(a[1-6]|bg[12]|t[12])_(l1|l2|d1|d2)$", c)
    if m:
        base, variant = m.groups()
        token = DSL_COLOR_ALIASES.get(base)
        base_hex = THEME_COLORS.get(token) if token else None
        if not base_hex:
            return None
        if variant == "l1":
            return lighten_hex(base_hex, 0.35)
        if variant == "l2":
            return lighten_hex(base_hex, 0.60)
        if variant == "d1":
            return darken_hex(base_hex, 0.25)
        if variant == "d2":
            return darken_hex(base_hex, 0.45)
    if re.match(r"^#[0-9a-f]{3}$", c):
        return "#" + "".join(ch * 2 for ch in c[1:])
    if re.match(r"^#[0-9a-f]{6}$", c):
        return c.upper()
    return None


def is_theme_color_token(c):
    return c in WORD_THEME_MAP


def get_active_word_theme_colors():
    """Read theme colors from the active Word document via COM."""
    colors = dict(THEME_COLORS)
    try:
        import pythoncom, win32com.client

        pythoncom.CoInitialize()
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            if word.Documents.Count == 0:
                return colors
            doc = word.ActiveDocument
            scheme = doc.DocumentTheme.ThemeColorScheme
            for token, idx in WORD_THEME_MAP.items():
                try:
                    hex_val = bgr_int_to_hex(scheme.Colors(idx).RGB)
                    if hex_val:
                        colors[token] = hex_val
                except Exception as inner:
                    print(f"[theme] {token} @ {idx}: {inner}")
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        print(f"[theme] fallback defaults: {e}")
    return colors


def refresh_word_theme_colors():
    global THEME_COLORS
    THEME_COLORS.update(get_active_word_theme_colors())
    print("[theme] Word theme colors refreshed:")
    for k, v in THEME_COLORS.items():
        print(f"  {k} = {v}")
