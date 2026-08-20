import os
import re
import hashlib
import tempfile
import requests
import urllib3

from word_copilot.constants import ICON_CACHE_DIR, FA_BASE_URL, VALID_ICON_STYLES

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def download_icon(icon_name, style="solid"):
    if style not in VALID_ICON_STYLES:
        style = "solid"
    style_dir = os.path.join(ICON_CACHE_DIR, style)
    os.makedirs(style_dir, exist_ok=True)
    file_path = os.path.join(style_dir, f"{icon_name}.svg")
    if os.path.exists(file_path):
        return file_path
    url = FA_BASE_URL.format(style=style, icon=icon_name)
    try:
        r = requests.get(url, verify=False, timeout=10)
        if r.status_code == 200:
            with open(file_path, "wb") as f:
                f.write(r.content)
            return file_path
        for alt in VALID_ICON_STYLES:
            if alt == style:
                continue
            r2 = requests.get(
                FA_BASE_URL.format(style=alt, icon=icon_name), verify=False, timeout=10
            )
            if r2.status_code == 200:
                print(f"[icon] '{icon_name}' available as style '{alt}'")
        return None
    except Exception as e:
        print(f"[icon] Download error '{icon_name}': {e}")
        return None


def colorize_svg(svg_path, hex_color):
    try:
        with open(svg_path, "r", encoding="utf-8") as f:
            content = f.read()
        content = re.sub(r'fill="[^"]*"', f'fill="{hex_color}"', content)
        if "fill=" not in content:
            content = content.replace("<svg", f'<svg fill="{hex_color}"', 1)
        out = svg_path.replace(".svg", f"_{hex_color.lstrip('#')}.svg")
        with open(out, "w", encoding="utf-8") as f:
            f.write(content)
        return out
    except Exception as e:
        print(f"[icon] colorize warning: {e}")
        return svg_path


def get_svg_aspect_ratio(svg_path):
    try:
        with open(svg_path, "r", encoding="utf-8") as f:
            content = f.read()
        m = re.search(r'viewBox="[^"]*\s([\d.]+)\s+([\d.]+)"', content)
        if m:
            w, h = float(m.group(1)), float(m.group(2))
            if w > 0 and h > 0:
                return w / h
        mw = re.search(r'width="([\d.]+)', content)
        mh = re.search(r'height="([\d.]+)', content)
        if mw and mh:
            w, h = float(mw.group(1)), float(mh.group(1))
            if w > 0 and h > 0:
                return w / h
    except Exception as e:
        print(f"[icon] aspect-ratio warning: {e}")
    return 1.0


def svg_to_png(svg_path, width_px=64, height_px=64):
    try:
        import cairosvg

        png_path = svg_path.replace(".svg", f"_{width_px}x{height_px}.png")
        cairosvg.svg2png(
            url=svg_path,
            write_to=png_path,
            output_width=width_px,
            output_height=height_px,
        )
        return png_path
    except ImportError:
        return None
    except Exception as e:
        print(f"[icon] svg→png warning: {e}")
        return None


def inject_svg_color(svg_markup, hex_color):
    if not hex_color:
        return svg_markup
    try:
        color = hex_color.strip()
        if not color.startswith("#"):
            color = f"#{color}"
        svg_markup = svg_markup.replace('fill="currentColor"', f'fill="{color}"')
        svg_markup = svg_markup.replace('stroke="currentColor"', f'stroke="{color}"')
        return svg_markup
    except Exception as e:
        print(f"[svg] color injection warning: {e}")
        return svg_markup


def inline_svg_to_tempfile(svg_markup):
    try:
        cache_dir = os.path.join(tempfile.gettempdir(), "word_dsl_svg")
        os.makedirs(cache_dir, exist_ok=True)
        digest = hashlib.sha1(svg_markup.encode("utf-8")).hexdigest()
        path = os.path.join(cache_dir, f"inline_{digest}.svg")
        if not os.path.exists(path):
            with open(path, "w", encoding="utf-8") as f:
                f.write(svg_markup)
        return path
    except Exception as e:
        print(f"[svg] temp file warning: {e}")
        return None
