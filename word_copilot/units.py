from word_copilot.constants import PX_TO_EMU, PT_TO_EMU


def px_to_emu(px):
    return int(float(px) * PX_TO_EMU)


def pt_to_emu(pt):
    return int(float(pt) * PT_TO_EMU)


def emu_to_px(emu):
    return emu / PX_TO_EMU


def emu_to_pt(emu):
    return emu / PT_TO_EMU


def pct_to_px(pct_str, total_px):
    s = pct_str.strip()
    if s.endswith("%"):
        return total_px * float(s[:-1]) / 100.0
    return float(s)
