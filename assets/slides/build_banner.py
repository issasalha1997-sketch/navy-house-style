"""Event banners and social images in the slide style: HTML and CSS rendered to PNG at an exact pixel size.

    python build_banner.py out_folder

Edit the CONTENT block, then run. Two designs are made (A: title block with a divider; B: a card in the primary colour with
labelled arcs), each at 1x (to upload) and 2x (to check). Needs: pip install playwright pillow; playwright install chromium.
"""
import pathlib
import sys

from PIL import ImageFont
from playwright.sync_api import sync_playwright

# ------------------------------------------------------------------ content (replace)
W, H = 796, 265                      # the platform's exact size; check its help page
EYEBROW = "Live session"
TITLE = ["Introducing the", "Team Dashboard"]          # two short lines at most
SUBTITLE = ["See how each team is doing,", "side by side and in one place."]
ARC_LABELS = ["By team", "By month", "By region"]     # the three things the arcs (B) or the strip (A) stand for
LOGO = None                          # path to an official logo PNG, only when asked (shown on white, unaltered)

# ------------------------------------------------------------------ style
THEME = "green"                      # or "navy"
PRIMARY, LINE_TINT = {"green": ("#00684A", "#A6CAC0"), "navy": ("#243168", "#D3D8EA")}[THEME]
INK = "#1D1D1B"
MARGIN, BAR = 30, 10


def _find_font(names):
    """The first of names found in the folder named by SLIDE_FONT_DIR or the usual font folders, or None."""
    import os
    dirs = [os.environ.get("SLIDE_FONT_DIR", ""), os.path.join(os.environ.get("WINDIR", ""), "Fonts"), "/Library/Fonts",
            "/System/Library/Fonts/Supplemental", os.path.expanduser("~/Library/Fonts"),
            "/usr/share/fonts/truetype/msttcorefonts"]
    for d in filter(None, dirs):
        for n in names:
            if os.path.exists(os.path.join(d, n)):
                return os.path.join(d, n)
    return None


FONT, FONT_BOLD = _find_font(["tahoma.ttf", "Tahoma.ttf"]), _find_font(["tahomabd.ttf", "Tahoma Bold.ttf"])


def text_w(text, px, bold=True):
    """Width of a line in CSS px, measured with the real font, so a title that would wrap stops the build."""
    try:
        return ImageFont.truetype(FONT_BOLD if bold else FONT, 400).getlength(text) / 400 * px
    except (OSError, AttributeError, TypeError):   # no Tahoma on this machine: a cautious estimate instead
        return len(text) * px * (0.62 if bold else 0.56)


CSS = f"""
*{{margin:0;padding:0;box-sizing:border-box}}
html,body{{width:{W}px;height:{H}px;overflow:hidden;background:#fff}}
body{{font-family:Tahoma,Verdana,sans-serif;color:{INK};position:relative}}
.abs{{position:absolute}}
.title{{font-weight:bold;color:{PRIMARY};letter-spacing:-0.2px}}
"""


def page(body):
    return f"<!doctype html><html><head><meta charset='utf-8'><style>{CSS}</style></head><body>{body}</body></html>"


def logo_html(height):
    """An official logo as a data URI, trimmed of its transparent margin only (never recoloured or stretched).
    Logos sit on white only; check the smallest legible size by zooming in on the fine detail."""
    import base64
    import io
    from PIL import Image
    im = Image.open(LOGO)
    im = im.crop(im.getchannel("A").getbbox()) if im.mode == "RGBA" else im
    buf = io.BytesIO()
    im.save(buf, "PNG")
    width = round(height * im.width / im.height)
    return f'<img src="data:image/png;base64,{base64.b64encode(buf.getvalue()).decode()}" ' \
           f'style="display:block;width:{width}px;height:{height}px">', width


def design_a():
    """Title block under a top bar in the primary colour: optional logo and a thin divider, then eyebrow, title and subtitle; a rule
    and a strip of up to four labels (what the session covers) close it."""
    x = MARGIN
    logo = ""
    if LOGO:
        img, lw = logo_html(64)
        logo = (f'<div class="abs" style="left:{MARGIN}px;top:{MARGIN + 30}px">{img}</div>'
                f'<div class="abs" style="left:{MARGIN + lw + 26}px;top:{MARGIN + 6}px;width:2px;height:150px;'
                f'background:{PRIMARY}"></div>')
        x = MARGIN + lw + 26 + 2 + 26
    tw = W - MARGIN - x
    for line in TITLE:
        assert text_w(line, 40) < tw, "title line too long: " + line
    rule_y = H - 57
    widths = [13 + 8 + text_w(lab, 14, bold=False) for lab in ARC_LABELS]
    gap = (W - 2 * MARGIN - sum(widths)) / max(1, len(ARC_LABELS) - 1)
    assert gap > 20, "labels too long for the strip"
    strip, lx = "", MARGIN
    for lab, w in zip(ARC_LABELS, widths):
        strip += (f'<div class="abs" style="left:{lx:.1f}px;top:{rule_y + 17}px;width:13px;height:13px;'
                  f'background:{PRIMARY}"></div><div class="abs" style="left:{lx + 21:.1f}px;top:{rule_y + 13}px;'
                  f'font-size:14px;line-height:21px">{lab}</div>')
        lx += w + gap
    body = (f'<div class="abs" style="left:0;top:0;width:{W}px;height:{BAR}px;background:{PRIMARY}"></div>{logo}'
            f'<div class="abs" style="left:{x}px;top:{MARGIN + 2}px;width:{tw}px">'
            f'<div style="font:bold 15px/20px Tahoma,Verdana,sans-serif;color:{PRIMARY}">{EYEBROW}</div>'
            f'<div class="title" style="font-size:40px;line-height:44px;margin-top:4px">{"<br>".join(TITLE)}</div>'
            f'<div style="font-size:16px;line-height:21px;margin-top:8px">{"<br>".join(SUBTITLE)}</div></div>'
            f'<div class="abs" style="left:{MARGIN}px;top:{rule_y}px;width:{W - 2 * MARGIN}px;height:1px;'
            f'background:{LINE_TINT}"></div>{strip}')
    return page(body)


def design_b():
    """A card in the primary colour with white text on the left; three labelled concentric arcs that bleed off the bottom edge."""
    card_w, card_h = 380, H - 2 * MARGIN
    for line in TITLE:
        assert text_w(line, 36) < card_w - 56, "title line too long for the card: " + line
    stroke, radii = 14, (140, 108, 76)
    cx = round(W - MARGIN - 40 - max(text_w(l, 16) for l in ARC_LABELS))   # longest label ends on the right margin
    cy = H - 3
    assert cx - radii[0] - stroke / 2 > MARGIN + card_w + 24, "arcs would touch the card"
    arcs = []
    for r, label in zip(radii, ARC_LABELS):
        arcs.append(f'<path d="M {cx - r} {H + 4} L {cx - r} {cy} A {r} {r} 0 0 1 {cx} {cy - r}" fill="none" '
                    f'stroke="{PRIMARY}" stroke-width="{stroke}" stroke-linecap="butt"/>')    # starts below the edge: bleeds
        y = cy - r
        arcs.append(f'<rect x="{cx + 14}" y="{y - 4.5}" width="9" height="9" fill="{PRIMARY}"/>'
                    f'<text x="{cx + 40}" y="{y + 6}" font-family="Tahoma,Verdana,sans-serif" font-weight="bold" '
                    f'font-size="16" fill="{PRIMARY}">{label}</text>')
    body = (f'<div class="abs" style="left:{MARGIN}px;top:{MARGIN}px;width:{card_w}px;height:{card_h}px;'
            f'background:{PRIMARY};border-radius:8px"></div>'
            f'<div class="abs" style="left:{MARGIN + 28}px;top:{MARGIN + 24}px;width:{card_w - 56}px;color:#fff">'
            f'<div style="font:bold 15px/20px Tahoma,Verdana,sans-serif">{EYEBROW}</div>'
            f'<div style="font-weight:bold;font-size:36px;line-height:40px;margin-top:6px;letter-spacing:-0.2px">{"<br>".join(TITLE)}</div>'
            f'<div style="font-size:15.5px;line-height:21px;margin-top:10px">{"<br>".join(SUBTITLE)}</div></div>'
            f'<svg class="abs" style="left:0;top:0" width="{W}" height="{H}" viewBox="0 0 {W} {H}">{"".join(arcs)}</svg>')
    return page(body)


def render(out):
    out = pathlib.Path(out)
    (out / "check-2x").mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        # grey anti-aliasing only: ClearType's coloured fringes turn into halos once the image is scaled
        browser = p.chromium.launch(args=["--disable-lcd-text"])
        for name, fn in (("Banner A - title block", design_a), ("Banner B - arcs", design_b)):
            html = fn()
            (out / f"{name}.html").write_text(html, encoding="utf-8")
            for scale, folder in ((1, out), (2, out / "check-2x")):
                pg = browser.new_page(viewport={"width": W, "height": H}, device_scale_factor=scale)
                pg.set_content(html, wait_until="load")
                pg.wait_for_timeout(150)
                pg.screenshot(path=str(folder / f"{name}.png"), clip={"x": 0, "y": 0, "width": W, "height": H})
                pg.close()
        browser.close()
    print("written to", out)


if __name__ == "__main__":
    render(sys.argv[1] if len(sys.argv) > 1 else "banners")
