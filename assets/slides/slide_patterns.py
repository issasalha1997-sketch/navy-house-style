"""Slide patterns: the devices that make a slide show the real thing instead of listing it.

Each function draws one device with the kit (slide_kit.py) and returns where it ends, so the next device can be
placed against it. Positions are inches. See references/patterns.md for when to use each one, and demo_deck.py
for every pattern on a slide.
"""
import math

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

from slide_kit import (DASH, INK, L, P, R, SH, SW, WHITE, B, N, layout_lines, lines_height_in, readable,
                       text_width_in)


# ---------------------------------------------------------------- headers
def fit_size(text, width, max_size, min_size, bold=True):
    """The largest whole point size (from max_size down to min_size) at which text fits one line of width."""
    size = max_size
    while size > min_size and text_width_in(text, size, bold) * 1.05 > width:
        size -= 1
    assert text_width_in(text, size, bold) * 1.05 <= width, f"title does not fit even at {min_size} pt: {text!r}"
    return size


def header(d, s, title, lead=None, band=True, logo=None, max_size=32, min_size=26):
    """A top band in the primary colour (optional), the title on one line, and an optional lead of one or two short sentences
    (one per line). logo: path to an official logo, shown top right. Returns the y where content can start."""
    if band:
        d.box(s, 0, 0, SW, 0.49, fill=P.primary, name="Top band")
    ty = 0.72 if band else 0.5
    w = R - L
    if logo:                                                    # measured first, added after the title (read order)
        (lx, ly, lw, lh), _ = d.logo_box(logo, R - 1.6, ty - 0.02, 1.6, 0.8, align="right", valign="top")
        w = lx - L - 0.3
    size = fit_size(title, w, max_size, min_size)
    d.title(s, L, ty, w, 0.62, title, size)
    if logo:
        d.logo(s, logo, R - 1.6, ty - 0.02, 1.6, 0.8, align="right", valign="top")
    y = ty + 0.62
    if lead:
        _, h = d.sentence_lines(s, L, y + 0.12, w, lead, 15)
        y += 0.12 + h
    return y + 0.45


def arcs_svg(w, h, cx, cy, radii, colour, width_pt=16, start=200, end=300):
    """Concentric arcs as an SVG in a box w x h inches (cx, cy and radii in the same box's inches); parts outside the
    box are clipped. Used as a quiet motif on a panel in the primary colour."""
    u = 100.0
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w * u:.1f} {h * u:.1f}" '
           f'width="{w * u:.1f}" height="{h * u:.1f}">']
    sw = width_pt / 72 * u
    for r in radii:
        a0, a1 = math.radians(start), math.radians(end)
        x0, y0 = (cx + r * math.cos(a0)) * u, (cy + r * math.sin(a0)) * u
        x1, y1 = (cx + r * math.cos(a1)) * u, (cy + r * math.sin(a1)) * u
        large = 1 if (end - start) % 360 > 180 else 0
        out.append(f'<path d="M {x0:.2f} {y0:.2f} A {r * u:.2f} {r * u:.2f} 0 {large} 1 {x1:.2f} {y1:.2f}" '
                   f'fill="none" stroke="#{colour}" stroke-width="{sw:.2f}" stroke-linecap="butt"/>')
    out.append("</svg>")
    return "".join(out)


def side_panel(d, s, pw, title_lines, subtitle_lines=None, lead=None, motif=True, title_size=30):
    """A full-height panel in the primary colour on the left holding the words: the question as the title (hand-set lines), who is
    asking, a short lead. A quiet arc motif sits at the bottom. Returns the x where the white side's content starts."""
    d.box(s, 0, 0, pw, SH, fill=P.primary, name="Colour panel")
    if motif:
        my = 4.5
        d.svg_picture(s, arcs_svg(pw, SH - my, 1.05, SH - my + 0.85, (1.55, 2.1, 2.65), P.decor), 0, my, pw, SH - my,
                      name="Panel motif", descr="decorative arcs")
    px, tw = 0.55, pw - 0.55 - 0.25
    th = lines_height_in(len(title_lines), title_size, 0.95) + 0.05
    d.title(s, px, 0.85, tw, th, " ".join(title_lines), title_size, colour=WHITE, lines=title_lines,
            anchor=MSO_ANCHOR.TOP, spacing=0.95)
    y = 0.85 + th
    if subtitle_lines:
        h = lines_height_in(len(subtitle_lines), 15) + 0.04
        d.text_lines(s, px, y + 0.3, tw, h, [[N(t, 15, WHITE)] for t in subtitle_lines], name="Subtitle")
        y += 0.3 + h
    if lead:
        _, h = d.para(s, px, y + 0.3, tw, [N(lead, 14, WHITE)], 6, "Lead")
        y += 0.3 + h
    assert y < 5.5, "panel text reaches the motif"
    return pw + 0.45


# ---------------------------------------------------------------- the real thing, drawn
def calendar(d, s, ax0, ax1, top, months, windows, year=None, year_x=None, name_h=0.34, bar_h=0.46, gap=0.1):
    """Months on an axis with each time window as a coloured bar sitting on it, named above in its own colour.
    windows: [{"start": 1, "end": 4, "label": "February to April", "name": "Pilot one", "colour": P.primary}], where
    start and end index into months (end exclusive). Returns the y under the month labels."""
    cw = (ax1 - ax0) / len(months)
    bar_y = top + name_h + gap
    axis_y = bar_y + bar_h
    for wdw in windows:
        x, w = ax0 + wdw["start"] * cw, (wdw["end"] - wdw["start"]) * cw
        if wdw.get("name"):
            d.one(s, x, top, w + 1.0, name_h, wdw["name"], 13, True, readable(wdw["colour"], 13, True),
                  anchor=MSO_ANCHOR.BOTTOM,
                  name=f"Window name {wdw['name']}")
        d.box(s, x, bar_y, w, bar_h, fill=wdw["colour"], name=f"Window {wdw['label']}")
        d.one(s, x, bar_y, w, bar_h, wdw["label"], 14, True, WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
              name=f"Window dates {wdw['label']}")
    d.line(s, ax0, axis_y, ax1, axis_y, P.primary, 1.25, name="Calendar axis")
    for i in range(len(months) + 1):
        d.line(s, ax0 + i * cw, axis_y, ax0 + i * cw, axis_y + 0.08, P.primary, 1.0, name=f"Tick {i}")
    for i, m in enumerate(months):
        d.one(s, ax0 + i * cw, axis_y + 0.1, cw, 0.28, m, 12, False, P.primary, align=PP_ALIGN.CENTER, name=f"Month {m}")
    if year:
        d.one(s, year_x, axis_y + 0.1, ax0 - year_x - 0.05, 0.28, year, 12, True, P.primary, name="Year")
    return axis_y + 0.38


def notification(d, s, x, y, w, app, right, message, border, size=12, pad=0.12, tile=0.3, right_bold=True,
                 right_colour=None, h=None):
    """A phone notification card: app tile with a bell, app name, a right-aligned label, then the exact wording.
    Returns its height (pass h to make two cards the same height)."""
    right_colour = right_colour or P.primary
    msg_w = w - 2 * pad
    lines = layout_lines([N(message, size)], msg_w)
    head_h = tile + 0.1
    h = h or (pad + head_h + lines_height_in(len(lines), size) + pad + 0.02)
    d.box(s, x, y, w, h, fill=WHITE, line=border, line_w=1.25, r=0.12, name="Notification")
    d.box(s, x + pad, y + pad, tile, tile, fill=P.primary, r=tile * 0.24, name="App tile")
    d.icon(s, "bell-simple", x + pad + tile * 0.18, y + pad + tile * 0.18, tile * 0.64, WHITE, descr="bell icon")
    rw = text_width_in(right, size, right_bold) * 1.05 + 0.1
    assert text_width_in(app, size, True) + tile + 0.08 + rw + 0.15 < msg_w, "card header too wide"
    d.one(s, x + pad + tile + 0.08, y + pad, msg_w - tile - 0.08 - rw, tile, app, size, True, INK,
          anchor=MSO_ANCHOR.MIDDLE, name="App name")
    d.one(s, x + w - pad - rw, y + pad, rw, tile, right, size, right_bold, right_colour, align=PP_ALIGN.RIGHT,
          anchor=MSO_ANCHOR.MIDDLE, name="Notification label")
    d.text_lines(s, x + pad, y + pad + head_h, msg_w, lines_height_in(len(lines), size) + 0.04, lines, name="Wording")
    return h


def notification_height(w, message, size=12, pad=0.12, tile=0.3):
    return pad + tile + 0.1 + lines_height_in(len(layout_lines([N(message, size)], w - 2 * pad)), size) + pad + 0.02


def phone(d, s, x, y, w, label, cards, app):
    """A phone rising from the bottom edge, holding the real wording people would see.
    cards: [{"right": "Mar 2026", "message": "...", "border": P.primary}]. Keep w about 3.6 in so the wording stays 12 pt."""
    d.shape(s, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, x, y, w, SH - y, fill=P.primary, name="Phone", adj=[0.42 / w, 0.0])
    d.shape(s, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, x + 0.1, y + 0.1, w - 0.2, SH - y - 0.1, fill=P.surface,
            name="Phone screen", adj=[0.33 / (w - 0.2), 0.0])
    d.box(s, x + w / 2 - 0.4, y + 0.2, 0.8, 0.13, fill=P.primary, r=0.065, name="Phone camera")
    d.one(s, x + 0.1, y + 0.42, w - 0.2, 0.3, label, 12, True, P.primary, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
          name="Phone label")
    cx, cw, cy = x + 0.24, w - 0.48, y + 0.8
    for c in cards:
        cy += notification(d, s, cx, cy, cw, app, c["right"], c["message"], c.get("border", P.primary)) + 0.12
    assert cy - 0.12 <= SH - 0.22, f"notifications reach the bottom edge ({cy - 0.12:.2f})"


def icon_text(d, s, icon, x, y, w, runs, max_lines, name, icon_size=0.4, gap=0.14, colour=None, balance=False):
    """An icon beside a short text, centred on each other. Returns the height used."""
    lines = layout_lines(runs, w - icon_size - gap, balance=balance)
    assert len(lines) <= max_lines, f"{name}: needs {len(lines)} lines, room for {max_lines}"
    size = max(r[1] for r in runs)
    th = lines_height_in(len(lines), size)
    top = y + max(0.0, (icon_size - th) / 2)
    d.icon(s, icon, x, y + max(0.0, (th - icon_size) / 2), icon_size, colour)
    d.text_lines(s, x + icon_size + gap, top, w - icon_size - gap, th + 0.04, lines, name=name)
    return max(th, icon_size)


def icon_row(d, s, items, x, y, w, colour=None, text_colour=INK, size=13, icon=0.4, max_lines=2, gap_icon=0.14):
    """Up to four short offers side by side, each an icon and a balanced two-line label; spare width is shared
    equally so the gaps look even. items: [(icon_name, text)]."""
    gap_item = 0.25
    need = []
    for _, text in items:
        tw = 0.8
        while len(layout_lines([N(text, size)], tw, balance=True)) > max_lines:
            tw += 0.02
        need.append(tw)
    spare = w - sum(need) - len(items) * (icon + gap_icon) - (len(items) - 1) * gap_item
    assert spare >= 0, f"icon row needs {-spare:.2f} in more"
    ix = x
    for i, ((ic, text), tw) in enumerate(zip(items, need)):
        tw += spare / len(items)
        d.icon(s, ic, ix, y, icon, colour)
        lines = layout_lines([N(text, size, text_colour)], tw, balance=True)
        th = lines_height_in(len(lines), size)
        d.text_lines(s, ix + icon + gap_icon, y + (icon - th) / 2, tw, th + 0.04, lines, name=f"Offer {i + 1}")
        ix += icon + gap_icon + tw + gap_item
    return icon


def icon_column(d, s, items, x, y, w, size=13, icon=0.38, gap=0.16):
    """The same offers stacked, for a narrow column. Returns the y under the last item."""
    for i, (ic, text) in enumerate(items):
        y += icon_text(d, s, ic, x, y, w, [N(text, size)], 2, f"Offer {i + 1}", icon) + gap
    return y - gap


def run_in(d, s, icon, x, y, w, lead_in, text, size=13, max_lines=3, icon_size=0.4):
    """A secondary point as one sentence with a bold lead-in in the primary colour ("Next step: ..."). Returns its height."""
    return icon_text(d, s, icon, x, y, w, [B(lead_in, size, P.primary), N(" " + text, size)], max_lines, "Secondary point",
                     icon_size)


def heading_square(d, s, x, y, w, text, size=16):
    """A section heading led by a small square in the primary colour (for parts that are not steps, so no numbers).
    Returns the y where its content starts."""
    d.square(s, x + 0.07, y + 0.18, 0.14, name="Heading square")
    d.one(s, x + 0.3, y, w - 0.3, 0.36, text, size, True, P.primary, anchor=MSO_ANCHOR.MIDDLE, name="Heading")
    return y + 0.46


def closing_band(d, s, y, heading, items, thanks):
    """A full-width band in the primary colour that ends the slide: what you offer (white icons) and a right-aligned thank-you."""
    d.box(s, 0, y, SW, SH - y, fill=P.primary, name="Closing band")
    d.one(s, L, y + 0.2, 6, 0.3, heading, 13, True, WHITE, name="Band heading")
    icon_row(d, s, items, L, y + 0.6, 8.1, WHITE, WHITE, size=13, icon=0.36)
    tx = L + 8.1 + 0.3
    lines = layout_lines([B(thanks, 16, WHITE)], R - tx)
    assert len(lines) <= 2, "thank-you line too long"
    th = lines_height_in(len(lines), 16)
    d.text_lines(s, tx, y + (SH - y - th) / 2, R - tx, th + 0.04, lines, align=PP_ALIGN.RIGHT, name="Thanks")


# ---------------------------------------------------------------- numbers, purpose, stages, structure
def icon_tile(d, s, icon, x, y, size, fill=None, colour=None):
    """An icon on a small rounded square of the pale surface colour (the glyph is 56% of the tile)."""
    d.box(s, x, y, size, size, fill=fill or P.surface, r=size * 0.18, name="Icon tile")
    g = size * 0.56
    d.icon(s, icon, x + (size - g) / 2, y + (size - g) / 2, g, colour)


def stats_row(d, s, y, stats, x=L, w=None):
    """Three headline numbers in a row: icon tile, a small qualifier above ("Nearly"), the number big, a label
    under it, thin rules between. stats: [(icon, qualifier, number, label)]. Returns the y under the row."""
    w = w or (R - x)
    n, gap = len(stats), 0.3
    iw = (w - gap * (n - 1)) / n
    for i, (ic, qual, num, label) in enumerate(stats):
        ix = x + i * (iw + gap)
        icon_tile(d, s, ic, ix, y + 0.14, 0.82)
        tx, tw = ix + 1.02, iw - 1.02
        if qual:
            d.one(s, tx, y, tw, 0.3, qual, 14, True, P.primary, name="Qualifier")
        d.one(s, tx, y + 0.28, tw, 0.6, num, 32, True, P.primary, anchor=MSO_ANCHOR.MIDDLE, name="Number")
        d.one(s, tx, y + 0.9, tw, 0.3, label, 13, False, INK, name="Number label")
        if i:
            gx = ix - gap / 2
            d.line(s, gx, y + 0.1, gx, y + 1.05, P.hairline, 1.25, name="Stat divider")
    return y + 1.25


def need_and_aim(d, s, x, y, w, need_label, need_runs, aim_label, aim_text, need_icon="magnifying-glass",
                 aim_icon="target"):
    """The problem in a pale card, an arrow down, then the purpose in a card in the primary colour: need, then aim.
    need_runs: runs with one bold key phrase. Returns the y under the aim card."""
    tx, tw = x + 1.05, w - 1.05 - 0.3
    need_lines = layout_lines(need_runs, tw)
    nh = 0.25 + 0.32 + lines_height_in(len(need_lines), 15) + 0.3
    d.box(s, x, y, w, nh, fill=P.surface, r=0.08, name="Need card")
    d.icon(s, need_icon, x + 0.3, y + 0.27, 0.5, P.primary)
    d.one(s, tx, y + 0.22, tw, 0.3, need_label, 13, True, P.primary, name="Need label")
    d.text_lines(s, tx, y + 0.55, tw, lines_height_in(len(need_lines), 15) + 0.04, need_lines, name="Need")
    ay = y + nh + 0.42
    d.line(s, x + 0.55, y + nh + 0.06, x + 0.55, ay - 0.06, P.primary, 2.25, name="Arrow", arrow_end=True)
    aim_lines = layout_lines([B(aim_text, 17, WHITE)], tw)
    ah = 0.25 + 0.32 + lines_height_in(len(aim_lines), 17) + 0.3
    d.box(s, x, ay, w, ah, fill=P.primary, r=0.08, name="Aim card")
    d.icon(s, aim_icon, x + 0.3, ay + 0.27, 0.5, WHITE)
    d.one(s, tx, ay + 0.22, tw, 0.3, aim_label, 13, True, WHITE, name="Aim label")
    d.text_lines(s, tx, ay + 0.55, tw, lines_height_in(len(aim_lines), 17) + 0.04, aim_lines, name="Aim")
    return ay + ah


def labelled_arcs(d, s, cx, cy, items, stroke_pt=15, size=20):
    """Three concentric quarter arcs, each ending in a small square and a bold label: three qualities, three ways
    in, three audiences. items: [(radius, label)] from the outside in. Native arcs, so they stay editable."""
    for r, label in items:
        d.arc(s, cx, cy, r, 180, 270, stroke_pt, P.primary, name=f"Arc {label}")
        top = cy - r
        d.square(s, cx + 0.3, top, 0.13, name=f"Marker {label}")
        d.one(s, cx + 0.55, top - 0.2, 3.0, 0.4, label, size, True, P.primary, anchor=MSO_ANCHOR.MIDDLE, name=f"Label {label}")


def chevron_stages(d, s, y, parts, x=L, w=None, h=0.68, gutter=0.38):
    """Stages as one chevron bar: done (primary), now (a lighter step) and next (white, dashed outline). A short note sits
    under each stage. parts: [{"title", "icon", "state", "note"}]. Returns (column_xs, column_width, y_under_notes)."""
    w = w or (R - x)
    n = len(parts)
    cw = (w - gutter * (n - 1)) / n
    cols = [x + i * (cw + gutter) for i in range(n)]
    for i, (p, cx) in enumerate(zip(parts, cols)):
        last = i == n - 1
        bx = cx if i == 0 else cx - 0.05
        bw = (x + w - bx) if last else (cw + 0.3 if i == 0 else cw + 0.35)
        kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        state = p.get("state", "done")
        if state == "next":
            d.shape(s, kind, bx, y, bw, h, fill=WHITE, line=P.primary, line_w=1.5, dash=DASH, name=f"Stage {p['title']}",
                    adj=[0.32])
            fg = P.primary
        else:
            d.shape(s, kind, bx, y, bw, h, fill=P.primary if state == "done" else P.now, name=f"Stage {p['title']}",
                    adj=[0.32])
            fg = WHITE
        ix = bx + (0.2 if i == 0 else 0.42)
        d.icon(s, p["icon"], ix, y + (h - 0.38) / 2, 0.38, fg)
        d.one(s, ix + 0.5, y, bw - (ix - bx) - 0.5 - 0.3, h, p["title"], 16, True, fg, anchor=MSO_ANCHOR.MIDDLE,
              name=f"Stage title {p['title']}")
        if p.get("note"):
            d.one(s, cx, y + h + 0.12, cw, 0.3, p["note"], 12, False, INK, name=f"Stage note {p['title']}")
    return cols, cw, y + h + 0.52


def gutter_dividers(d, s, cols, cw, y1, y2):
    """Dotted rules down the middle of each gutter, so the parts read as clearly separate."""
    for a, b in zip(cols, cols[1:]):
        d.vdivider(s, (a + cw + b) / 2, y1, y2, name="Gutter divider")


def card(d, s, x, y, w, h, state="live", r=0.08, name="Card"):
    """live: white with a hairline border; next: white with a dashed border in the primary colour (not here yet); quiet: pale
    surface fill, no border. One radius everywhere."""
    if state == "next":
        return d.box(s, x, y, w, h, fill=WHITE, line=P.primary, line_w=1.25, dash=DASH, r=r, name=name)
    if state == "quiet":
        return d.box(s, x, y, w, h, fill=P.surface, r=r, name=name)
    return d.box(s, x, y, w, h, fill=WHITE, line=P.hairline, line_w=1.25, r=r, name=name)


def two_zone_structure(d, s, top, levels, joined, members, x=L, box_w=4.6, box_h=0.9, gap=0.3, cols=4,
                       member_colours=None):
    """A structure chart in two zones: the levels down the left, each a box in the primary colour with its name and role inside (no
    leader lines needed), joined by a short spine; one level's members sit in a panel to the right, joined to it
    by a line. levels: [(name, role)]; joined: index of the level that has members; members: [name].
    member_colours: {member: colour}, only when each member has its own colour that means something (a brand or
    category colour); otherwise every band is in the primary colour."""
    cx = x + box_w / 2
    ys = [top + i * (box_h + gap) for i in range(len(levels))]
    for i in range(len(levels) - 1):
        d.line(s, cx, ys[i] + box_h, cx, ys[i + 1], P.primary, 3, name="Spine")
    for (name, role), by in zip(levels, ys):
        d.box(s, x, by, box_w, box_h, fill=P.primary, r=0.12, name=f"Level {name}")
        d.text_lines(s, x + 0.18, by + 0.14, box_w - 0.36, 0.3, [[B(name, 14, WHITE)]], name=f"Level name {name}")
        d.text_lines(s, x + 0.18, by + 0.48, box_w - 0.36, 0.28, [[N(role, 12, WHITE)]], name=f"Level role {name}")
    rows = math.ceil(len(members) / cols)
    px0, pad, cgap = x + box_w + 0.5, 0.2, 0.14
    pw = R - px0
    mw = (pw - 2 * pad - (cols - 1) * cgap) / cols
    mh = 0.9
    ph = 2 * pad + rows * mh + (rows - 1) * cgap
    jy = ys[joined] + box_h / 2
    py = max(top, min(jy - ph / 2, SH - 0.4 - ph))
    d.box(s, px0, py, pw, ph, line=P.primary, line_w=3, r=0.2, name="Members panel")
    d.line(s, x + box_w, jy, px0, jy, P.primary, 3, name="Joining line")
    for i, m in enumerate(members):
        row, col = divmod(i, cols)
        in_row = min(cols, len(members) - row * cols)
        rx0 = px0 + pad + (cols - in_row) * (mw + cgap) / 2          # a short last row is centred
        mx, my = rx0 + col * (mw + cgap), py + pad + row * (mh + cgap)
        d.box(s, mx, my, mw, mh, fill=WHITE, line=P.hairline, line_w=1.0, r=0.1, name=f"Member {m}")
        band = member_colours.get(m, P.primary) if member_colours else P.primary
        d.box(s, mx + 0.12, my + 0.12, mw - 0.24, 0.06, fill=band, name="Member band")
        lines = layout_lines([B(m, 12, P.primary)], mw - 0.24, balance=True)
        assert len(lines) <= 2, f"member name too long: {m}"
        th = lines_height_in(len(lines), 12)
        d.text_lines(s, mx + 0.12, my + 0.18 + (mh - 0.18 - th) / 2, mw - 0.24, th + 0.04, lines,
                     align=PP_ALIGN.CENTER, name=f"Member name {m}")
    return max(ys[-1] + box_h, py + ph)


def title_slide(d, s, title_lines, subtitle, byline, parts=None, pw=6.4):
    """Opening slide: a wide panel in the primary colour with the title, a subtitle and who/when, over the quiet arc motif.
    parts: up to three short labels for what the talk covers, drawn on the white side as labelled arcs."""
    d.box(s, 0, 0, pw, SH, fill=P.primary, name="Colour panel")
    my = 3.9
    d.svg_picture(s, arcs_svg(pw, SH - my, 1.3, SH - my + 1.1, (1.9, 2.55, 3.2), P.decor), 0, my, pw, SH - my,
                  name="Panel motif", descr="decorative arcs")
    tw = pw - 0.7 - 0.4
    th = lines_height_in(len(title_lines), 40, 0.95) + 0.05
    d.title(s, 0.7, 0.9, tw, th, " ".join(title_lines), 40, colour=WHITE, lines=title_lines, anchor=MSO_ANCHOR.TOP,
            spacing=0.95)
    y = 0.9 + th + 0.3
    _, h = d.para(s, 0.7, y, tw, [N(subtitle, 18, WHITE)], 3, "Subtitle")
    d.one(s, 0.7, y + h + 0.3, tw, 0.32, byline, 14, False, WHITE, name="Byline")
    if parts:                                                    # after the title, so it is read after it
        assert len(parts) <= 3, "three parts at most"
        radii = (2.4, 1.85, 1.3)[:len(parts)]
        labelled_arcs(d, s, pw + 0.6 + radii[0], SH - 0.55, list(zip(radii, parts)), stroke_pt=13, size=18)
