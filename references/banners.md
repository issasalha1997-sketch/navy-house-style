# Event banners and social images

For a banner or card image in the same style (an event page, a registration page, a post). Starter: `assets/slides/build_banner.py` (two designs, HTML and CSS rendered to PNG at an exact pixel size; needs `pip install playwright pillow` and `playwright install chromium`). Green (#00684A) by default; set `THEME = "navy"` at the top of the builder for navy.

## How

- Lay the banner out in HTML with absolute pixel positions on a page exactly the platform's size, and render it with a headless browser (Playwright) at 1x for upload and 2x for checking. Keep the HTML beside the PNG so it can be re-rendered.
- **Launch Chromium with `--disable-lcd-text`.** Otherwise text is drawn with ClearType colour fringes that turn into halos once the image is scaled on a phone or a high-DPI screen.
- Measure every title and subtitle line with the real font before rendering, and stop the build if a line would wrap or collide. The builder looks for Tahoma in the usual font folders and in `SLIDE_FONT_DIR`; without it, it uses a cautious width estimate, so check the 2x render closely.
- Export a flat RGB PNG (no transparency, no colour profile), at the exact size the platform asks for. When the platform's help page and its own screen disagree on the size, ship the exact one plus a backup at the other size made by scaling and centring the same design, not by redesigning it.

## The two designs

- **A. Title block:** a thin green top bar; eyebrow (15 px bold green), title (40 px bold green, two lines, slightly tight letter spacing), subtitle (16 px, two lines); a hairline rule near the bottom and a strip of up to four labels with small green squares, spread evenly across the width. Optional official logo on the left with a 2 px green divider, the text block starting after it.
- **B. Green card and labelled arcs:** a green card at the left (radius 8 px) with the eyebrow, title (36 px) and subtitle in white; three concentric quarter arcs (14 px strokes, flat ends) that run off the bottom edge, each ending in a small square and a bold 16 px label for one thing the session covers. The longest label ends on the right margin. The stronger of the two: the card anchors a plain white page and the arcs say what the session is about.

## Rules

- 30 px margins on every side; shared top lines (a logo's top = the card's top); a label's right edge on the margin.
- Anything that runs to the edge must run past it; a few pixels short reads as a mistake.
- Only colours from the palette; text on green is pure white (not a pale tint).
- No decoration without meaning: every arc or square carries a label; drop anything the size cannot label.
- Logos unaltered, only on white, never below the size at which their fine detail (dots, small letters) still reads; check at 3x zoom against a clean resize of the original.
- On a phone only the title survives: the title carries the message; details live in the platform's own text. Leave dates off a banner used for several sessions.
- Same words as the invitation; nothing the source does not say.
