# Building, checking and finishing

## 1. Workflow

1. **Read the source** and keep it untouched (work on copies; never rebuild over a deck someone has edited by hand without asking).
2. **Write the words first:** for each slide, the title (the message, often a question or a claim) and at most two short lead sentences, one per line. Use the source's own words and figures; anything not in the source is left out or marked "to be confirmed".
3. **Pick a device per slide** (`slide-patterns.md`).
4. **Sample first** for a new deck: build two representative slides (one simple, one complex), render them, show them, then build the rest.
5. **Build with the kit** (section 2).
6. **Render every slide and look at each image** (section 5). Fix, render again.
7. **Independent check:** a fresh reviewer (or a separate agent) looks only at the renders with the checklist in section 6, and reports must-fix and polish items with exact numbers ("0.24 in gap, make it 0.30"). Apply, re-render, re-check.
8. **Options when the direction is open:** two to four real, editable versions in one file (slides A, B, C, D), one recommendation, and a simple way to answer (for example a page whose radio buttons compose a line such as `1=B 2=A` to paste back). The user may mix versions; build the mix as a new slide and keep the others unchanged.
9. **Speaker notes** on every slide (section 4).
10. **Finish the file** (section 7) and hand over the file and the renders.

## 2. The kit

`assets/slides/slide_kit.py` (drawing, text, pictures, notes, saving) and `assets/slides/slide_patterns.py` (the devices). Requires `pip install python-pptx pillow pymupdf`; Windows checks need `pywin32` and PowerPoint. Start from `assets/slides/demo_deck.py`.

```python
import sys; sys.path.insert(0, "path/to/navy-house-style/assets/slides")
from slide_kit import Deck, P, L, R, SH, N, B, readable
from slide_patterns import header, calendar, phone, icon_row, run_in

d = Deck()                                   # 16:9, one layout, green theme; Deck(theme="navy") for navy
# colours come from the palette P: P.primary, P.surface, P.hairline, P.now, P.ramp[0..3];
# readable(colour, size, bold) gives the primary instead of a series colour that is too light for small text
s = d.slide()                                # white, with the locked white backing
top = header(d, s, "Could the app help announce both pilots?",
             ["Two pilots open this year.", "A notice in the app could help people join."])
...
d.notes(s, [("Script (about one minute)", True), ("...", False), (), ("Background, if asked", True), ("...", False)])
d.save("deck.pptx", author="Owner's name", title="Deck title")
```

Building blocks: `d.title` (fills the title placeholder; call it right after background shapes so it is read first), `d.one` (one-line label, refuses to overflow), `d.para` (wraps to a width, refuses more than `max_lines`), `d.sentence_lines`, `d.text_lines` (lines set by hand), `d.box` (radius in inches), `d.line` (dotted with `dash=SQUARE_DOT`, `arrow_end=True`), `d.divider` / `d.vdivider`, `d.square`, `d.circle`, `d.arc`, `d.icon` (Phosphor name from `assets/slides/icons`), `d.svg_picture`, `d.logo` (crops to the artwork, never stretches), `d.notes`.

How the kit keeps text right:

- Every text box has zero insets, no autofit, wrap on and an explicit anchor, so a box's x is the visible text edge and alignment is pure arithmetic.
- Text is measured with the real font file; the line breaker chooses the fewest lines, then the most even ones, and never leaves a last line under half the width. Line ends are written as explicit breaks, so PowerPoint never re-wraps them.
- Fonts: the kit finds Tahoma in the usual font folders on Windows, macOS and Linux, or in the folder named by the `SLIDE_FONT_DIR` environment variable (it needs `tahoma.ttf` and `tahomabd.ttf`). Without Tahoma it measures with Verdana or DejaVu Sans and warns, since the line breaks are then less exact; with no font at all it stops with a clear message. `Deck(font="Verdana")` writes and measures every line in Verdana (for Google Slides); `use_font(name, regular, bold)` does the same for any other font, by name or from its two files.
- Non-breaking spaces (`nb("product name")`) hold names, a name and its abbreviation, and pairs like "take part" together. Never put them inside wording quoted from a source.
- A build stops (assertion) when a label does not fit, a paragraph needs more lines than allowed, or a device would cross the bottom margin.

## 3. Guest slides (a slide for someone else's deck)

PowerPoint's default paste (Use Destination Theme) drops a slide's own background and its "hide background graphics" setting, and lets the host deck's title style in (for example ALL CAPS). The kit already guards against this:

- a locked, full-slide white rectangle at the very back (locked, but never "no select": the host team must be able to delete it);
- every run pinned: font, size, bold, italic off, colour, language, no capitals, no letter spacing, no effects;
- every line pinned (cap, compound, alignment, dash, join) and every shape and connector given an empty effect list, so a host theme's shadows cannot appear;
- the title placeholder given no fill, no line and no effects.

Test it: paste the slide into a copy of the real host deck and into a deck with a deliberately clashing master (coloured background, a different title font in capitals), export both, and compare with the original pixel by pixel. Borrow one structural cue from the host deck (for example its top band) so the guest slide looks at home. Untested risks to mention: footers and slide numbers applied to all slides after pasting, 4:3 hosts, "Keep Source Formatting" pastes, PowerPoint for the web.

## 4. Speaker notes

- A bold **"Script (about one minute)"** section: about 150 to 180 words in short paragraphs: who is asking (names spelled out at first mention), the facts, the request as a question and exactly what would be provided, any second ask, then the thank-you.
- An empty line, then a bold **"Background, if asked"** section: terms explained, the verbatim wording with its source and date, what any placeholder means, details that would crowd the slide, and claims tied to their dated source ("The brief, dated 3 March, says ...").
- Label both parts, so a guest presenter never reads the background aloud. Match the source's modality ("would provide", not "will provide"); never invent a contact or a commitment.

## 5. Rendering and checking

- `python assets/slides/render_slides.py deck.pptx render_folder` exports every slide at 1600 px wide (PowerPoint on Windows; LibreOffice elsewhere). On Windows it also asks PowerPoint how many lines each text box really lays out and compares that with the lines the build set: any difference, or text taller than its box, is reported.
- Look at every image at full size, then at phone or thumbnail size (does the title still read?).
- Without PowerPoint or LibreOffice, ask the user to export the slides (File > Export > PNG) and look at those.

## 6. Checklist for the independent review (on the renders)

- One message per slide, in the title; the device shows the content (no title-and-bullets slide, no decorative card grid).
- Nothing invented: every name, date, figure and quotation is in the source; quoted wording is exact.
- No overflow, no stranded single word on a last line, no split name, no text under 12 pt (13 pt for explanatory text).
- Exact alignment: shared left edges, tops and baselines; equal gaps between peers; even gaps above and below dividers; equal margins.
- Colour carries meaning only; category colours pass contrast for text or sit only on marks; no grey; no gradients, shadows or 3D.
- One radius, one icon family and one icon size per block; line weights match where they meet.
- Devices touch the slide edge on purpose or stay clearly inside.
- Readable from the back of a room and in a video call window.

## 7. Finishing the file

- `d.save(path, author=..., title=...)` sets the author and dates. Keep files free of tool names: no "generated by" text in properties, notes or file names.
- Optional on Windows: open and save once in PowerPoint for a clean `app.xml`; that save drops the PNG fallbacks of SVG pictures, so run `slide_kit.finish(path, "render/slide-01.png", author)` afterwards, which restores them and sets the thumbnail from a render of slide 1.
- Strip dead linked pictures inherited from pasted content (a "Security information / active content" bar in PowerPoint): remove `r:link` attributes whose relationship is external, then confirm no external relationships remain.
- Name shapes meaningfully (the Selection Pane shows them), give pictures alt text, and keep reading order = creation order: background, title, then content in the order it is read.

## 8. Traps

- `add_connector` takes the theme's line effect: on some company templates that is a drop shadow, so every thin rule looks blurred. The kit clears it (`shadow.inherit = False`); do the same on anything drawn outside the kit.
- Two effect lists in one shape make PowerPoint refuse the file.
- A title placeholder comes with every new slide and sits under shapes drawn later: the kit moves it to the right place when you call `d.title`.
- `run.text = "a\vb"` writes a literal `_x000B_`; use the kit's explicit breaks.
- Hyperlinks take the theme's link colour unless told otherwise.
- PowerPoint fully enforces "no select" on a shape (even scripts cannot touch it): never use it on shapes others may need to delete.
- A slide added with python-pptx does not copy a slide-level background.
- Writing scripts through shell here-documents can strip backslashes: write builder files with a file tool, and patch them with replacements that must match exactly once.

## 9. Google Slides

Build in PowerPoint and convert at the end: upload the .pptx (or File > Import slides), then look at every slide again. Tahoma is not on Google's font lists, so set Verdana (Tahoma's wider sibling, about 14% wider) and re-check every line break. Also check, because the conversion was not tested here: icons (they may arrive as plain pictures), dashed and dotted lines and arrowheads, rounded corners, and the speaker notes.
