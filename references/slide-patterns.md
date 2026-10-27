# Slide patterns

Pick the device that shows the content; never fall back to a title and bullets. Green below means the deck's primary colour: MongoDB's Forest Green #00684A by default, navy #243168 with `Deck(theme="navy")`. Every pattern below is a function in `assets/slides/slide_patterns.py` (or a few kit calls) and appears in the demo deck (`assets/slides/demo_deck.py`, renders in `assets/slides/examples/`).

## 1. Choose by the shape of the content

| The content is... | Show it as | Function |
|---|---|---|
| a request or proposal (one ask) | the ask as the title question, the real artefact (calendar, wording on a phone), what you would provide, a thank-you | `calendar`, `phone`, `icon_row`, `run_in`; variants below |
| time windows, phases with dates | a calendar strip: months on an axis, each window a coloured bar on it, named above | `calendar` |
| exact wording people would see (a message, an email, a notice) | the wording inside its real container: a phone with notification cards, a card per message | `phone`, `notification` |
| what you offer or need (two to four items) | icon plus a short balanced label, in a row or a column | `icon_row`, `icon_column` |
| stages, done / now / next | one chevron bar, solid for done, lighter for now, dashed for not yet; what belongs to each stage underneath | `chevron_stages`, `card`, `gutter_dividers` |
| three headline numbers | a stats row: icon tile, small qualifier above ("Nearly"), the number big, a label under, thin rules between | `stats_row` |
| a problem and a purpose | a pale "need" card, an arrow down, a green "aim" card | `need_and_aim` |
| three qualities, three ways in, three audiences | three labelled concentric arcs | `labelled_arcs` |
| who decides what (levels and members) | two zones: levels down the left with the role inside each box, one level's members in a panel joined to it | `two_zone_structure` |
| equal peers, each with a status | an equal grid of white cards (live) and dashed cards (coming), all the same size | `card` in a grid |
| a sequence of real steps | numbered steps in order (numbers only for real sequences) | kit boxes and arrows |
| the opening | a wide green panel with the title, subtitle and who/when, the talk's parts as labelled arcs | `title_slide` |

## 2. The request slide, four ways

The same content (title question, one or two lead sentences, the real artefact, what you would provide, an optional second ask, a thank-you) in four layouts. Offer them side by side when the direction is open; the user often mixes them.

- **A. Band, calendar and phone:** a thin green band across the top; title and lead; the calendar under the title on the left; the phone rising from the bottom edge on the right; offers in a row; a dotted divider; the second ask as one run-in sentence; the thank-you at the bottom left. Reads top to bottom.
- **B. Two columns and a closing band:** title and lead; two mirrored columns (each window's dates in its colour, who it is for, then its notification card, both cards the same height); a dotted divider in the gutter; the second ask; a full-width green band at the bottom with the offers in white and a right-aligned thank-you. Best for a one-minute guest slide.
- **C. Green side panel and three parts:** a full-height green panel at the left holds the question (large white type, line breaks set by hand), who is asking and a quiet arc motif; the white side is a brief in three parts, each a square-bullet heading, separated by dotted hairlines.
- **D. Panel plus pictures:** C's panel carries all the words; the white side carries the pictures: the calendar across the top, the phone at the bottom right, the offers as a column beside the phone (its heading level with the phone's label). Reads left to right, from the question to the evidence.

What makes them work: the audience sees the thing it is being asked to approve; the dates are a picture read in one glance; invisible precision (the first window starts exactly under the title's left edge; the phone top equals the window names' top); a structural cue borrowed from the host deck (its top band) so a guest slide looks native.

## 3. Pattern details

**Calendar strip** (`calendar`). Months evenly spaced on a 1.25 pt green axis, ticks of 1 pt on the month boundaries, month labels 12 pt green centred between ticks, the year 12 pt bold at the far left on the label row. Each window is a bar 0.46 in tall sitting on the axis, filled with its colour, its dates in 14 pt bold white centred; its name sits above it in 13 pt bold in the same colour, or in the primary when that colour is too light for small text (the kit's `readable()` decides; write names as text; never rely on a logo's small wordmark to name something). Start the axis one month-width left of the title edge when you want the first window to line up under the title.

**Phone** (`phone`, `notification`). Frame: rounded top corners only (0.42 in radius), 3.6 in wide (narrower forces the wording below 12 pt), running exactly to the bottom edge, so it reads as a phone held up into view. Screen inset 0.1 in in pale green; a small green pill for the camera; a 12 pt bold green label ("Proposed wording"). Cards: white, 1.25 pt outline in the window's colour, radius 0.12 in; a green app tile with a white bell; app name bold; a right-aligned label (the month, or "Proposed wording"); then the exact wording at 12 pt or more. The wording is quoted exactly; never add non-breaking spaces or edits inside it.

**Offers** (`icon_row`, `icon_column`). Each item is a line icon (0.36 to 0.42 in) and a label balanced on two lines; in a row, each label takes the narrowest width that balances, and the spare width is shared so the gaps look even. One icon size per block, so text edges line up.

**Second ask** (`run_in`). One sentence with a bold green lead-in ("Next step:") and an icon. Keep it to two lines.

**Closing band** (`closing_band`). Full width, green, from about 6.1 in to the bottom edge: a 13 pt bold white heading, the offers in white, and the thank-you right-aligned in 16 pt bold white on two lines.

**Green side panel** (`side_panel`). 4.4 to 4.75 in wide (about a third of the slide), full height. Title 30 pt bold white, line spacing 95%, lines chosen by hand so a name stays whole; subtitle 15 pt; lead 14 pt; text ends above 5.5 in, where the motif begins. The motif (three concentric arcs in a slightly lighter green, 16 pt strokes, clipped by the slide edge) is an SVG, so nothing sits off the slide.

**Stats row** (`stats_row`). Three items across the content width: an 0.82 in pale green icon tile, then a column with the qualifier (14 pt bold green: "Since", "Nearly", "Over"), the number (32 pt bold green; 44 pt when the numbers are the slide) and a label (13 pt). Short solid hairlines between items. Use only numbers from the source, with the source's own qualifier.

**Need and aim** (`need_and_aim`). The need in a pale green card (icon, 13 pt bold label, 15 pt text with one bold key phrase), a short green arrow down from the icon's centre line, the aim in a green card (white icon, label, 17 pt bold white statement). Put the labelled arcs or a picture beside it; plain paragraphs beside a strong visual look unfinished.

**Labelled arcs** (`labelled_arcs`). Three concentric quarter arcs (upper-left quarter, flat caps, green, 13 to 15 pt strokes, radii in steps of about 22% of the outer radius), each ending in a small green square and an 18 to 20 pt bold green label. Native arcs, so the user can edit them: Insert > Shapes > Arc, Flip Horizontal, a thick green outline with flat caps, sizes aligned centre and middle.

**Chevron stages** (`chevron_stages`, `gutter_dividers`). One bar 0.68 in tall: the first segment a pentagon, the rest chevrons (tip depth 0.32 of the height), each with a 0.38 in icon and a 16 pt bold heading inside, inset equally clear of the notch. Done = green, now = lighter green, next = white with a 1.5 pt dashed green outline. A 12 pt note under each segment. Below, each stage's items as cards in its column (a single card centred between taller stacks), and dotted rules down the gutters that end where the cards end. The last tip stays inside the right margin. This is the device for "clearly divided" parts; no number boxes, the shape already carries the order.

**Cards** (`card`). One radius for the whole deck (0.08 in). Live: white with a 1.25 pt hairline border. Not yet: white with a 1.25 pt dashed green border. Quiet: pale green fill, no border. A card must stand for something real (a team, a product, a message); never a grid of identical decorative cards each with an icon on top.

**Two-zone structure** (`two_zone_structure`). Levels down the left as green boxes (name 14 pt bold white, role 12 pt white inside the box, so no leader lines), joined by a 3 pt green spine; the level that has members is joined by a 3 pt line to a panel (3 pt green outline, radius 0.2 in) holding the members as white cards with a green band and a 12 pt bold green name, in rows of four, a short last row centred. Line weights match where they meet (3 pt line into a 3 pt border). Give each member its own band colour only when that colour means something (a brand or category).

**Title slide** (`title_slide`). A wide green panel (about 6.4 in) with the title in 40 pt bold white on two or three hand-set lines, the subtitle 18 pt, who and when 14 pt, the arc motif at the bottom; on the white side, the talk's three parts as labelled arcs. End the talk on the slide people need to act on (the ask, a sign-up code), not on "Thank you".

**Call to action or sign-up code.** The code large (about 40% of the slide height), dark on white, no logo inside; a short web address under it; say what people get and how long it takes; if there are two codes, give each a distinct label. Leave it up through the questions.

## 4. Technical content

The same idea, show the real thing, for the content technical decks carry most:

- **Architecture or data flow:** components as boxes in reading order (left to right, or top to bottom), white with a hairline border, the one that matters filled green; 3 pt green arrows between them, each labelled with what flows ("change events", "nightly export"); group boxes inside a thin container when they run together (a cluster, a region). One small icon per component at most; no clip-art servers or clouds. Build it with `d.box`, `d.line(..., arrow_end=True)` and `d.one` labels.
- **A document, a query or code:** a "screen" panel (pale surface fill, radius 0.12 in) holding only the lines that matter (5 to 12), in a monospace font (Consolas, or Menlo on a Mac) at 12 pt or more, indented as in the real thing; mark the key field or line with a pale green band behind it and a short callout to the right ("the index on this field"). Never a screenshot of an editor, and never more code than the audience can read.
- **Numbers by group or over time:** a native PowerPoint chart (python-pptx `add_chart`) so the numbers stay editable: bars or columns sorted by value, green first and the lighter steps after, no 3D, no legend when the bars are labelled, value labels on the bars when there are few, axis text 12 pt, gridlines off or in the hairline colour. Put the takeaway in the title ("Queries run 40% faster after the change"), not "Results".
- **A results table:** header row in green with bold white text, hairlines between rows, numbers right-aligned, the row that matters in bold, 12 pt or more. Six rows or fewer on a slide; the rest goes in an appendix or the notes.
- **A comparison** (before and after, option A or B): two mirrored columns with the same rows in the same order, the differences in bold, and a one-line verdict under them; a dotted divider down the gutter.

## 5. Rules from what worked and what did not

- **"Clearly divided"** means one chevron bar for sequential parts, with visible gaps, plus dotted rules down the gutters starting under each gap, and equal heading insets.
- **"No distinction"** came from mixing three card styles. Separation comes from structure (a header bar, gaps, dividers), not from varying card styles: one light card style on both sides of a single dark focal panel, one radius, aligned column tops and bottoms.
- **Equal peers go in an equal grid**, each with its own status; a hub-and-spokes diagram makes positions unequal when the centre adds nothing. Never show the same list twice, and never stretch a picture.
- **Parts that are not steps get small green squares, not numbers;** numbered badges on non-steps read as generic.
- **Make the left side visual:** text-only paragraphs beside a strong visual look unfinished; put the text into a card with an icon, or a flow.
- **Keep the approved wording while iterating the visuals,** so a new version changes one thing.
- **Readable on a projector and in a video call:** text at 12 pt or more (13 pt for explanatory text), hairlines no paler than the palette's own (#A6CAC0 in green, #BCC3DE in navy), card outlines on pale backgrounds.
- **Align exactly or offset clearly;** a near miss of a few hundredths of an inch reads as an error.
- **Mirror column grids:** the same text offset in both columns, whatever the logo widths.
- **Run devices to the slide edge on purpose** (the phone, the band, the panel) or keep them clearly inside; never nearly.
