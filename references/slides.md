# Slides, decks and banners

Slides people remember because the content itself is the picture: the real dates on a calendar, the exact wording on a phone, the stages as one chevron bar, the numbers big. Everything native and editable, calm, green. Look at `assets/slides/examples/` before building: that is the target. Slides default to green; the rest of the house style (documents, spreadsheets, HTML pages) stays navy.

## Generic AI slides, and what to do instead

| Generic | Instead |
|---|---|
| A title over four bullets | The message as the title (a question or a claim) and the content drawn: calendar, phone, chevrons, numbers |
| A row of identical tinted cards, each with an icon on top | A device that shows the relationship: stages as chevrons, peers as an equal grid, a need flowing into an aim, a level joined to its members |
| Icons as decoration | One icon per real item, from one line-icon family, one size per block, beside the words it explains |
| Stock photos, gradients, shadows, 3D | Flat green, its pale tints and white space; other colours only where they mean something |
| Numbered circles for things that are not steps | Small green squares for parts; numbers only for a real sequence |
| Everything centred | Left-aligned text on a grid with exact shared edges |
| "Overview", "Agenda", "Key takeaways" | A title that says the point |
| Ending on "Thank you" | Ending on what people should do (the ask, a sign-up code) |

## The rules

1. **One message per slide**, written as the title; one or two short lead sentences under it, one per line.
2. **Show the real thing.** The audience should see what it is being asked about: the actual dates, the actual wording in its real container, the actual numbers with their own qualifiers. Use the source's words; never invent a figure, date, name or quotation.
3. **Make the structure visible.** Chevrons for stages (solid done, lighter now, dashed not yet), dotted dividers between parts, a green panel or band to separate the ask from the detail, containers for membership.
4. **One motif per deck**, used sparingly: three concentric arcs (quiet on a green panel, or labelled as a small diagram). Never build a motif from someone else's logo (MongoDB's leaf may never be part of an illustration).
5. **Type does the work.** Big bold green titles, a calm lead, bold run-in labels ("Next step:"), big numbers; nothing under 12 pt.
6. **Restraint.** White space, green and its tints; extra colours only to tell real categories apart, each with one meaning.
7. **Measured, rendered, looked at.** Build with the kit (text is measured with the real font and line breaks are set), render every slide, look at every image, then have a fresh reviewer check the renders.

## Workflow

1. Write each slide's title and lead from the source first.
2. Pick the device that shows the content: `references/slide-patterns.md`.
3. Build with `assets/slides/slide_kit.py` and `assets/slides/slide_patterns.py` (python-pptx); start from `assets/slides/demo_deck.py`.
4. Render and check: `python assets/slides/render_slides.py deck.pptx render` (on Windows it also checks PowerPoint's own line layout). Look at every slide.
5. Independent review of the renders with the checklist in `references/slide-build.md`; fix and re-render.
6. When the direction is open, give two to four versions in one file (A, B, C, D) and one recommendation.
7. Speaker notes on every slide: "Script (about one minute)", then "Background, if asked".
8. A slide for someone else's deck: the kit makes it paste-proof; test it as `references/slide-build.md` describes.

## Look

- **Canvas:** 13.333 x 7.5 in (16:9); margins 0.7 in left and right; content from about 2.2 in (under the title and lead) to about 7.0 in.
- **Type:** Tahoma. For a deck going to Google Slides, build with `Deck(font="Verdana")`, so the text is written and measured in Verdana. Title 28 to 32 pt bold in the primary colour on one line (30 to 40 pt in a panel or on the title slide); lead 15 pt; headings 14 to 16 pt bold; body 13 pt; labels and notes 12 pt; headline numbers 32 to 44 pt bold.
- **Colour (green, the default):** primary #00684A (MongoDB's Forest Green; white text on it 6.8:1); pale surface #E3FCF7 (MongoDB's Mist) for phone screens, quiet panels and icon tiles; dotted dividers #A6CAC0; the motif on green and the "now" stage #2E836B; further series #4C9580 and #8CBBAE; body text black; no grey. A series colour too light for small text gives way to the primary automatically (`readable()` in the kit) and stays on bars and outlines.
- **When green, when navy:** green is the default because the work is about MongoDB, and its greens sit well with that. Use navy when the material is not about MongoDB, or when MongoDB's colours next to its logo could suggest that MongoDB made or endorsed the deck. Official MongoDB material uses MongoDB's own template, colours and fonts instead of this style.
- **Colour (navy, the option):** `Deck(theme="navy")` and `THEME = "navy"` in the banner builder switch everything to navy #243168 with its tints (#EEF0F7, #BCC3DE, #3A4A86, #5B6AA6, #8F99C2), the same palette as the rest of this skill.
- **Shapes:** one corner radius per deck (0.08 to 0.12 in); 3 pt lines for structure, 1.25 pt card outlines, 1 pt dotted hairlines; solid = done or live, dashed = not yet.
- **Icons:** `assets/slides/icons` (Phosphor, regular weight, MIT licence), 0.36 to 0.42 in beside text, in the primary on white and white on the primary, inserted so PowerPoint can recolour them.

## Logos

No logo by default. When asked, use only the official file, unaltered, once per slide, small, on white (or the off-white version on a green or navy panel). MongoDB: `assets/logos/mongodb/` (Slate on white is the default; Off-white on a coloured panel; see the README there). In material not made by MongoDB the logo only identifies MongoDB's product, never the deck's own brand, and the author's own name stays more prominent. If the deck must follow a template the user supplies (for example a company template), build on its layouts, fonts and colours and keep these rules for the content.

## Files

- `references/slide-patterns.md`: which device for which content, with the details and the lessons learned.
- `references/slide-build.md`: workflow, kit, guest slides, notes, rendering, review checklist, finishing, traps, Google Slides.
- `references/banners.md`: event banners and social images (HTML rendered to PNG).
- `assets/slides/slide_kit.py`, `assets/slides/slide_patterns.py`, `assets/slides/demo_deck.py`, `assets/slides/render_slides.py`, `assets/slides/build_banner.py`, `assets/slides/icons/`, `assets/logos/mongodb/`, `assets/slides/examples/` (renders of the demo deck in green, the request slide in navy, and two banners).
