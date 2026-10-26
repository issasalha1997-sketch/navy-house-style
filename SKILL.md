---
name: navy-house-style
description: House style for documents, spreadsheets, figures and HTML pages - navy 243168 as the one colour, Tahoma, calm spacing, plain British English, no grey. Use it by default whenever you create or restyle a Word or Google Docs document, an Excel workbook or Google Sheet (dashboards, trackers, registers, checklists, buttons, Apps Script dialogs), a chart or figure, or a local HTML explainer, guide, walkthrough or pack, unless the user asks for another style or the work must match an existing template. Also use it when the user says "navy style" or "house style", or asks for the MongoDB logo on a deliverable (the official logo files are included).
---

# Navy house style

One calm, readable look for everything: documents, spreadsheets (Excel and Google Sheets), figures and HTML explainer pages. Navy is the only brand colour. Near-black text, white space and a few colours that carry a meaning do the rest.

## The five ideas

1. **One colour carries the structure.** Navy #243168 marks titles, headings, header rows, captions, rules and buttons. Body text stays black or near-black.
2. **Calm before clever.** One typeface, generous spacing, left-aligned text, hairlines instead of boxes, nothing that decorates without a meaning.
3. **The reader's story sets the order.** Start with what the reader needs first (where things stand, what to do now), then the detail.
4. **Plain English.** Short sentences, everyday words, British spelling.
5. **Nothing people type into can break.** Data stays plain; counts and checks live on a read-only dashboard; scripts and buttons only help.

## Pick the guide for the job

| Making | Read | Starter |
|---|---|---|
| A document (Word, Google Docs, a PDF from either) | `references/documents.md` | python-docx recipe inside |
| An Excel workbook | `references/spreadsheets.md` | `assets/excel/navy_workbook.py` (run it for a demo) |
| A Google Sheet | `references/spreadsheets.md`, then `references/google-sheets.md` | `assets/google-sheets/` |
| An HTML explainer, guide, walkthrough or pack | `references/html-explainers.md` | `assets/html/explainer-template.html` |
| A chart, diagram or figure | `references/figures.md` | |

Read only the files the task needs.

## Colour

| Role | Value | Where |
|---|---|---|
| Navy | #243168 (RGB 36, 49, 104) | titles, headings, captions, table header rows (bold white text), first-level bullets, rules, buttons, the lead chart colour |
| Text | black in documents; ink #1B1F2A in sheets; #1D1D1B on HTML pages | all body and data text |
| Hairline | #D3D8EA | rules and lines under rows (box and field edges on sheets and screens use #C9CEDF) |
| Secondary text | soft #2C3A78 in sheets; #4A4F63 on pages | hints, notes, captions |
| Links | #0563C1 | so links never look like headings |
| Status (font colour) | done #1E7B3C, waiting #8A5300, with others #00618A, stopped #A4262C | lists, registers, dashboards |
| Attention | #FFE699 needs a look, #FFF4CC or #FFF3B0 still to fill in | only on the exact cell or words concerned |

Every role, tint and contrast ratio: `assets/palette.json`. A few values differ slightly between documents, sheets and pages because each comes from its own proven design; use the values given for the medium you are working in.

- **No grey anywhere:** no grey text, fills, borders or chart series. Every neutral is a pale or deep tint of navy.
- Colour only where it carries a meaning, and never as the only signal: add a word, a mark or a pattern.
- A tint of navy is a design treatment, not a new brand colour. Do not invent brand colours.
- The navy stays even when the work is about a brand with other colours (MongoDB's greens included). A logo can name that brand; the style does not change. The exception: official material that must follow MongoDB's own brand, or a template MongoDB supplies. There, MongoDB's brand and template win.

## Type

Tahoma everywhere; fallbacks Verdana, then sans-serif.

- **Documents:** body 12 pt; Heading 1 14 pt bold, Heading 2 12 pt bold, Heading 3 12 pt regular, all navy; tables at least 10 pt; footnotes 10 pt.
- **Spreadsheets:** body and data 10 pt; header row 10 pt bold white on navy; titles 14 to 20 pt bold navy; notes 8.5 to 9 pt italic.
- **HTML pages:** body 15 px with line height 1.55; h1 28 px, h2 21 px, h3 16 px; small text 13 px.
- **Google Docs and Sheets:** see `references/google-sheets.md` for the font that replaces Tahoma there.
- Headings: Heading 1 in documents in Title Case, everything else in sentence case. No ALL CAPS, no letter-spaced capitals.

## Writing (every medium)

- Plain English for a reader who is not a specialist; explain a technical term the first time it appears.
- British spelling by default (organise, colour, programme, centre); US spelling only if the user asks for it.
- Formal documents are impersonal (avoid "we" and "our" where possible); guides and explainers may speak to the reader ("Open the sheet, then ...").
- No em dashes in sentences: use a comma, a colon, a full stop or a new sentence. En dashes in ranges are fine.
- Lead with the point. Summaries are short prose under plain headings, not bullet lists or questions and answers.
- Present the final state; never narrate what was found and fixed along the way.
- Never invent facts, numbers, names, dates or approvals. Label proposals as proposals and mark anything unconfirmed "to be confirmed" (navy italics in sheets and pages).
- Dates in prose in full ("Monday 2 March 2026"); in sheets `d mmm yyyy` ("2 Mar 2026").
- Names of things (sheets, dashboards, pages, buttons) are plain and informative, never metaphors.
- **MongoDB's name** (its trademark rules): never joined to another word or hyphenated ("based on MongoDB", not "MongoDB-based"), never abbreviated ("Mongo"); in external material add ® where the mark is registered (MongoDB's trademark page lists which are).

## Never

- Grey, or pure black text on sheets and pages (use the ink).
- The generic "AI design" look: tinted cards, rounded pills and badges, chips, letter-spaced uppercase labels, decorative legends of coloured squares, accent bars down the side of boxes (one exception: the amber notice banner at the top of a workbook), gradients, drop shadows, emoji used as icons.
- Justified text, text boxes in Word, ALL CAPS.
- A logo the user did not ask for.

## Logos

No logo by default: the navy carries the style. When the user asks for a logo:

- Use only official files, unaltered: never recolour, stretch, crop, outline or add effects. Show it once per page or document, small, in a prominent but quiet place (top right of a page header, a cover, a title band), with clear space round it.
- **MongoDB** (official press-kit files in `assets/logos/mongodb/`, see the README there):
  - on white, the default for this style: `mongodb-logo-slate` (contrast 17.2:1); for a touch of the brand's green, `mongodb-logo-evergreen` (13.7:1);
  - on a navy band: `mongodb-logo-off-white` (12.1:1);
  - the press kit's Spring Green and Chartreuse versions are left out on purpose: 1.6:1 and 1.2:1 on white, and loud on navy;
  - the leaf alone only where the full logo is already shown or the audience clearly knows it (a favicon, a footer); neither the logo nor the leaf ever becomes part of an illustration;
  - in material not made by MongoDB, the logo only identifies MongoDB's product or service: it is never the document's own brand, nothing may suggest that MongoDB sponsors or endorses the work, and the author's own name or logo stays more prominent than MongoDB's;
  - placing it: an HTML page header about 24 px tall, inlined (see `references/html-explainers.md`); Word, the PNG with `doc.add_picture(path, width=Cm(4))` (Word does not take the SVG); Excel, the PNG through `openpyxl.drawing.image.Image` sized to about 140 x 36 px, top right of the dashboard; Google Sheets, Insert > Image > Image over cells at about 140 x 36 px;
  - rules and questions: https://www.mongodb.com/legal/trademark-usage-guidelines and trademark@mongodb.com.
- Any other organisation: ask the user for the official file or take it from that organisation's brand page; the same rules apply.

## Before handing anything over

- One navy, used the same way everywhere; no stray default colours (Office theme blues, default link blue on headings, default chart palettes).
- Tahoma (or the agreed substitute in Google) everywhere, including headers, footers, charts, buttons and dialogs.
- No grey, no em dash in sentences, no ALL CAPS, no invented colours.
- Look at the result as the reader will: render it, view a picture of it, click every button and link, try it at phone width for pages.
- Everything the reader must open is a link they can click.
