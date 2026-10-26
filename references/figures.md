# Figures and charts

For any chart, diagram or picture that goes into a document or an HTML page.

## 1. What a good figure is here

- Static, sharp and accurate. Export SVG plus a PNG at about 300 dpi, sized for the page: about 16 cm wide for a portrait page, 22.5 cm for a landscape one.
- As little text as possible. A short title at the top and straight into the picture. Any "how to read this" line goes at the bottom, in plain English, followed by a key for every colour used.
- Every case (each row, person, item or group) is visibly distinct.
- One figure, one message. Never pack several small visualisations into one picture.
- When asked for "less text", cut words, not cases. Each new version gets simpler, not denser.
- Offer one clear recommendation (which figure, and why) rather than more styles to choose from.
- Test it before handing it over: someone who sees only the picture should be able to answer the questions it is meant to answer.

## 2. Start from the shape of the data

- **Who and what** (people or roles against what they can see or do): a grid. Rows are who, columns are what. Where the answer is yes, a coloured box with one or two words; blank only ever means "no". If the source does not say, use a separate visible mark for "not stated", never a blank.
- **Structure that must be explained** (how one case works): small panels, one per case, each with a name, the picture and one short line. Explain the picture once, in a key at the top, not in every panel. No paragraphs inside the figure.
- **Compact overview** of many cases: a simple heat map, offered as the alternative to the grid.
- Source notes go in a small footnote under the figure.

## 3. The look: a report table, not an app

- Square cells. Header words sit on a coloured rule: rules, not containers.
- One colour (navy) for all header words; colour lives only in the rules and the cells.
- A fine line between rows. Every separator of the same kind looks identical: pin lines to whole pixels and keep heights even.
- A plain tick, not a heavy checkbox mark. No bold words in solid blocks (they read as buttons).
- Nothing in the first column spills into the next one: break long headings onto two lines.
- Clear division between groups of rows (a slightly heavier rule or a gap), and a group name never repeats a column name.
- The key is one clean row: black words beside small boxes of the exact colours and marks used. The key's marks match the figure exactly, one symbol means one thing, and a colour in the key is not reused for anything else (for example header text).

## 4. Colour

- Navy #243168 leads: the first series, the main marks, the header words.
- More series: lighter tints of navy first (see the palette in `assets/palette.json`). When categories need different hues, use the status colours only where they carry their status meaning, and keep every colour's meaning the same across the whole document.
- No grey: no grey series, gridlines, fills or muted labels. Use black, navy, a navy tint or white space.
- Pale colours are fills under black text, never text themselves.
- Never rely on colour alone: add labels, shapes or patterns so the figure still works printed in black and white.

## 5. Building it

- Draw the figure as SVG (by hand, from Python, or with a charting library), using the real font so text measures correctly, then export PNG at about 300 dpi.
- matplotlib: `rcParams["font.family"] = "Tahoma"` (fallbacks `Verdana`, `DejaVu Sans`), navy as the first colour, no 3D, no shadows, no gradient fills, thin black axis lines, light navy tint gridlines only when they help reading.
- Write numbers on the bars or cells when there are few of them, instead of making the reader use an axis.
- For a figure that will sit in an HTML page, inline SVG keeps it sharp at any zoom; give it a `<title>` and a caption.
