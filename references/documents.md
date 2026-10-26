# Documents (Word and Google Docs)

Reports, briefs, proposals, policies, one-page summaries and any other formal document.

## 1. Page

- A4 (21 x 29.7 cm). Margins 2.54 cm on all sides. A single section may have slightly narrower margins to fit a wide table or figure (use section breaks around it).
- Left aligned. Never justified.
- Paragraphs: 0 pt before, 10 pt after, line spacing 1.15 (Word: Multiple, 1.15).
- One space after full stops and other punctuation.
- No text boxes: they break the flow, move around and cannot take comments. If something must sit in a box, use a one-cell table in the flow of text.

## 2. Type

One typeface throughout: Tahoma (Google Docs: see section 9).

| Element | Size | Colour | Weight | Case |
|---|---|---|---|---|
| Cover title | up to 30 pt | navy | | |
| Cover subtitle | 16 to 24 pt | black or navy | regular | |
| Heading 1 | 14 pt | navy | bold | Title Case |
| Heading 2 | 12 pt | navy | bold | Sentence case |
| Heading 3 | 12 pt | navy | regular | Sentence case |
| Body text | 12 pt | black | regular | Sentence case |
| Table text | 10 pt minimum | black | regular | |
| Footnotes | 10 pt | black | regular | |
| Page numbers | 11 pt | black | numbers bold | "Page **3** of **12**" |
| Running header | 9 or 10 pt | black | regular | |

Use one style per heading level and keep it identical everywhere. If sections are numbered, number every one the same way.

## 3. Colour

Navy #243168 goes in exactly these places:

- the title and every heading level
- table and figure captions (bold)
- the table header row: navy fill, bold white text
- first-level bullet markers (square)
- appendix titles
- the lead colour in charts and figures

Everything else stays black on white: body text, the table of contents, second-level bullet markers, table body text.

No grey in new styling: no grey text, fills, borders, gridlines or chart series. Use white space, black, navy or a navy tint instead. Check the document still reads when printed in black and white: never let colour be the only signal.

## 4. Lists

- First level: square bullet, navy.
- Second level: introduce it with a colon; the marker is an indented black dash (the em dash glyph, used only as a list marker, never as punctuation in sentences).
- A list of short phrases: introduce it with a colon, start each item in lower case (unless it is a proper noun), no punctuation after the items, and a full stop after the last item.
- A list of full sentences: punctuate each item as a normal sentence.

## 5. Tables and figures

- Introduce every table and figure in the text before it. When referring to a specific one, capitalise it: Table 2, Figure 1.
- Caption above the table or figure, bold navy, no full stop at the end: `Table 1. Title of the table`, `Figure 1. Title of the figure`. Wrapped caption lines line up with the start of the title text, not with "Table 1.". In Word use Insert Caption so the numbers update; a script that writes captions as text must renumber them itself.
- Notes, keys and sources go below the table or figure.
- Header row: navy fill, bold white text. Table text at least 10 pt. A thin navy or black border, the same on every table.
- Never split a row across two pages; set the header row to repeat on each page.
- Figures follow `figures.md`.

## 6. Quotes, numbers and names

- Double quotation marks for what a person said; single quotation marks when quoting a document (double inside single for a quote within a quote). Indent a long quote instead of using quotation marks. Never italicise quotes.
- Numbers: one to nine in words, 10 and above in digits. Spell out a number that starts a sentence. Always digits for dates, years, times, money, percentages, and when comparing numbers where one is 10 or more.
- Names: spell a name out on first use with its short form in brackets, for example "service level agreement (SLA)", then use the short form. An organisation is "it", never "they".
- No ALL CAPS for emphasis or headings. Real acronyms (API, SQL, ISO) are fine.

## 7. References and footnotes

- In-text: Vancouver style, superscript numbers in round brackets placed straight after the punctuation, no space: `...was released in 2025.⁽¹⁾` or `...in two studies.⁽²,³⁾`
- Reference list: numbered in order of first appearance, after the main body and before the appendices.
- Footnotes use symbols, not numbers: *, +, ⁑. Placed straight after the punctuation, 10 pt, continuous through the document.
- Reference every fact that comes from a source, including glossary definitions and footnotes.
- A "no references" request applies to that document only, not to the style.

## 8. Document order

Formal report:

1. Cover
2. About (the team or organisation that wrote it)
3. Contents
4. Main body
5. Glossary
6. References
7. Appendices
8. Back cover

Short internal brief: title band, body, sources.

- Contents: black text, built from the heading styles.
- Glossary: alphabetical, only terms that appear in the document.
- Appendices: lettered, `Appendix A. Title`, bold navy, wrapped lines aligned under the first word of the title. Refer to every appendix somewhere in the main body.
- Running header (optional, internal reports): the document title on the top line, the team's name underneath, both 9 or 10 pt, divided by a solid line.
- Footer: "Page X of Y" centred, 11 pt, the numbers bold. The cover counts in the total but shows no number.

## 9. Google Docs

- Font: Verdana, Tahoma's wider sibling (same x-height, about 14% wider), which Google Docs offers in its font menu; Tahoma is not reliably offered there. Same sizes as above.

- Set Normal text, Title, Subtitle and Headings 1 to 3 once (format a line, then Format > Paragraph styles > Heading 1 > Update "Heading 1" to match), then Format > Paragraph styles > Options > Save as my default styles.
- File > Page setup: A4, margins 2.54 cm on all sides.
- Line spacing: Format > Line and paragraph spacing > Custom: 1.15, 0 before, 10 after.
- Table header row: select it, set the cell background to navy #243168 (custom colour) and the text to bold white. Table properties > Pin header rows so the header repeats on each page.
- Captions: Google Docs has no caption numbering, so type `Table 1. Title` above the table, bold navy, and check the numbers before sharing.
- Page numbers: Insert > Page numbers, then type "Page ", the number, " of " and Insert > Page count.

## 10. python-docx recipe

```python
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

NAVY = RGBColor(0x24, 0x31, 0x68)
NAVY_HEX = "243168"
FONT = "Tahoma"

doc = Document()
s = doc.sections[0]
s.page_width, s.page_height = Cm(21.0), Cm(29.7)
s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Cm(2.54)

def set_font(style, size, bold=False, colour=None):
    style.font.name = FONT                  # creates w:rFonts with ascii and hAnsi
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = False
    fonts = style.element.rPr.rFonts
    for attr in ("w:cs", "w:eastAsia"):     # explicit, so no theme font sneaks in
        fonts.set(qn(attr), FONT)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        fonts.attrib.pop(qn(attr), None)
    if colour is not None:
        style.font.color.rgb = colour

normal = doc.styles["Normal"]
set_font(normal, 12, colour=RGBColor(0, 0, 0))
pf = normal.paragraph_format
pf.space_before, pf.space_after, pf.line_spacing = Pt(0), Pt(10), 1.15
pf.alignment = WD_ALIGN_PARAGRAPH.LEFT

set_font(doc.styles["Heading 1"], 14, bold=True, colour=NAVY)
set_font(doc.styles["Heading 2"], 12, bold=True, colour=NAVY)
set_font(doc.styles["Heading 3"], 12, bold=False, colour=NAVY)
set_font(doc.styles["Title"], 30, colour=NAVY)                 # cover title, up to 30 pt
set_font(doc.styles["Subtitle"], 16, colour=NAVY)
set_font(doc.styles["Caption"], 12, bold=True, colour=NAVY)
set_font(doc.styles["Footer"], 11)
title_ppr = doc.styles["Title"].element.pPr                     # the default Title has a blue rule under it
if title_ppr is not None and title_ppr.find(qn("w:pBdr")) is not None:
    title_ppr.remove(title_ppr.find(qn("w:pBdr")))

def row_props(row, header=False):
    trpr = row._tr.get_or_add_trPr()
    keep = OxmlElement("w:cantSplit")        # never split a row across pages
    trpr.append(keep)
    if header:
        rep = OxmlElement("w:tblHeader")     # repeat the header row on each page
        rep.set(qn("w:val"), "true")
        trpr.append(rep)

def shade(cell, hex_fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcpr.append(shd)

def caption(text):
    """Caption above a table or figure: 'Table 1. Title', bold navy, no full stop. The number is plain text."""
    p = doc.add_paragraph(style="Caption")
    r = p.add_run(text)
    r.bold = True
    r.font.color.rgb = NAVY                 # explicit as well: Word can drop style colours on save
    p.paragraph_format.keep_with_next = True
    return p

def table(header, rows):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"                  # thin black border on every cell
    for cell, text in zip(t.rows[0].cells, header):
        cell.text = ""
        run = cell.paragraphs[0].add_run(text)
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(cell, NAVY_HEX)
    row_props(t.rows[0], header=True)
    for values in rows:
        row = t.add_row()
        row_props(row)
        for cell, text in zip(row.cells, values):
            cell.text = ""
            cell.paragraphs[0].add_run(str(text)).font.size = Pt(10)
    return t

def field(paragraph, instr):
    """A Word field such as PAGE or NUMPAGES, in bold (Word fills in the number)."""
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    run, rpr, bold, txt = OxmlElement("w:r"), OxmlElement("w:rPr"), OxmlElement("w:b"), OxmlElement("w:t")
    rpr.append(bold)
    txt.text = "1"
    run.append(rpr)
    run.append(txt)
    fld.append(run)
    paragraph._p.append(fld)

def page_x_of_y(section):
    """'Page X of Y' centred in the footer, 11 pt, the numbers bold."""
    p = section.footer.paragraphs[0]
    p.style = doc.styles["Footer"]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Page ")
    field(p, "PAGE")
    p.add_run(" of ")
    field(p, "NUMPAGES")
```

Use it as `page_x_of_y(doc.sections[0])`. For a cover with no number, set `doc.sections[0].different_first_page_header_footer = True` and leave the first-page footer empty.

Notes:

- Word can drop style-level colours on save: the recipe sets the navy on the caption runs as well; do the same on any title or heading run you format by hand, and check the rendered result.
- Captions written by a script carry fixed numbers. Word's Insert Caption numbers itself; with fixed numbers, renumber and re-check every mention whenever a table or figure moves.
- python-docx's List Bullet style gives a round black bullet. The navy square bullet comes from the list's bullet definition: make it once in Word (Define New Bullet > Symbol > Wingdings, character 167, the small square; Font > colour navy), save that document as a template, and build from it with `Document("template.docx")`.

## 11. Final check

- Tables, figures and appendices numbered in order of first appearance; every in-text mention matches.
- Captions above, bold navy, no full stop; notes below.
- Headings use the heading styles (the contents page and the navigation pane depend on them).
- No grey, no em dashes in sentences, no ALL CAPS, nothing justified, no text boxes.
- Page numbers present; the date on the cover is right; the file name is short and accurate.
- Spelling checked with the right English (British by default).
- Print preview: no blank pages, no orphaned headings at the bottom of a page, no row split across pages.
