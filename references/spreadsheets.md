# Spreadsheets (Excel and Google Sheets)

Registers, trackers, checklists, contact lists, dashboards. This file is the look and the structure; `google-sheets.md` says how to get the same result in Google Sheets. A tested openpyxl starter is in `assets/excel/navy_workbook.py` (run it to see a demo workbook).

## 1. The idea

- The workbook reads like a quiet web page, not a grid: one typeface, navy for structure, near-black text, pale navy hairlines, white space.
- **Nobody can break the data.** Data sheets hold plain, typeable values only. Every count, check and calculation lives on a read-only dashboard. Deleting or overtyping any cell in the data can never break anything.
- **Colour means something.** Status, attention and links get colour; nothing gets colour for decoration. No grey anywhere: every neutral is a tint of navy.
- **Automation only helps.** Buttons and screens (VBA in Excel, Apps Script in Google Sheets) make entry easier, but everything must still work by typing, because some people will open the file where scripts do not run (Excel in a browser, a phone, a copy).

## 2. Pick the shape of the workbook

| Shape | Sheets, in this order | Use for |
|---|---|---|
| Register or tracker | Dashboard, the list (one table), then reference tabs (Lists, and any lookup table) | records that move through statuses: requests, access, issues, assets |
| Plan or checklist | An explanation page, then a checklist template (copied once per project) | a repeatable process with steps, dates and owners |
| Simple list | Overview (counts), the list (one table), Lists | contacts, inventories, invitations |

- The first sheet is the one people read: a dashboard or overview that works like a web page.
- No separate "Read me" or "Summary" tab: explanations live where they are needed (a subtitle, a note under a box, a hint line, a prompt on a cell).
- Columns of the list follow the form or document the information comes from, in the same order, so people recognise them.

## 3. Type

Tahoma everywhere, set as the workbook's own font (the Normal style) and as the theme's heading and body fonts, so new cells, links and anything added later come out in Tahoma too.

| Element | Size | Style | Colour |
|---|---|---|---|
| Dashboard title | 20 | bold | navy |
| Title on other sheets | 14 to 16 (18 on an explanation page) | bold | navy |
| Subtitle (one sentence) | 10 | regular | soft |
| Section title on a dashboard | 13 | bold | navy |
| Sub-section title | 11 to 12 | bold | navy |
| Group label over columns | 10 | bold | navy |
| Header row of a table | 10 | bold | white on navy fill |
| Header row inside a dashboard box | 9.5 | bold | navy, thin navy underline |
| Data and body | 10 (10.5 to 11 inside dashboard boxes) | regular | ink |
| Row labels | same size | bold | ink |
| Total row | 11 | bold | navy |
| Tile number | 22 | bold | navy |
| Tile caption | 8.5 | regular | soft |
| Hints and notes | 8.5 to 9 | italic | soft |
| Anything to be confirmed | 10 | italic | navy |
| Print footer | 8 | regular | default |

Bold is for titles, headings, header rows, row labels, totals and numbers that matter. Italic is only for secondary information: notes, hints, "to be confirmed", "not needed".

## 4. Colour

| Role | Hex | Use |
|---|---|---|
| Navy | #243168 | titles, headings, header-row fill, rules, totals, tile numbers, buttons, tab colour of the sheets people work in |
| Ink | #1B1F2A | all data and body text (never pure black) |
| Soft | #2C3A78 | subtitles, hints, notes, tile captions |
| Line | #D3D8EA | hairlines under rows that hold data |
| Rule | #C9CEDF | tile borders, field borders on screens, the second line of text on a navy button |
| Picker fill | #EEF0F7 | the one input cell on a dashboard |
| Tint | #F4F5FA | read-only boxes on screens |
| Bar | #5B6AA6 | data bars and bar charts |
| Reference tab | #8F99C2 | tab colour of reference sheets (Lists, lookups) |
| Faded | #7A82AE | whole-row font of ended or closed records |

**Status, as font colour only** (never a fill): done #1E7B3C (bold for the main "live" state); waiting on us #8A5300 (bold when on hold, italic for old records nobody has checked); with an outside party #00618A; stopped, refused or ended #A4262C (bold while it is still being ended).

**Attention:** needs a look #FFE699 as a fill on the exact cell that is wrong; a problem count turns bold #A4262C only when it is above 0; a waiting count turns bold navy above 0; "fill in once" cells #FFF4CC inside a thin rule-colour box; a notice banner #FFF4CE with a medium #C27C0E left edge.

**Progress checklists** use fills on just two cells per row: finished #375623 text on #E2EFDA with a green tick; in progress #FCE4D6 fill and #B5520F text; a finish date before the start date in bold #C00000 with no tick; "Not needed" as strikethrough on the task and navy italic in the date cell.

Every text colour above passes 4.5:1 contrast on white, with two deliberate exceptions: the faded colour for ended records (3.7:1, meant to recede) and the in-progress orange on its own fill (4.1:1). Machine-readable values: `assets/palette.json`.

## 5. Layout

- **Page-like sheets** (dashboards, explanation pages): gridlines off, a narrow gutter column first (width 2) and last, content from column B.
- **Data sheets:** gridlines may stay on (they help typing), or draw a hairline grid in the line colour. Never merge cells in a data sheet.
- **Column widths** suit the content: names 14 to 22, email 32, organisation 26 to 30, dates 12 to 13, short codes 6 to 15, notes 60 to 80.
- **Row heights:** rows people type into keep Excel's automatic height, so they grow with the text; never fix them. Fixed heights only for rows nobody types into: titles (32), header rows (26), group labels (20), box titles (24), spacers (6 to 16), and wrapped notes (lines x 12.75 + 5 pt for Tahoma 10).
- **Header row:** 26 pt, bold white Tahoma 10 on navy, vertically centred, wrapped where needed.
- **Group labels** above the header ("Who", "Tracking", "Notes"): 10 bold navy, centred across the group's columns without merging (Center Across Selection), with a medium navy rule under the whole row.
- **Freeze** the header rows and the columns that identify a row (for example name and email), so they stay in view.
- **Tables:** one Excel table for the list (filter buttons on, `TableStyleLight1`, no row or column stripes). The navy header comes from direct formatting. Rows are separated by a hairline that appears only under rows that hold data (conditional format), never by banding.
- **Alignment:** labels left, numbers right, tiles centred; long text top-aligned and wrapped.
- **Dates** always shown as `d mmm yyyy` (16 Mar 2026), so a misread month is visible, and asked for "typed like 16 Mar 2026": day-first dates such as 16/03/2026 are not read as dates on computers set to US dates.

## 6. The dashboard

Build it top to bottom like a page:

1. Optional notice banner (row height 36), the one boxed notice this style allows: #FFF4CE fill, medium #C27C0E left edge, 9.5 ink, first sentence bold. For example: "**Open this file in the Excel desktop app to use the buttons.** In the browser, type in the list as usual."
2. Title (20 bold navy, row 32) and a one-sentence subtitle (10 soft).
3. A row of buttons (row 46, buttons 40 pt tall).
4. A medium navy rule across the content (a 6 pt row with a bottom border).
5. Boxes separated by blank rows. Each box: a title (13 bold navy, row 24), a header row (9.5 bold navy, wrapped, thin navy underline, row 28), body rows (11 ink, first column bold, row 20), a hairline under each filled row only, and a total row (11 bold navy between thin navy rules).
6. Tiles for the few numbers that matter: a 22 bold navy number over an 8.5 soft caption, both inside one thin rule-colour box.
7. One input at most, the picker (for example "choose a team"): #EEF0F7 fill, medium navy box, 13 bold navy, a right-aligned label to its left and a 9 italic soft hint ("click the box, then the arrow"). Everything on the page answers to it.
8. Data bars in #5B6AA6, solid, no border, starting at 0, with the value hidden: the bar column holds a copy of a count that is already shown in its own column.
9. A "Needs a look" box: one row per data-quality check, using exactly the same tests as the amber fills on the list. A "To do" box: lines waiting at each step and the oldest date.
10. Short notes under boxes: 8.5 italic soft, wrapped, top-aligned.

Protect only the dashboard (no password needed), with the picker unlocked; tables do not grow on a protected sheet, so never protect the data sheets. Formula cells can also refuse typing without protection: a custom validation rule of `FALSE` with the message "Worked out automatically".

## 7. The data sheet (the list)

- Plain values only: no formulas, no helper or "check" columns, no lookups. A region or category is typed like any other value.
- **Drop-downs:** strict ("Stop") only where a dashboard count depends on the value; a warning everywhere else, ending "Click Yes to keep it anyway." The values live on the Lists tab.
- **Status** shown by font colour; ended records fade to #7A82AE across the row and stay in the list with a date and a reason (records are never deleted).
- **Needs a look:** the amber fill sits on the exact cell that is wrong (not a whole row, not a helper column) and clears itself when the cell is fixed.
- **Dates:** validation from 1 Jan 2020 to 31 Dec 2040, or up to today for things that already happened.
- **Notes and History:** a Notes column for what people write; if scripts change lines, a History column at the far right for their stamps, `d mmm yyyy, name: what changed`.
- Text that looks like something else (8 Jun, +1 555 0100, -x, =x) is kept exactly as typed.
- **Lists tab:** the drop-down values as small side-by-side tables (2-wide spacer columns), navy headers, and a red 8.5 bold italic "Fixed:" warning above any list the dashboard counts depend on.

## 8. The checklist template

- Top block, frozen with the header: a title (16 bold navy) and a short hint (9 navy, right-aligned); three labelled blocks ("Project", "Team", "Where things stand") with the values filled once in #FFF4CC cells inside thin boxes, and the "where things stand" values worked out by formula.
- Header row 22 pt with thin white separators between the cells, repeated on every printed page.
- Stage rows: the stage number 12 bold navy, a medium navy rule above, the stage name 11 bold navy followed by its timing in 9 italic navy, and a count on the right ("2 of 8 finished") in 9 italic navy.
- Step rows: a hairline box round every cell, top-aligned and wrapped, automatic height.
- **Status comes from the dates alone**, so nobody keeps a status column up to date: a finish date gives the green tick and fill; a start date without a finish gives orange; "Not needed" strikes the task through; a finish before the start shows red and suppresses the tick, so a mistake never looks finished.
- Formulas, drop-downs and colours cover rows well beyond the last step (for example to row 200), so added rows are counted.

## 9. Messages on cells

- **Title:** the field name or the problem, at most 32 characters ("Not on the list", "Type a date").
- **Message:** a plain instruction with an example in the exact display format, the keyboard shortcut and the reason, at most 255 characters. For example: "Type a date, for example 16 Mar 2026, or press Ctrl+; for today." / "This place is not on the Places tab yet. Click Yes to keep it anyway." / "Pick a status from the list (the Dashboard counts use it)."
- No jargon, no exclamation marks.

## 10. Buttons (Excel desktop)

- Flat rounded rectangles (a shape with a macro assigned), never form controls or ActiveX. Free floating (they never move or resize with the cells), alt text set to the label, no shadow.
- **Primary:** navy fill, no outline, white bold Tahoma. **Secondary** (only for going back, for example "Dashboard"): white fill, 1 pt navy outline, navy bold text. **Danger** actions belong on a screen, not on the sheet.
- **On a dashboard:** one row of equal buttons, 40 pt tall, 8 pt apart, filling the content width; two lines each: the label (10 bold) and a short caption (8 regular, #C9CEDF on navy): "Add a request / one person", "Find a person / search and edit".
- **On a list sheet:** 24 pt tall, 9.5 bold, 6 pt apart, in the frozen top row so they stay in view.
- Labels are short verb phrases in sentence case.
- The code that draws them (Python driving Excel through pywin32; `part()` calls a property that takes arguments, which plain dispatch cannot):

```python
import pythoncom, win32com.client as win32
NAVY, WHITE, SOFT, CAPTION_ON_NAVY = 0x683124, 0xFFFFFF, 0x783A2C, 0xDFCEC9   # Excel stores colours as BGR

def part(obj, name, *args):
    disp = obj._oleobj_
    r = disp.Invoke(disp.GetIDsOfNames(name), 0, pythoncom.DISPATCH_PROPERTYGET | pythoncom.DISPATCH_METHOD, True, *args)
    return win32.Dispatch(r)

def sheet_button(ws, left, top, w, h, label, macro, primary=True, caption=None, size=10):
    s = ws.Shapes.AddShape(5, left, top, w, h)          # 5 = rounded rectangle
    s.Name = "btn" + macro
    s.Fill.Visible = True
    s.Fill.Solid()
    s.Fill.ForeColor.RGB = NAVY if primary else WHITE
    if primary:
        s.Line.Visible = False
    else:
        s.Line.Visible = True
        s.Line.ForeColor.RGB = NAVY
        s.Line.Weight = 1
    s.Shadow.Visible = False
    tf = s.TextFrame2
    tf.MarginLeft = tf.MarginRight = 3
    tf.MarginTop = tf.MarginBottom = 1
    tf.VerticalAnchor = 3                               # middle
    tf.WordWrap = True
    tr = tf.TextRange
    tr.Text = label + ((chr(13) + caption) if caption else "")
    tr.ParagraphFormat.Alignment = 2                    # centre
    tr.Font.Name = "Tahoma"
    tr.Font.Size = size
    tr.Font.Bold = True
    tr.Font.Fill.ForeColor.RGB = WHITE if primary else NAVY
    if caption:
        sub = part(tr, "Characters", len(label) + 2, len(caption))
        sub.Font.Size = 8
        sub.Font.Bold = False
        sub.Font.Fill.ForeColor.RGB = CAPTION_ON_NAVY if primary else SOFT
    s.OnAction = macro
    s.Placement = 3                                     # free floating
    s.AlternativeText = label + ((": " + caption) if caption else "")
    s.Locked = True
    return s
```

Place a dashboard row of buttons by measuring the content width: `w = (right - left - gap * (n - 1)) / n`, top 3 pt below the row top.

## 11. Screens (Excel UserForms; Google Sheets dialogs)

Every screen has the same look:

- A navy header bar, 30 pt tall, with the title in 12 bold white; a 9 soft subtitle under it.
- White body, 18 pt margins; on wide screens a three-column grid with 10 pt gutters.
- Section headings 10 bold navy over a 1 pt #C9CEDF rule.
- Field captions 9 ink; fields flat, 18 pt tall, white, with a #C9CEDF border and 9 pt ink text; read-only boxes on the #F4F5FA tint.
- The main action is a navy button with white text at the bottom right; other actions to its left, 8 pt apart; Cancel or Close at the bottom left (Esc closes); a dangerous action is a white button with red #A4262C text at the bottom left, and Cancel then moves next to the main button.
- A message line above the buttons, 9 bold: red for a problem, green for success.
- A hover note on every box and button; required fields marked "*", explained in the subtitle.
- Screens stay under about 490 pt tall, so they fit a 1080p laptop screen at 150% scaling.

## 12. Print

A4; landscape for wide sheets; fit to one page wide; margins 0.5 in left and right, 0.6 in top, 0.7 in bottom; a centred Tahoma 8 footer "Sheet, page N of M" (`&A, page &P of &N`); the header row repeated on every page; page breaks only where a section starts, so no heading is left alone at the foot of a page.

## 13. File hygiene

- Save a script-built workbook once in Excel, so the calculated values are stored and counts show in previews and in any viewer.
- Set the author fields to the person who owns the file, and remove any local folder path Excel stores.
- Once people work in a workbook it is the live record: never rebuild it from a script over their data. Change it with a small, repeatable upgrade step on a copy, then swap.
- Formulas have a hard 8,192-character limit. If one gets close, stop and suggest helper cells on the dashboard; do not refactor a working formula unless asked.
- openpyxl cannot write dynamic-array formulas (FILTER, UNIQUE) or calculated table columns, cannot draw buttons, and drops some cross-sheet drop-downs it did not create when re-saving a file. Functions newer than Excel 2007 need the `_xlfn.` prefix when written from openpyxl (`_xlfn.MINIFS`, `_xlfn.TEXTJOIN`, and in conditional formats `_xlfn.ISFORMULA`), or Excel shows #NAME? or silently ignores the rule.
- Excel drops named styles nothing uses when it saves, so set the theme fonts and theme link colour (#0563C1) rather than relying on a custom link style. `navy_workbook.finish()` does both.

## 14. Final check

- Tahoma everywhere, including the Normal style, the theme, buttons and the print footer.
- No grey: no grey fills, borders, gridline-coloured lines or text. No pure black text (ink instead).
- Data sheets hold no formulas; the dashboard holds them all and is the only protected sheet.
- Strict drop-downs only where counts depend on them; every message has an example.
- Dates show as `d mmm yyyy`.
- Typed rows grow with their text; nothing is cut off.
- The workbook still works by typing alone, with scripts switched off.
