# Google Sheets

The look and the structure are the ones in `spreadsheets.md`. This file says what changes in Google Sheets and how to build it. Starter: `assets/google-sheets/NavyStyle.gs` (theme, menu, a Dashboard, Tracker and Lists builder, the code behind a screen) and `assets/google-sheets/Dialog.html` (a screen in the house look).

To use them: in the spreadsheet, Extensions > Apps Script; paste `NavyStyle.gs`, add an HTML file named `Dialog` with the contents of `Dialog.html`, save, reload the spreadsheet. Run `buildStarter()` once on a new, empty spreadsheet to see the full look, then adapt names and columns (it removes the empty first sheet only when it is called "Sheet1"; in another language, delete it by hand). The first run asks the owner to authorise the script.

## 1. Font

- Use **Verdana** in Google Sheets (and Google Docs). It is in the font menu and in the spreadsheet theme's font list; Tahoma is not on the theme list, and whether it can be picked cell by cell is not documented. Verdana is Tahoma's wider sibling: the same x-height, about 14% wider (12% in bold). Keep the same point sizes and make columns about 15% wider.
- Set it twice: as the spreadsheet theme font, so anything typed later uses it (Format > Theme > Customise, or `applyNavyTheme()`), and on every styled range.
- Screens (HTML dialogs) are drawn by the browser: keep `Tahoma, Verdana, sans-serif` there.

## 2. Theme and units

- `applyNavyTheme()` sets the theme font, the text colour (ink #1B1F2A), the first accent (navy) and the link colour (#0563C1). Cells with a format of their own keep it.
- Widths and heights are in pixels. Column width: (Excel width x 7 + 5) x 1.15, for example an Excel width of 22 becomes 183 px. Row height: points x 4/3, so a 26 pt header row is 35 px, a 20 pt row 27 px, a 32 pt title row 43 px.

## 3. What carries over directly

| In Excel | In Google Sheets |
|---|---|
| White-on-navy header row | fill and font on the header range (`header_()` in the starter) |
| Freeze panes | View > Freeze, or `setFrozenRows` / `setFrozenColumns` |
| Gridlines off on page-like sheets | View > Show > Gridlines, or `setHiddenGridlines(true)` |
| Tab colours | `setTabColor` |
| Dates `d mmm yyyy` | Format > Number > Custom date and time, or `setNumberFormat('d mmm yyyy')` |
| Strict drop-down | Data validation, "Reject the input" (`setAllowInvalid(false)`) |
| Warning drop-down | "Show a warning" (`setAllowInvalid(true)`, the default) |
| Input prompt on a cell | help text, shown when the cell is selected (`setHelpText`) |
| Dates from 2020 up to today | custom formula rule `=AND(ISNUMBER(E3),E3>=DATE(2020,1,1),E3<=TODAY())`, reject |
| Status font colours, amber "needs a look" fill, faded rows, strikethrough | conditional formatting with custom formulas |
| Protect the dashboard only | `protect().setWarningOnly(true)`: anyone can still edit after a warning; for a hard lock remove the other editors (the owner can always edit) |
| Filter buttons on the list | Data > Create a filter on the header row, or `range.createFilter()` |
| FILTER, UNIQUE, COUNTIFS, MINIFS | the same functions; FILTER and UNIQUE spill on their own; refer to open-ended ranges such as `Tracker!$D$3:$D` |

## 4. What works differently

- **Only the first true rule formats a cell.** Google evaluates conditional formatting rules in the order listed and "the first rule found to be true will define the format". Every rule behaves like Excel's stop-if-true, and formats from several rules do not add up. Add rules most specific first: the "needs a look" fills, then the status colours, then the whole-row fade. A cell shows one rule's format only.
- **Conditional formatting sets only** bold, italic, strikethrough, text colour and fill. No borders, so the "hairline only under rows that hold data" cannot be done: keep gridlines on in data sheets (as the Excel version does) and draw fixed hairlines under the rows of each dashboard box.
- **Rules cannot point at another sheet** directly. Use `INDIRECT`: `=AND($D3<>"",COUNTIF(INDIRECT("Lists!B5:B"),$D3)=0)` marks a status that is not on the Lists tab.
- **No icon sets and no data bars.** Bars: a `SPARKLINE` in the bar column, `=SPARKLINE(C10,{"charttype","bar";"max",MAX($C$10:$C$14);"color1","#5B6AA6"})`. A checklist's finished step: the green fill and dark green text on the finish date; if a tick mark is wanted, put it in a narrow column worked out by formula, never typed.
- **Drop-downs can show as coloured chips.** This style has no chips: in Data validation > Advanced options > Display style, choose Arrow (closest to Excel) or Plain text. Scripts cannot set this; set it by hand once per rule.
- **No Center Across Selection.** Put each group label in the group's first cell, left-aligned, with the medium navy rule under the whole row.
- **Print** settings are chosen in File > Print (paper A4, landscape for wide sheets, scale "Fit to width", sheet name and page numbers in headers and footers, custom page breaks); scripts cannot set them.
- **Dates typed by people** are read according to the spreadsheet's locale (File > Settings). Set it once (for example United Kingdom), then type a test date such as 16 Mar 2026 and check it lines up on the right like a date.

## 5. Buttons

Buttons in Google Sheets are drawings or images with a script assigned. They run only in a web browser, never in the mobile apps, so the sheet must still work by typing.

- **A drawing** (the usual way): Insert > Drawing > Shape > Rounded rectangle. Fill: custom colour #243168. Border: transparent. Text: the label in Verdana 10 bold white, centred; an optional second line in Verdana 8 regular #C9CEDF ("one person", "search and edit"). Save and close, size it (about 53 px tall on a dashboard, 32 px on a list sheet), then open the drawing's three-dot menu > Assign script and type the function name without brackets. Make the next button by copying the first one, so they all match. Secondary button ("Dashboard", to go back): white fill, 1 px navy border, navy bold text.
- **An image:** a script can insert a PNG over the cells (`sheet.insertImage(blob, column, row)`) and wire it with `assignScript('functionName')`; scripts cannot create drawings.
- **Always add the same actions to the custom menu** (`onOpen()` in the starter). Menus need a script that belongs to the spreadsheet (Extensions > Apps Script inside it).
- Place a dashboard's buttons in one row under the title, equal widths, about 10 px apart, and leave the row free of text.

## 6. Screens

Excel's UserForms become HTML dialogs (`Dialog.html`, opened by `showAddDialog()`):

- Google's dialog frame shows the title. The page adds a thin navy bar, a one-line subtitle, sections with a bold navy heading over a fine rule, flat fields with labels above them (a #7A84AE border, stronger than the Excel screens' #C9CEDF so fields stay visible on any monitor), a message line (red for a problem, green for success), the main action as a navy button at the bottom right and Cancel at the bottom left. Esc closes the screen. Every field and button has a hover note (`title`).
- Drop-down values are read from the Lists tab (`listValues()`), so the screen and the sheet offer the same choices.
- The server side writes plain values only, keeps text that starts with `=`, `+`, `-` or `@` exactly as typed, and refuses a date it cannot read with an example ("Type the date like 16 Mar 2026.").

## 7. Charts

Insert > Chart, then in Customise: series colours navy #243168 first, then #5B6AA6, #8F99C2, #C9CEDF; chart and axis font Verdana; no 3D, no shadows; gridlines #D3D8EA or none; a short title (or none, if the sheet's heading already says it); data labels when there are only a few values. One message per chart, as in `figures.md`.

## 8. Working through the Sheets API or a connector

If you change the spreadsheet through the Google Sheets API (for example a connector that sends `spreadsheets.batchUpdate` requests) instead of pasting Apps Script, the same look takes these requests. Colours are fractions of 255: navy #243168 is `{"red": 0.141, "green": 0.192, "blue": 0.408}`.

| Look | Request |
|---|---|
| Theme font Verdana, theme colours | `updateSpreadsheetProperties` with `spreadsheetTheme` (`primaryFontFamily`, and all `themeColors` pairs: TEXT, BACKGROUND, ACCENT1 to ACCENT6, LINK) |
| Header row, fonts, fills, alignment, wrap, number formats | `repeatCell` with `cell.userEnteredFormat` (`backgroundColor`, `textFormat`, `horizontalAlignment`, `verticalAlignment`, `wrapStrategy`, `numberFormat`) and a matching `fields` mask |
| Rules and hairlines | `updateBorders` (`top`, `bottom`, ... each with `style` and `color`) |
| Column widths, row heights (pixels) | `updateDimensionProperties` with `pixelSize` |
| Frozen rows and columns, gridlines off, tab colour | `updateSheetProperties` (`gridProperties.frozenRowCount`, `frozenColumnCount`, `hideGridlines`, `tabColor`) |
| Status colours, "needs a look" fill | `addConditionalFormatRule` with a `booleanRule` of type `CUSTOM_FORMULA`, added most specific first |
| Drop-downs and date rules | `setDataValidation` (`ONE_OF_RANGE` or `CUSTOM_FORMULA`, `strict` true to reject, `inputMessage`, `showCustomUi` true) |
| Protected dashboard | `addProtectedRange` with `warningOnly` true |
| Filter buttons | `setBasicFilter` |

Then check the result in the browser, as for any build: the API sets values exactly as sent, so fonts, colours and rule order need the same final check.

## 9. From an Excel workbook

VBA macros, Excel shape buttons and UserForms do not run in Google Sheets. When a workbook in this style moves to Google Sheets, rebuild the buttons as drawings and the screens as dialogs, and check the fonts, the conditional formats (icon sets and data bars have no equivalent) and the drop-down display style after uploading.

## 10. Final check

- One font everywhere (theme and styled ranges); columns widened for Verdana.
- Drop-downs set to Arrow or Plain text, not chips.
- Conditional formatting rules ordered most specific first.
- The menu and the buttons work in a browser; on a phone the sheets still work by typing.
- The dashboard is protected (at least with a warning); the data sheets are not.
- The locale is set, and a typed date reads as a date.
