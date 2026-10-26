# HTML explainers, guides and packs

Local, self-contained HTML pages: an explainer, a how-to guide, a walkthrough of changes, a "pack" that gathers everything someone needs for a task (text to copy, steps to follow, choices to make, files to download).

Start from `assets/html/explainer-template.html`. It holds the patterns below, already styled and tested (copy buttons, rich copy for email, remembered ticks and picks, the answer line, a searchable list, a download, phone width, print); only picture options are described here without an example. Delete the patterns a page does not need.

## 1. The feel

The page reads like a well-set document that happens to have buttons.

- Everything sits on white, in one column (at most 1040 px wide; paragraphs at most 860 px).
- The only strong mark is a 10 px navy bar at the top.
- Navy is for structure alone: headings, labels, table header rows, button outlines and the answer bar.
- Text is a soft near-black, Tahoma 15 px, line height 1.55.
- Every rule and box edge is a hairline in a pale navy tint, so nothing looks grey or boxed in.
- Colour appears only when it means something: yellow for a gap, brown for "check this", green for "done", blue for a link.
- Nothing competes for attention: no shadows, gradients, icons, badges, pills, tinted cards, accent bars or capital-letter labels.

## 2. Tokens

| Token | Value | Use |
|---|---|---|
| `--navy` | #243168 | 10 px top bar, h1 to h3, field labels, run-in labels ("Why:"), button text and outline, table header fill, total-row rule, answer bar |
| `--ink` | #1D1D1B | all body text and typed text |
| `--line` | #D3D8EA | every hairline: rule above each h2, box borders, image borders, table row rules |
| `--link` | #0563C1 | links (so links never look like the navy headings) |
| `--flag` | #8A4B00 | brown "Check:" notes, the warning state of a button |
| meta text | #4A4F63 | `.meta`, `.hint`, `.why`, sources, captions under small pictures |
| input border | #7A84AE | inputs, selects, text areas (3:1 against white) |
| done | #1E6B3A | "Copied" and "Downloaded" button state |
| missing | #FFF3B0 | behind a live field or link that is still empty |

Type: `Tahoma, Verdana, sans-serif` (system fonts only, no web fonts). Sizes: h1 28 px (line height 1.2), h2 21 px, h3 16 px, lede 16 px, body 15 px, tables 13.5 px, small text 13 px, captions under small pictures 12 px (a figure caption above a figure is body size, bold navy). Sentence case everywhere. One radius, 3 px. No shadows.

Spacing: h2 has 44 px above, then a 1 px `--line` rule, then 14 px; h3 22 px above and 6 px below; paragraphs 6 px above and 10 px below; figures 22 px above and 30 px below.

## 3. Page order

1. **Header:** the 10 px navy bar, the title (h1), a lede (what the page is for, key dates in bold, how to answer), a `.meta` status line ("Prepared 2 March 2026. Nothing on this page has been sent or shared yet."), a "Start with ..." sentence, then a contents list.
2. **Contents:** an `ol.toc` in two columns (one on phones). Its numbers match the numbered h2 headings.
3. **What to settle first** (optional, unnumbered): points to agree before anything is sent, with a Copy list button. Time-critical requests go here as ready-to-send messages in copy boxes ("Do today").
4. **Numbered sections:** h2 "1. Name", "2. Name". Each opens with one or two sentences: what is in it and what every option already respects.
5. **Steps inside a section are lettered** (A, B, C) so they never clash with the numbered sections. Cross-references name both: "section 2 (Page title)".
6. **End:** a small footer, "How it was checked: ...", stating the checks as the final state, never as a story of fixes. On pages with decisions, the fixed navy answer bar closes the page instead.

## 4. Patterns

**Copy box.** Text on the left, its button top right; on phones the button drops under the text. Button labels say what they copy: "Copy list", "Copy message", "Copy description", "Copy email for Outlook", "Copy as plain text". `.copybox.narrow` for a short value such as a colour code. Never put a copy box inside a `<p>`.

**Rich copy for email** (`data-rich="1"`): copies HTML and plain text together so bold, lists and links survive a paste into Outlook or Gmail. Empty links copy as "[link to follow]", never as the page's own file path, and the button warns in brown ("Copied, links still missing"). In the plain version every link is written "link text: URL".

**Options.** The suggested option is pre-selected and marked "(suggested)" in plain text, never "my pick". Each option: a number or letter in bold, the wording, a Copy button if it will be pasted somewhere, then run-in "Why:" and "Watch out:" in bold navy. Picture options put the radio in the caption and a phone-size preview beside the full picture.

**Answer bar** (only when the page asks for decisions): fixed at the bottom, navy, it composes a one-line answer such as `1=B 2=2 3=a` from the radio groups, starting at the suggestions, with a Copy button and a key ("1 = layout, 2 = title, 3 = audience"). Sending it unchanged means "I accept the suggestions". Picks are remembered across reloads. Name the radio groups `q1`, `q2`, `q3` and list them in the script's `CONFIG.picks`: the number after the q becomes the key in the line.

**Lettered checklist.** `ul.checks`: a tick box, the action in the words the screen uses, then a smaller `.why` line (why it matters or what goes wrong if skipped). Group headings carry the menu path in brackets: "B. Options (Settings, then Sharing)". "Tick as you go; this browser remembers the ticks."

**Fill in once.** A box of labelled fields placed right before what it fills, saying exactly what it fills. Live values appear in the text below; anything still missing is shaded yellow. Use real input types (`type="number"` with limits). "This browser remembers what you type."

**Tables.** Navy header row with white bold text; one hairline under each row; no vertical rules, no stripes. Numbers centred; a bold total row under a 2 px navy rule. Gaps are named, never blank ("none published", "To be confirmed."). Wide tables sit in `.hscroll` with "Scroll sideways for more columns." on phones; long lists sit in `.scroll` with a sticky header, a search box (accent-insensitive), a group filter and a live "N shown".

**Figures.** Caption above, "Figure 1. Title" in bold navy; the picture in a hairline border; inline SVG with a `<title>`, or a PNG at twice the display size. See `figures.md`.

**Before and after.** A table with Before and After columns: added words highlighted (`ins`, #FFF1A8, no underline), removed words red with strikethrough (`del`, #C00000), row headers in bold navy, a one-line result under the table ("Applied in part: ..."). The full list can sit in a `<details>`.

**Terms.** A definition grid: bold navy terms in a 150 px column beside one-sentence meanings.

**Disclosures.** Provenance and secondary material go in `<details><summary>`: "Where each answer comes from", "Pages these steps come from". The print view opens them.

**Downloads.** Files are embedded as base64 and saved through a Blob, so the page works offline; the button then says "Downloaded". Give the file facts beside it ("1200 x 400 px, 48 KB, PNG"). Embed each file once and reuse it; heavy pages are slow to open and to email.

**Sources.** Write sources out in full, never as codes, and add a one-line key. An unknown answer says "To be confirmed." with the source "Answer not in the sources". External links open in a new tab (`target="_blank" rel="noopener"`).

## 5. Writing on the page

- The ten-second test: a busy reader sees what the thing is and what to do.
- Lead each paragraph with its point and split dense ones: one point per paragraph, the key fact first, never a date buried at the end.
- Absolute dates ("Monday 2 March"), never "tomorrow"; 24-hour times with the time zone beside them.
- Plans are written as plans ("is due to be published on ..."); nothing promised that is not agreed.
- Conditions come before the Copy button they apply to ("Before you paste it: ...").
- Captions never claim more than the picture shows.
- Neutral wording: "(suggested)", never "my pick"; no "prepared for [name]" if the page may be passed on.
- Only state what a source says; say "the documentation does not mention ..." rather than "there is no ...".

## 6. Build rules

- One self-contained file: CSS in `<style>`, the script at the end of `<body>`, images and downloads inlined. No external fonts, scripts or requests, so it works from a USB stick or an email attachment.
- Set `lang` (`en-GB` by default) and an author meta tag.
- Keep the wording in one place (a content file or a block at the top of the builder), each fact with its source beside it, and let the builder only assemble.
- Before writing the file, assert: no leftover placeholders, no duplicate ids, no em dashes, no tool names or internal shorthand.
- No dark mode: the page sets a white background on purpose.
- Check it in a browser before sharing: no console errors, every Copy button puts the right text on the clipboard, nothing scrolls sideways at 390 px, print preview shows everything.
- Never link to local files (`file:///`) in a page that may travel: the path breaks on any other machine and shows the user name.

## 7. Logo (only when asked)

The page has no logo by default: the navy bar is the brand. When a logo is wanted (for example a subtle MongoDB logo), put it once, top right of the header, about 24 px tall, on white, using the official file unaltered (see "Logos" in SKILL.md). Inline it so the page stays one file:

```python
import base64, pathlib
svg = pathlib.Path("assets/logos/mongodb/mongodb-logo-slate.svg").read_bytes()
src = "data:image/svg+xml;base64," + base64.b64encode(svg).decode()
# then: <img class="logo" alt="MongoDB" src="{src}"> inside <div class="brand">, after the h1
```
