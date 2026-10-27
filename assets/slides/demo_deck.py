"""A demo deck with every pattern, in neutral sample wording:  python demo_deck.py demo.pptx [--navy]

Each slide shows one way to make the content itself the picture. All names, dates and figures are made up:
replace every one with the real source before anything is shared.
"""
import sys

from pptx.enum.text import MSO_ANCHOR

from slide_kit import P, R, SH, SW, B, Deck, L, N, lines_height_in, readable
from slide_patterns import (calendar, card, chevron_stages, closing_band, gutter_dividers, header, icon_column,
                            icon_row, labelled_arcs, need_and_aim, notification, notification_height, phone, run_in,
                            side_panel, stats_row, title_slide, two_zone_structure)

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"]
TITLE = "May we announce both pilots in the customer app?"
LEAD = ["Two pilots run this year, one in spring and one in summer.",
        "A short notice in the app while each is open would reach the teams that can use it."]
CARD_TEXT = [{"right": "Mar 2026",
              "message": "Reports pilot: see your team's month at a glance. Join from the Reports tab and tell us "
                         "what you would change."},
             {"right": "Jul 2026",
              "message": "Alerts pilot: get a message when a number moves. Teams that signed up get a set-up link "
                         "by email on the first day."}]
OFFERS = [("file-text", "A one-page guide for each pilot"),
          ("users-three", "Two drop-in sessions"),
          ("paper-plane-tilt", "A sign-up link for each pilot")]
NEXT = ("Next step:", "a page in the app that lists every pilot and how to join.")
THANKS = "Thank you for your help with this."
OFFERS_HEADING = "What the team can share"
NOTES = [("Script (about one minute)", True),
         ("Sample content: every name, date and figure on these slides is made up.", False),
         ("Say the request in one sentence, then what you can share, then thank them.", False),
         (),
         ("Background, if asked", True),
         ("The exact wording, where it comes from, and any detail that would crowd the slide.", False)]
WINDOWS, CARDS = [], []


def sample_colours():
    """The sample windows and cards, coloured from the current theme (call after Deck())."""
    WINDOWS[:] = [{"start": 2, "end": 5, "label": "March to May", "name": "Pilot one: reports", "colour": P.ramp[0]},
                  {"start": 6, "end": 8, "label": "July to August", "name": "Pilot two: alerts", "colour": P.ramp[1]}]
    CARDS[:] = [dict(c, border=P.ramp[i]) for i, c in enumerate(CARD_TEXT)]


def request_band(d):
    """Version A: top band, title and lead; the real calendar; the real wording on a phone; what is offered."""
    s = d.slide()
    d.box(s, 0, 0, SW, 0.49, fill=P.primary, name="Top band")
    d.title(s, L, 0.72, R - L, 0.62, TITLE, 28)
    d.sentence_lines(s, L, 1.4, R - L, LEAD, 15)
    ph_w = 3.6
    ph_x = R - ph_w
    ax1 = ph_x - 0.55
    y = calendar(d, s, L + 0.85, ax1, 2.3, MONTHS, WINDOWS, year="2026", year_x=L)
    phone(d, s, ph_x, 2.3, ph_w, "Proposed wording", CARDS, app="Customer app")
    d.divider(s, L, y + 0.12, ax1)
    d.one(s, L, y + 0.28, 5, 0.3, OFFERS_HEADING, 14, True, P.primary, name="Offers heading")
    icon_row(d, s, OFFERS, L, y + 0.72, ax1 - L)
    d.divider(s, L, y + 1.42, ax1)
    h = run_in(d, s, "squares-four", L, y + 1.58, ax1 - L, *NEXT)
    d.one(s, L, y + 1.58 + h + 0.32, ax1 - L, 0.34, THANKS, 16, True, P.primary, name="Thanks")
    d.notes(s, NOTES)


def request_panel(d):
    """Version D: the words on a panel in the primary colour; the calendar and the phone on the white side."""
    s = d.slide()
    x0 = side_panel(d, s, 4.4, ["May we", "announce both", "pilots in the", "customer app?"],
                    subtitle_lines=["A request from the", "platform team"], lead=" ".join(LEAD))
    ph_w = 3.6
    ph_x = R - ph_w
    cal_bottom = calendar(d, s, x0 + 0.8, R, 0.5, MONTHS, WINDOWS, year="2026", year_x=x0)
    ph_y = cal_bottom + 0.25
    phone(d, s, ph_x, ph_y, ph_w, "Proposed wording", CARDS, app="Customer app")
    cw = ph_x - 0.3 - x0
    d.one(s, x0, ph_y + 0.42, cw, 0.3, OFFERS_HEADING, 14, True, P.primary, anchor=MSO_ANCHOR.MIDDLE,
          name="Offers heading")                                  # level with the phone's label
    y = icon_column(d, s, OFFERS, x0, ph_y + 0.9, cw)
    d.divider(s, x0, y + 0.22, x0 + cw)
    h = run_in(d, s, "squares-four", x0, y + 0.42, cw, *NEXT, max_lines=4)
    d.text_lines(s, x0, y + 0.42 + h + 0.3, cw, lines_height_in(2, 16) + 0.04,
                 [[B("Thank you for your", 16, P.primary)], [B("help with this.", 16, P.primary)]], name="Thanks")
    d.notes(s, NOTES)


def request_columns(d):
    """Version B: two mirrored columns, each window with its own wording; a band in the primary colour closes."""
    s = d.slide()
    top = header(d, s, TITLE, LEAD, band=False)
    gap = 0.46
    col_w = (R - L - gap) / 2
    whos = ["For every team on monthly reports", "For teams that signed up by email"]
    card_h = max(notification_height(col_w, c["message"], size=13, pad=0.18) for c in CARDS)
    block_h = 0.82 + card_h + 0.3 + 0.4                          # dates, who, cards, then the next step
    top += max(0.0, (6.12 - 0.25 - top - block_h) / 2)            # centred in the space above the closing band
    for i, (wdw, c, who) in enumerate(zip(WINDOWS, CARDS, whos)):
        cx = L + i * (col_w + gap)
        d.one(s, cx, top - 0.04, col_w, 0.34, f"{wdw['label']} 2026", 16, True, readable(wdw["colour"], 16, True),
              name="Dates")
        d.para(s, cx, top + 0.34, col_w, [N(who, 13)], 1, "Who")
        notification(d, s, cx, top + 0.82, col_w, "Customer app", "Proposed wording", c["message"], c["border"],
                     size=13, pad=0.18, right_bold=False, h=card_h)
    d.vdivider(s, L + col_w + gap / 2, top - 0.04, top + 0.82 + card_h)
    run_in(d, s, "squares-four", L, top + 0.82 + card_h + 0.3, R - L, *NEXT, max_lines=2, icon_size=0.36)
    closing_band(d, s, 6.12, OFFERS_HEADING, OFFERS, THANKS)
    d.notes(s, NOTES)


def background_and_aim(d):
    """Three numbers lead into the need, the need into the aim; three labelled arcs say what 'good' means."""
    s = d.slide()
    top = header(d, s, "One dashboard would save every team time")
    y = stats_row(d, s, top - 0.15, [("chart-bar", "", "38", "teams use monthly reports"),
                                      ("clock", "Under", "2 min", "to find any number"),
                                      ("users-three", "About", "450", "people use them each week")])
    need_and_aim(d, s, L, y + 0.35, 6.7, "The need",
                 [N("Teams check ", 15), B("up to five tools", 15),
                  N(" to see how their work is going, and the answers do not always agree.", 15)],
                 "The aim", "Give every team one place to see its own numbers and decide what to do next.")
    labelled_arcs(d, s, 10.4, 6.95, [(2.35, "One place"), (1.82, "Up to date"), (1.29, "Shareable")], stroke_pt=13,
                  size=18)
    d.notes(s, NOTES)


def stages(d):
    """Where things stand: done, now and next as one chevron bar; what belongs to each stage underneath."""
    s = d.slide()
    top = header(d, s, "Two products are live and three more are on the way",
                 ["Reports and the usage dashboard are live, alerts are being set up."], band=True)
    parts = [{"title": "Live now", "icon": "check-circle", "state": "done", "note": "Running since January"},
             {"title": "Being set up", "icon": "plus-circle", "state": "now", "note": "Live from April"},
             {"title": "Coming next", "icon": "hourglass-medium", "state": "next", "note": "Planned for the summer"}]
    cols, cw, y = chevron_stages(d, s, top - 0.1, parts)
    groups = [[("file-text", "Reports", "38 teams"), ("chart-bar", "Usage dashboard", "Every team")],
              [("bell-simple", "Alerts", "First 10 teams")],
              [("chart-line-up", "Forecasts", "Design agreed"), ("download-simple", "Exports", "To be confirmed")]]
    states = ["live", "live", "next"]
    ch, cgap = 1.18, 0.2
    for cx, items, st in zip(cols, groups, states):
        cy = y + 0.2 + ((2 - len(items)) * (ch + cgap)) / 2      # a single card sits centred between the stacks
        for ic, name, caption in items:
            card(d, s, cx, cy, cw, ch, state=st, name=f"Card {name}")
            d.icon(s, ic, cx + 0.25, cy + 0.3, 0.42)
            d.one(s, cx + 0.85, cy + 0.24, cw - 1.05, 0.34, name, 16, True, P.primary, name=f"Card name {name}")
            d.one(s, cx + 0.85, cy + 0.62, cw - 1.05, 0.3, caption, 13, False, "000000", name=f"Card caption {name}")
            cy += ch + cgap
    gutter_dividers(d, s, cols, cw, y + 0.05, y + 0.2 + 2 * ch + cgap)   # the rules end where the cards end
    d.notes(s, NOTES)


def structure(d):
    """Who decides what: levels down the left with their role inside each box; one level's members in a panel."""
    s = d.slide()
    top = header(d, s, "Three groups steer the work and one team delivers it", band=False)
    two_zone_structure(d, s, top - 0.2, [("Steering group", "Sets direction and approves funding"),
                                          ("Programme board", "Agrees priorities each quarter"),
                                          ("Working groups", "Bring specialist knowledge"),
                                          ("Delivery team", "Builds, runs and reports")],
                       joined=2, members=["Data platform", "Security", "Customer success", "Finance", "Legal",
                                          "Support", "Design"])
    d.notes(s, NOTES)


def main(out, theme="green"):
    d = Deck(theme=theme)
    sample_colours()
    s = d.slide()
    title_slide(d, s, ["Two pilots,", "one clear", "picture"],
                "How the reports and alerts pilots will reach the teams who can use them.",
                "Platform team, 2 March 2026", parts=["Where we are", "What we ask", "What happens next"])
    d.notes(s, NOTES)
    request_band(d)
    request_panel(d)
    request_columns(d)
    background_and_aim(d)
    stages(d)
    structure(d)
    d.save(out, author="", title="Demo deck")
    print("saved", out)


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    main(args[0] if args else "demo.pptx", "navy" if "--navy" in sys.argv else "green")
