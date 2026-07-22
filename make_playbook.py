"""
make_playbook.py
===============================================================================
Generate a self-contained HTML report — the BUY-BEFORE / SELL-AFTER playbook for
every Nifty 50 stock, across all four event types. Reuses stock_server's ranking
engine so the numbers match the dashboard exactly.

    python make_playbook.py            # -> buy_sell_playbook.html

Purely additive: writes a new HTML file, touches nothing else.
"""
from __future__ import annotations

import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parent

import stock_server as S
from download_feeds import NIFTY50_FALLBACK

S.SYMBOLS = sorted(set(NIFTY50_FALLBACK))

EVENTS = [("RESULTS", "Quarterly Results"), ("BOARD_MEETING", "Board Meeting"),
          ("CORPORATE_ACTION", "Corporate Action"), ("ANNOUNCEMENT", "Announcements")]


def verdict(win, avg):
    if avg <= 0:
        return ("avoid", "Avoid")
    if win >= 60 and avg >= 0.5:
        return ("strong", "Strong")
    if win >= 55:
        return ("ok", "Some edge")
    return ("weak", "Weak")


def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


print("Building rankings for the Nifty 50 (analysing every stock — a few seconds)…")
data = S.build_rankings()
ranks = data["ranks"]

# invert -> per-stock map for the master matrix
by_stock: dict[str, dict] = {}
for et, _ in EVENTS:
    for r in ranks.get(et, []):
        by_stock.setdefault(r["symbol"], {})[et] = r
stocks = sorted(by_stock)

# ── per-event ranked tables ─────────────────────────────────────────────────
def event_table(et):
    rows = []
    for i, r in enumerate(ranks.get(et, []), 1):
        vc, vt = verdict(r["win"], r["avg"])
        medal = {1: "🥇", 2: "🥈", 3: "🥉"}.get(i, str(i))
        wcol = "pos" if r["win"] >= 55 else ("amb" if r["win"] >= 50 else "neg")
        acol = "pos" if r["avg"] >= 0 else "neg"
        rows.append(
            f'<tr><td class="rk">{medal}</td><td class="sym">{esc(r["symbol"])}</td>'
            f'<td class="mono">buy <b>{r["db"]}d</b> before</td>'
            f'<td class="mono">sell <b>{r["da"]}d</b> after</td>'
            f'<td class="num {wcol}"><b>{r["win"]:.0f}%</b></td>'
            f'<td class="num {acol}">{"+" if r["avg"]>=0 else ""}{r["avg"]:.2f}%</td>'
            f'<td class="num">{r["events"]}</td>'
            f'<td><span class="pill {vc}">{vt}</span></td></tr>')
    return "".join(rows)


# ── master matrix: one row per stock, all events ────────────────────────────
def matrix_rows():
    rows = []
    for sym in stocks:
        cells = [f'<td class="sym">{esc(sym)}</td>']
        for et, _ in EVENTS:
            r = by_stock[sym].get(et)
            if not r:
                cells.append('<td class="mono muted">—</td>')
                continue
            wcol = "pos" if r["win"] >= 55 else ("amb" if r["win"] >= 50 else "neg")
            cells.append(
                f'<td class="cell"><span class="rule">buy {r["db"]}d · sell {r["da"]}d</span>'
                f'<span class="num {wcol}">{r["win"]:.0f}% · {"+" if r["avg"]>=0 else ""}{r["avg"]:.1f}%</span></td>')
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return "".join(rows)


CSS = """
:root{--bg:#f7f8fa;--panel:#fff;--ink:#12161d;--mute:#5b6472;--line:#e4e7ec;--blue:#2563eb;
  --pos:#15803d;--neg:#b91c1c;--amb:#b45309;--shadow:0 1px 2px rgba(16,24,40,.06),0 1px 3px rgba(16,24,40,.1);
  --mono:'SFMono-Regular',Consolas,Menlo,monospace;}
@media (prefers-color-scheme:dark){:root{--bg:#0c0f14;--panel:#141922;--ink:#e6e9ef;--mute:#98a2b3;
  --line:#232a35;--blue:#5b9bff;--pos:#5ec98a;--neg:#f87171;--amb:#fbbf5b;--shadow:none;}}
:root[data-theme="dark"]{--bg:#0c0f14;--panel:#141922;--ink:#e6e9ef;--mute:#98a2b3;--line:#232a35;
  --blue:#5b9bff;--pos:#5ec98a;--neg:#f87171;--amb:#fbbf5b;--shadow:none;}
:root[data-theme="light"]{--bg:#f7f8fa;--panel:#fff;--ink:#12161d;--mute:#5b6472;--line:#e4e7ec;
  --blue:#2563eb;--pos:#15803d;--neg:#b91c1c;--amb:#b45309;--shadow:0 1px 2px rgba(16,24,40,.06),0 1px 3px rgba(16,24,40,.1);}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);line-height:1.6;font-size:15px;
  font-family:'Inter',-apple-system,'Segoe UI',Roboto,Helvetica,Arial,sans-serif}
.wrap{max-width:1040px;margin:0 auto;padding:0 20px 90px}
header{padding:48px 0 26px;border-bottom:1px solid var(--line)}
.kicker{font-family:var(--mono);font-size:.78rem;letter-spacing:.08em;text-transform:uppercase;color:var(--blue);font-weight:600}
h1{font-size:1.95rem;margin:.3em 0 .2em;letter-spacing:-.02em}
.lede{font-size:1.05rem;color:var(--mute);margin:0;max-width:70ch}
h2{font-size:1.28rem;margin:44px 0 4px}
h2 .n{color:var(--blue);font-family:var(--mono);font-size:.95rem;margin-right:.5em}
p{margin:.55em 0}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px 20px;margin:14px 0;box-shadow:var(--shadow)}
.how{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:14px 0}
@media(max-width:640px){.how{grid-template-columns:1fr}}
.how .c{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px 16px;box-shadow:var(--shadow)}
.how .c b{color:var(--ink)}
.callout{border-left:4px solid var(--amb);background:var(--panel);border-radius:0 10px 10px 0;padding:13px 18px;margin:14px 0;box-shadow:var(--shadow);color:var(--mute)}
.tblscroll{overflow-x:auto;border:1px solid var(--line);border-radius:12px;margin:14px 0;box-shadow:var(--shadow)}
table{width:100%;border-collapse:collapse;font-size:.86rem;background:var(--panel)}
th,td{padding:8px 11px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap}
th{font-size:.72rem;text-transform:uppercase;letter-spacing:.04em;color:var(--mute);font-weight:600;
  position:sticky;top:0;background:var(--panel)}
.num{text-align:right;font-family:var(--mono)}
.mono{font-family:var(--mono);font-size:.82rem}
.rk{text-align:center;font-family:var(--mono);width:40px}
.sym{font-weight:700;color:var(--ink)}
.pos{color:var(--pos)}.neg{color:var(--neg)}.amb{color:var(--amb)}.muted{color:var(--mute)}
tr:hover td{background:rgba(37,99,235,.05)}
.pill{font-family:var(--mono);font-size:.7rem;font-weight:600;padding:.14em .6em;border-radius:20px}
.pill.strong{background:rgba(21,128,61,.14);color:var(--pos)}
.pill.ok{background:rgba(180,83,9,.15);color:var(--amb)}
.pill.weak{background:rgba(139,152,173,.16);color:var(--mute)}
.pill.avoid{background:rgba(185,28,28,.14);color:var(--neg)}
.cell .rule{display:block;font-family:var(--mono);font-size:.78rem;color:var(--ink)}
.cell .num{display:block;font-size:.78rem;margin-top:2px}
.foot{color:var(--mute);font-size:.82rem;margin-top:34px;font-family:var(--mono)}
@media print{:root{--bg:#fff;--panel:#fff;--ink:#12161d;--mute:#4b5563;--line:#d7dbe0;--shadow:none;}
  @page{size:A4 landscape;margin:12mm}html,body{background:#fff!important;font-size:9.5pt}
  .wrap{max-width:100%;padding:0}h2{break-after:avoid}.tblscroll{break-inside:auto}tr{break-inside:avoid}}
"""

sections = "".join(
    f'<h2><span class="n">·</span>{esc(lab)}</h2>'
    f'<div class="tblscroll"><table><thead><tr>'
    f'<th>#</th><th>Stock</th><th>Buy signal</th><th>Sell signal</th>'
    f'<th class="num">Win rate</th><th class="num">Avg gain</th><th class="num">Events</th><th>Signal</th>'
    f'</tr></thead><tbody>{event_table(et)}</tbody></table></div>'
    for et, lab in EVENTS)

matrix_head = "".join(f'<th>{esc(lab)}</th>' for _, lab in EVENTS)
n = data["n_stocks"]

HTML = f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Buy-Before / Sell-After Playbook — Nifty 50</title><style>{CSS}</style></head><body><div class="wrap">
<header>
  <div class="kicker">NSE Behaviour Analysis · Playbook</div>
  <h1>Buy-before / sell-after — all {n} tracked stocks</h1>
  <p class="lede">For each stock and each corporate event, the historical “buy N days before, sell M days
    after” rule with the best win rate — plus how you actually know the event date in advance, and the
    honest caveat on what these numbers do and don’t mean. Universe = the Nifty 50 plus a couple of
    recent ex-constituents kept for their event history ({n} in all).</p>
</header>

<h2><span class="n">01</span>How to “predict” the event — and read this report</h2>
<div class="how">
  <div class="c"><b>You don’t guess the date — you read it.</b> For <b>scheduled</b> events (results, board
    meetings, dividends) SEBI Regulation 29 forces the company to intimate the exchange ~5 trading days ahead.
    That intimation <em>is</em> your buy signal window.</div>
  <div class="c"><b>Buy-before</b> harvests the run-up as the market <em>anticipates</em> a known event;
    <b>sell-after</b> harvests the drift as the market digests the news. Announcements (press releases,
    acquisitions) are <b>unscheduled</b> — you can only act <em>after</em> they’re public.</div>
</div>
<div class="callout"><b>Read this honestly.</b> These win rates are <b>in-sample</b> — the rule was picked on
  the same history it’s scored against, so they flatter. A high past win rate does not guarantee future
  profit, especially where the event count is small. Treat this as a map of past behaviour, not a signal
  to trade blindly. See the concept note for why.</div>

<h2><span class="n">02</span>Master table — all stocks, all events</h2>
<p class="lede" style="font-size:.92rem">Each cell: the best rule and its <b>win% · avg gain</b>. Sorted A–Z.</p>
<div class="tblscroll"><table><thead><tr><th>Stock</th>{matrix_head}</tr></thead>
<tbody>{matrix_rows()}</tbody></table></div>

{sections}

<p class="foot">Generated from the same engine as the dashboard’s 🏆 Rankings tab · {data['n_stocks']} stocks ·
win rate = share of past events where the best rule made money · numbers are in-sample.</p>
</div></body></html>"""

out = ROOT / "buy_sell_playbook.html"
out.write_text(HTML, encoding="utf-8")
print(f"wrote {out}  ({len(HTML)//1024} KB, {len(stocks)} stocks)")
