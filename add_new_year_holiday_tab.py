"""Add a 'New Year 2027' event (hrow-newyear_2027) to the Stock Holidays tab.

Usage: python add_new_year_holiday_tab.py <ROOT> <POSITION_BOOK_XLSX>
Idempotent, additive only: inserts a new holiday entry into
dashboard_data/event_dashboard_data.json and anchors-inserts a master-table row
plus the id into the three hard-coded id lists of event_trading_dashboard.html.
Then run build_dashboard.py in ROOT to regenerate the .js bundle and docs/.
"""
import sys, json, os, pandas as pd

ROOT, PB = sys.argv[1], sys.argv[2]
JP = os.path.join(ROOT, "dashboard_data", "event_dashboard_data.json")
HP = os.path.join(ROOT, "event_trading_dashboard.html")

d = json.load(open(JP, encoding="utf-8"))
fo = {x["symbol"]: x for x in json.load(open(os.path.join(ROOT, "dashboard_data", "fo_stocks_211.json"), encoding="utf-8"))}
other = {}
for h in d["holidays"]:
    for s in h["stocks"]:
        other.setdefault(s["symbol"], s)

ST = pd.read_csv(os.path.join(os.path.dirname(os.path.abspath(PB)), "New_Year_live_15yr_stats.csv")).set_index("Symbol")
pb = pd.read_excel(PB, sheet_name="Position_Book", header=3).dropna(subset=["Symbol"])
pb = pb.sort_values("Rank")


def fmt(ts):
    t = pd.Timestamp(ts)
    return t.strftime("%d-%b-%Y (%A)")


stocks = []
for _, r in pb.iterrows():
    sym = r["Symbol"]
    n = int(r["Enter N before T"])
    hold = int(r["Hold days"])
    after = hold - n
    win = f"T-{n} to T" + (f"+{after}" if after > 0 else (f"{after}" if after < 0 else ""))
    f = fo.get(sym, {})
    o = other.get(sym, {})
    direction = r["Direction"]
    lot = int(f.get("lot_size") or o.get("lot") or 0)
    margin = float(f.get("margin_20pct") or o.get("margin") or 0)
    st = ST.loc[sym]
    wr = round(float(st.WR15), 1)          # 15-yr (2012-2026) win rate, NET of 0.10% cost, at the live window
    avg = round(float(st.AVG15), 2)
    oos_n = int(st.OOS_n); oos_wr = None if pd.isna(st.OOS_wr) else round(float(st.OOS_wr), 1)
    t = float(r["Hist t"])
    conv = "STRONG" if t >= 3 else ("MODERATE" if t >= 2 else "WEAK")
    mon = "JAN" if pd.Timestamp(r["Exit Date"]).month == 1 else "DEC"
    stocks.append({
        "symbol": sym,
        "name": f.get("name") or o.get("name") or sym,
        "direction": direction,
        "window": win,
        "lead": n,
        "hold": after,
        "entry_date": fmt(r["Entry Date"]),
        "exit_date": fmt(r["Exit Date"]),
        "total_years": int(st.N15), "hist_years": 15,
        "full_19y_wr": wr,
        "full_19y_avg_ret": avg,
        "last_4y_wr": None,
        "last_4y_avg_ret": round(float(r["Last5 Avg pct"]), 2),
        "lot": lot,
        "margin": margin,
        "fut_action": f"{'BUY' if direction == 'LONG' else 'SELL'} FUTURES ({direction} {lot} qty)",
        "opt_action": f"BUY {mon} {'CALL' if direction == 'LONG' else 'PUT'} OPTION",
        "sector": o.get("sector") or "Unclassified",
        "direction_basis": (f"New Year 1-Jan-2027: best window {win}. Last 15 yrs (2012-2026), net of 0.10% cost: win {wr:.0f}% over {int(st.N15)} yrs, avg {avg:+.2f}% (in-sample for this window). "
                            f"Genuine out-of-sample (walk-forward, window chosen only from earlier years): "
                            + (f"{oos_wr:.0f}% win over {oos_n} traded yrs." if oos_wr is not None else "not selected in any test year.")),
        "n_years": int(st.N15),
        "conviction": conv,
        "wf_oos_win_rate": oos_wr,
        "wf_oos_trades": oos_n,
        "confidence_flag": "OK",
        "confidence_note": "",
    })

nl = sum(s["direction"] == "LONG" for s in stocks)
ns = len(stocks) - nl
top = sorted(stocks, key=lambda s: -s["full_19y_avg_ret"])[:2]
entry = {
    "id": "newyear_2027",
    "name": "New Year (1-Jan)",
    "date": "01-Jan-2027 (Friday)",
    "category": "Year-End Event",
    "status": "UPCOMING (Dec 2026 - Jan 2027)",
    "badge": "UPCOMING",
    "bias": (f"New Year 1-Jan-2027 -- top-50 F&O, per-stock best window, long/short on a 15-yr walk-forward: "
             f"{len(stocks)} positions ({nl} LONG / {ns} SHORT). Book backtest +1.83%/yr, 14 of 15 years positive, "
             f"max drawdown -0.29%. T = Fri 1-Jan-2027 (25-Dec holiday)."),
    "hist_years": 15,
    "stocks_count": len(stocks),
    "top_pick": " & ".join(f"{s['symbol']} ({s['full_19y_avg_ret']:+.2f}% avg)" for s in top),
    "stocks": stocks,
}
ids = [h["id"] for h in d["holidays"]]
if "newyear_2027" in ids:
    d["holidays"][ids.index("newyear_2027")] = entry
else:
    d["holidays"].append(entry)
json.dump(d, open(JP, "w", encoding="utf-8", newline="\n"), ensure_ascii=False, indent=2)
print("json ok", len(stocks), nl, ns)

# ---------------- html ----------------
h = open(HP, encoding="utf-8", newline="").read()
crlf = "\r\n" in h
h = h.replace("\r\n", "\n")

ROW = f"""              <tr class="event-master-row" id="hrow-newyear_2027" onclick="selectHolidayRow('newyear_2027', this)">
                <td><strong>New Year (1-Jan)</strong></td>
                <td><span class="date-badge">01-Jan-2027 (Friday)</span></td>
                <td>Year-End Event</td>
                <td><span style="color:#94A3B8; font-size:0.75rem;">UPCOMING IN JAN</span></td>
                <td><strong style="color:#38BDF8;">Tuesday, 22-Dec (T-7)</strong></td>
                <td><span class="badge-long">{nl} LONG</span> / <span class="badge-short">{ns} SHORT</span> <span style="font-size:0.62rem; color:#94A3B8;">(15-yr walk-forward)</span></td>
                <td><button class="btn-select-event" onclick="event.stopPropagation(); selectHolidayRow('newyear_2027', this.closest('tr'))">View Playbook</button></td>
              </tr>
"""
if 'id="hrow-newyear_2027"' not in h:
    a = h.index('id="hrow-christmas_2026"')
    end = h.index("</tr>\n", a) + len("</tr>\n")
    h = h[:end] + ROW + h[end:]

old = "['gandhi_2026', 'dussehra_2026', 'diwali_2026', 'gurunanak_2026', 'christmas_2026'].forEach"
new = "['gandhi_2026', 'dussehra_2026', 'diwali_2026', 'gurunanak_2026', 'christmas_2026', 'newyear_2027'].forEach"
if old in h:
    h = h.replace(old, new)
oldr = "{ id: 'christmas_2026', closureDate: '25-Dec-2026', entryDate: '15-Dec-2026', name: 'Christmas & Year-End' }\n"
newr = ("{ id: 'christmas_2026', closureDate: '25-Dec-2026', entryDate: '15-Dec-2026', name: 'Christmas & Year-End' },\n"
        "        { id: 'newyear_2027', closureDate: '01-Jan-2027', entryDate: '22-Dec-2026', name: 'New Year (1-Jan)' }\n")
if oldr in h and "id: 'newyear_2027'" not in h:
    h = h.replace(oldr, newr)

def rep(a, b):
    global h
    if a in h and b not in h:
        h = h.replace(a, b, 1)
    elif a not in h and b not in h:
        print("WARN anchor missing:", a[:60])


rep("<th>19Y Win Rate</th>", '<th id="h-th-wr">19Y Win Rate</th>')
rep("<th>19Y Avg Ret</th>", '<th id="h-th-ret">19Y Avg Ret</th>')
rep("      updateHolidayFilterPills(h);\n    }\n\n    let currentRbiCatFilter",
    "      const _thw = document.getElementById('h-th-wr'), _thr = document.getElementById('h-th-ret');\n"
    "      if (_thw) _thw.textContent = (h.hist_years || 19) + 'Y Win Rate' + (h.hist_years ? ' (net)' : '');\n"
    "      if (_thr) _thr.textContent = (h.hist_years || 19) + 'Y Avg Ret';\n"
    "      updateHolidayFilterPills(h);\n    }\n\n    let currentRbiCatFilter")
rep("""<td><strong style="color:${wrVal >= 80 ? '#FBBF24' : '#E2E8F0'}">${wrVal.toFixed(1)}%</strong></td>""",
    """<td><strong style="color:${wrVal >= 80 ? '#FBBF24' : '#E2E8F0'}">${wrVal.toFixed(1)}%</strong>${(h.hist_years && s.wf_oos_win_rate !== null && s.wf_oos_win_rate !== undefined) ? `<br><small style="color:#94A3B8;" title="Walk-forward out-of-sample: window chosen only from earlier years">OOS ${Number(s.wf_oos_win_rate).toFixed(0)}% (${s.wf_oos_trades}y)</small>` : ''}</td>""")
rep("""<strong>📈 19-Year Model Win Rate:</strong> <strong style="color:#FBBF24;">${stock.full_19y_wr}%</strong> (${stock.total_years || 19} Years Backtested)""",
    """<strong>📈 ${stock.hist_years || 19}-Year ${stock.hist_years ? 'Net Win Rate' : 'Model Win Rate'}:</strong> <strong style="color:#FBBF24;">${stock.full_19y_wr}%</strong> (${stock.total_years || 19} Years Backtested)${(stock.hist_years && stock.wf_oos_win_rate !== null && stock.wf_oos_win_rate !== undefined) ? ` &bull; genuine out-of-sample ${stock.wf_oos_win_rate}% (${stock.wf_oos_trades} yrs)` : ''}""")
rep("""<strong>🎯 19-Year Avg Return %:</strong>""", """<strong>🎯 ${stock.hist_years || 19}-Year Avg Return %:</strong>""")
rep("""if (subheading) subheading.innerHTML = `Empirically Derived""",
    """if (subheading) subheading.innerHTML = hObj.hist_years ? `Last ${hObj.hist_years} years (2012&ndash;2026) of daily bars &bull; win rate is NET of 0.10% round-trip cost; "OOS" = genuine walk-forward out-of-sample win rate (window chosen only from earlier years). Book backtest +1.83%/yr, 14 of 15 years positive.` : `Empirically Derived""")

if crlf:
    h = h.replace("\n", "\r\n")
open(HP, "w", encoding="utf-8", newline="").write(h)
print("html ok", 'hrow-newyear_2027' in h, "newyear_2027'].forEach" in h, "id: 'newyear_2027'" in h)
