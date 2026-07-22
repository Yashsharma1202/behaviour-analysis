"""
dashboard_gui.py
"""
import tkinter as tk
from pathlib import Path
from tkinter import ttk

import pandas as pd

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"

# --- "trading terminal" dark palette ---------------------------------------
BG = "#0f131b"        # window background
PANEL = "#171d2b"     # card / chart background
GRID = "#2a3346"      # gridlines
INK = "#e6edf3"       # primary text
MUTE = "#8b98ad"      # secondary text
POS = "#3fb950"       # green  (up / beat)
NEG = "#f85149"       # red    (down / miss)
BLUE = "#58a6ff"
AMBER = "#d29922"
PURPLE = "#bc8cff"
SERIES = [BLUE, AMBER, PURPLE, POS, NEG]


# ===========================================================================
# Tiny Canvas charting toolkit (no external libs)
# ===========================================================================
class Chart(tk.Canvas):
    """A canvas that knows how to draw simple axes, line series and bars."""

    def __init__(self, master, width=760, height=340, **kw):
        super().__init__(master, width=width, height=height, bg=PANEL,
                         highlightthickness=0, **kw)
        self.W, self.H = width, height
        self.ml, self.mr, self.mt, self.mb = 64, 24, 44, 40   # margins

    # inner plotting box in pixels
    def _box(self):
        return self.ml, self.mt, self.W - self.mr, self.H - self.mb

    def _sx(self, v, lo, hi):
        x0, _, x1, _ = self._box()
        return x0 + (v - lo) / ((hi - lo) or 1) * (x1 - x0)

    def _sy(self, v, lo, hi):
        _, y0, _, y1 = self._box()
        return y1 - (v - lo) / ((hi - lo) or 1) * (y1 - y0)   # inverted

    def title(self, text, sub=""):
        self.create_text(self.ml, 14, text=text, fill=INK, anchor="w",
                         font=("Segoe UI Semibold", 12))
        if sub:
            self.create_text(self.W - self.mr, 14, text=sub, fill=MUTE, anchor="e",
                             font=("Segoe UI", 9))

    def _yaxis(self, lo, hi, fmt="{:.1f}", ticks=5):
        x0, y0, x1, y1 = self._box()
        for i in range(ticks + 1):
            v = lo + (hi - lo) * i / ticks
            y = self._sy(v, lo, hi)
            self.create_line(x0, y, x1, y, fill=GRID)
            self.create_text(x0 - 8, y, text=fmt.format(v), fill=MUTE, anchor="e",
                             font=("Consolas", 8))

    def line_series(self, x, series: dict, lo=None, hi=None, xlabels=None,
                    zero=True, yfmt="{:+.1f}"):
        """x: list of x positions. series: {name: [y...]} mapped to colours."""
        allv = [v for ys in series.values() for v in ys]
        lo = min(allv + ([0] if zero else [])) if lo is None else lo
        hi = max(allv + ([0] if zero else [])) if hi is None else hi
        pad = (hi - lo) * 0.12 or 1
        lo, hi = lo - pad, hi + pad
        self._yaxis(lo, hi, yfmt)
        xlo, xhi = min(x), max(x)
        if zero and lo <= 0 <= hi:                       # bold zero line
            y = self._sy(0, lo, hi)
            self.create_line(self.ml, y, self.W - self.mr, y, fill=MUTE, width=1)
        # x tick labels
        for i, xv in enumerate(x):
            lab = xlabels[i] if xlabels else str(xv)
            self.create_text(self._sx(xv, xlo, xhi), self.H - self.mb + 14,
                             text=lab, fill=MUTE, font=("Consolas", 8))
        # each series
        for idx, (name, ys) in enumerate(series.items()):
            col = SERIES[idx % len(SERIES)]
            pts = []
            for xv, yv in zip(x, ys):
                pts += [self._sx(xv, xlo, xhi), self._sy(yv, lo, hi)]
            if len(pts) >= 4:
                self.create_line(*pts, fill=col, width=2, smooth=False)
            for xv, yv in zip(x, ys):
                px, py = self._sx(xv, xlo, xhi), self._sy(yv, lo, hi)
                self.create_oval(px - 3, py - 3, px + 3, py + 3, fill=col, outline=PANEL)
            # legend
            ly = self.mt + 4 + idx * 16
            self.create_line(self.W - self.mr - 120, ly, self.W - self.mr - 100, ly,
                             fill=col, width=3)
            self.create_text(self.W - self.mr - 94, ly, text=name, fill=INK,
                             anchor="w", font=("Segoe UI", 9))

    def bars(self, cats, values, colors=None, yfmt="{:.0f}", value_tags=True):
        vals = [0 if pd.isna(v) else v for v in values]
        lo, hi = min(vals + [0]), max(vals + [0])
        pad = (hi - lo) * 0.15 or 1
        lo, hi = lo - (pad if lo < 0 else 0), hi + pad
        self._yaxis(lo, hi, yfmt)
        if lo <= 0 <= hi:
            yz = self._sy(0, lo, hi)
            self.create_line(self.ml, yz, self.W - self.mr, yz, fill=MUTE)
        else:
            yz = self._sy(lo, lo, hi)
        x0, _, x1, _ = self._box()
        n = len(cats)
        slot = (x1 - x0) / (n or 1)
        bw = slot * 0.6
        for i, (c, v) in enumerate(zip(cats, vals)):
            cx = x0 + slot * (i + 0.5)
            col = (colors[i] if colors else BLUE)
            y = self._sy(v, lo, hi)
            self.create_rectangle(cx - bw / 2, yz, cx + bw / 2, y, fill=col, outline="")
            self.create_text(cx, self.H - self.mb + 14, text=str(c), fill=MUTE,
                             font=("Consolas", 8))
            if value_tags:
                self.create_text(cx, y - 8 if v >= 0 else y + 8,
                                 text=yfmt.format(v), fill=INK, font=("Consolas", 8))


# ===========================================================================
# Helpers to load each result file (fail soft with a friendly message)
# ===========================================================================
def _read(name):
    p = PROC / name
    return pd.read_csv(p) if p.exists() else None


def car_from_study(fname):
    df = _read(fname)
    if df is None or df.empty:
        return None
    cols = [c for c in df.columns if c.lstrip("-").isdigit()]
    mean = df[cols].mean()
    days = [int(c) for c in cols]
    order = sorted(range(len(days)), key=lambda i: days[i])
    days = [days[i] for i in order]
    car = (mean.iloc[order].cumsum() * 100).tolist()
    return days, car


# ===========================================================================
# The application
# ===========================================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RELIANCE  –  Corporate-Event Behaviour Dashboard")
        self.configure(bg=BG)
        self.geometry("860x640")

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TNotebook", background=BG, borderwidth=0)
        style.configure("TNotebook.Tab", background=PANEL, foreground=MUTE,
                        padding=(16, 8), font=("Segoe UI Semibold", 10))
        style.map("TNotebook.Tab", background=[("selected", BG)],
                  foreground=[("selected", BLUE)])
        style.configure("Treeview", background=PANEL, foreground=INK,
                        fieldbackground=PANEL, rowheight=22, borderwidth=0)
        style.configure("Treeview.Heading", background=BG, foreground=MUTE,
                        font=("Segoe UI Semibold", 9))
        style.map("Treeview", background=[("selected", "#243044")],
                  foreground=[("selected", INK)])

        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(header, text="Corporate-Event Behaviour Analysis",
                 bg=BG, fg=INK, font=("Segoe UI Semibold", 16)).pack(side="left")
        tk.Label(header, text="RELIANCE INDUSTRIES  ·  NSE  ·  2007–2026",
                 bg=BG, fg=MUTE, font=("Segoe UI", 10)).pack(side="left", padx=12)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=8)
        self.nb = nb

        self._tab_overview()
        self._tab_records()
        self._tab_event_study()
        self._tab_recovery()
        self._tab_surprise()
        self._tab_backtest()
        self._tab_fundamentals()

    # -- reusable tab scaffold ------------------------------------------
    def _panel(self, title):
        frame = tk.Frame(self.nb, bg=BG)
        self.nb.add(frame, text=title)
        return frame

    def _card(self, parent, title, value, sub, color=INK):
        c = tk.Frame(parent, bg=PANEL, padx=16, pady=12)
        tk.Label(c, text=title, bg=PANEL, fg=MUTE,
                 font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(c, text=value, bg=PANEL, fg=color,
                 font=("Segoe UI Semibold", 22)).pack(anchor="w")
        tk.Label(c, text=sub, bg=PANEL, fg=MUTE,
                 font=("Segoe UI", 8)).pack(anchor="w")
        return c

    # -- OVERVIEW -------------------------------------------------------
    def _tab_overview(self):
        f = self._panel("Overview")
        uni = _read("unified_events.csv")
        rec = _read("dividend_recovery.csv")
        sur = _read("earnings_surprise.csv")

        # KPI cards
        cards = tk.Frame(f, bg=BG)
        cards.pack(fill="x", padx=8, pady=8)
        n_events = len(uni) if uni is not None else 0
        n_div = int((uni["event_type"] == "DIVIDEND").sum()) if uni is not None else 0
        med_rec = (f'{rec["recovery_days"].median():.0f} d'
                   if rec is not None and rec["recovery_days"].notna().any() else "n/a")
        if sur is not None and not sur.empty:
            beat = sur[sur["label"] == "BEAT"]["post5_ar_%"].mean()
            miss = sur[sur["label"] == "MISS"]["post5_ar_%"].mean()
            spread = f"{beat - miss:+.1f}%"
        else:
            spread = "n/a"
        for i, (t, v, s, col) in enumerate([
            ("Dated events", f"{n_events:,}", "across 4 NSE feeds", INK),
            ("Dividends", str(n_div), "ex-dates analysed", BLUE),
            ("Median div. recovery", med_rec, "trading days to regain price", POS),
            ("Beat−Miss drift", spread, "5-day post-results spread", AMBER),
        ]):
            card = self._card(cards, t, v, s, col)
            card.grid(row=0, column=i, padx=6, sticky="nsew")
            cards.columnconfigure(i, weight=1)

        # event-type bar chart
        ch = Chart(f, width=800, height=360)
        ch.pack(padx=8, pady=8)
        ch.title("Event mix in the master timeline", "top categories by count")
        if uni is not None:
            vc = uni["event_type"].value_counts().head(8)
            labels = [s[:10] for s in vc.index]
            ch.bars(labels, vc.values.tolist(), yfmt="{:.0f}")

    # -- RECORDS (detailed event history) -------------------------------
    def _tab_records(self):
        f = self._panel("Records")
        rec = _read("event_records.csv")
        tk.Label(f, text="Reliance event history — click any row to see what happened",
                 bg=BG, fg=INK, font=("Segoe UI Semibold", 11)).pack(
                     anchor="w", padx=12, pady=(8, 4))

        body = tk.Frame(f, bg=BG)
        body.pack(fill="both", expand=True, padx=12)
        cols = ("date", "type", "event", "before", "adj", "after")
        heads = {"date": ("Date", 90), "type": ("Type", 85),
                 "event": ("Event", 240), "before": ("Before 5d", 78),
                 "adj": ("Ex-date adj", 92), "after": ("After 5d", 78)}
        tv = ttk.Treeview(body, columns=cols, show="headings", height=13)
        for c, (t, w) in heads.items():
            tv.heading(c, text=t)
            tv.column(c, width=w, anchor="w" if c in ("event", "type") else "center")
        sb = ttk.Scrollbar(body, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        tv.tag_configure("up", foreground=POS)
        tv.tag_configure("down", foreground=NEG)
        tv.tag_configure("flat", foreground=INK)

        self._rec_rows = {}
        if rec is not None:
            for _, r in rec.iterrows():
                after = r["after_%"]
                tag = "up" if after > 0.5 else ("down" if after < -0.5 else "flat")
                adj = f"{r['adj_%']:+.1f}%" if pd.notna(r["adj_%"]) else "—"
                vals = (pd.to_datetime(r["date"]).strftime("%d-%b-%Y"), r["category"],
                        str(r["event"])[:40], f"{r['before_%']:+.1f}%", adj,
                        f"{after:+.1f}%")
                iid = tv.insert("", "end", values=vals, tags=(tag,))
                self._rec_rows[iid] = r

        detail = tk.Label(f, text="Select an event above for a plain-English "
                          "explanation of what happened and the price adjustment.",
                          bg=PANEL, fg=INK, font=("Segoe UI", 10), justify="left",
                          wraplength=830, anchor="w", padx=14, pady=12)
        detail.pack(fill="x", padx=12, pady=10)

        def explain(_=None):
            sel = tv.selection()
            if sel:
                detail.config(text=self._explain_record(self._rec_rows[sel[0]]))
        tv.bind("<<TreeviewSelect>>", explain)

    def _explain_record(self, r) -> str:
        d = pd.to_datetime(r["date"]).strftime("%d-%b-%Y")
        cat, before, after, adj = r["category"], r["before_%"], r["after_%"], r["adj_%"]
        beh = (f"In the 5 trading days BEFORE it, the stock moved {before:+.1f}%; "
               f"in the 5 days AFTER, {after:+.1f}%.")
        if cat == "DIVIDEND":
            return (f"{d}  —  EX-DIVIDEND ({r['detail']}).  On the ex-date the price was "
                    f"adjusted {adj:+.1f}% — roughly the dividend leaving the share (you "
                    f"keep the cash, so it's not a real loss).  {beh}  For Reliance the "
                    f"dividend is tiny versus the price, so this 'discount' is barely "
                    f"visible and usually recovered within a day.")
        if cat in ("BONUS", "SPLIT"):
            return (f"{d}  —  EX-{cat} ({r['detail']}).  The price is mechanically "
                    f"adjusted {adj:+.0f}% because you receive extra shares — a 1:1 bonus "
                    f"doubles your share count and halves the price, so your total value "
                    f"is unchanged.  {beh}")
        if cat == "EARNINGS":
            extra = f"  Result: {r['detail']}." if r["detail"] else ""
            drift = ("kept rising — a positive reaction" if after > 0.5 else
                     "fell back — classic 'sell the news'" if after < -0.5 else "was flat")
            return (f"{d}  —  QUARTERLY RESULTS.{extra}  {beh}  Afterwards the stock "
                    f"{drift}.  Reliance tends to drift UP into results then fade after, "
                    f"so historically the edge is being positioned BEFORE the announcement.")
        if cat == "DEMERGER":
            return (f"{d}  —  DEMERGER.  Reliance spun off a business (Jio Financial "
                    f"Services) to shareholders as a separate listed company.  {beh}  "
                    f"A demerger carves value out of the parent into the new entity.")
        if cat == "RIGHTS":
            return (f"{d}  —  RIGHTS ISSUE ({r['detail']}).  Shareholders got the right to "
                    f"buy new shares at a discount; the price adjusts for the extra shares. "
                    f"  {beh}")
        return f"{d}  —  {cat}: {r['event']}.  {beh}"

    # -- EVENT STUDY ----------------------------------------------------
    def _tab_event_study(self):
        f = self._panel("Event Study")
        ch = Chart(f, width=800, height=440)
        ch.pack(padx=8, pady=12)
        ch.title("Cumulative abnormal return around events (vs Nifty 50)",
                 "market-adjusted, day 0 = event")
        series = {}
        days_ref = None
        for name, fn in [("Dividends", "event_study_dividends.csv"),
                         ("Earnings", "event_study_earnings.csv"),
                         ("Acquisitions", "event_study_acquisitions.csv")]:
            r = car_from_study(fn)
            if r:
                days_ref, car = r
                series[name] = car
        if series:
            xl = [f"t{d:+d}" for d in days_ref]
            ch.line_series(days_ref, series, xlabels=xl, yfmt="{:+.1f}%")
        tk.Label(f, text="Bars/points above the zero line = stock beat the market. "
                 "A rise before t0 hints at anticipation; drift after t0 = ongoing re-pricing.",
                 bg=BG, fg=MUTE, font=("Segoe UI", 9), wraplength=800,
                 justify="left").pack(padx=12, anchor="w")

    # -- DIVIDEND RECOVERY ----------------------------------------------
    def _tab_recovery(self):
        f = self._panel("Dividend Recovery")
        rec = _read("dividend_recovery.csv")
        top = tk.Frame(f, bg=BG)
        top.pack(fill="x", padx=8, pady=8)

        ch1 = Chart(top, width=400, height=300)
        ch1.pack(side="left", padx=4)
        ch1.title("Dividend yield trend", "dividend as % of price")
        ch2 = Chart(top, width=400, height=300)
        ch2.pack(side="left", padx=4)
        ch2.title("Recovery time", "trading days to regain pre-div price")

        if rec is not None and not rec.empty:
            yrs = [str(pd.to_datetime(d).year) for d in rec["ex_date"]]
            ch1.line_series(list(range(len(rec))), {"yield %": rec["div_yield_%"].tolist()},
                            xlabels=yrs, zero=True, yfmt="{:.1f}%")
            days = rec["recovery_days"].fillna(120).tolist()
            cols = [POS if d <= 5 else (AMBER if d <= 30 else NEG) for d in days]
            ch2.bars(yrs, days, colors=cols, yfmt="{:.0f}")

        tk.Label(f, text="Reliance's dividend yield fell from ~3.7% (2011) to ~0.4% (2025) "
                 "as the price grew ~6×. The ex-date 'discount' shrank to less than one "
                 "day's normal move → most dividends are recovered the same day. Green bars "
                 "= recovered ≤5 days.",
                 bg=BG, fg=MUTE, font=("Segoe UI", 9), wraplength=800,
                 justify="left").pack(padx=12, anchor="w", pady=6)

    # -- EARNINGS SURPRISE ----------------------------------------------
    def _tab_surprise(self):
        f = self._panel("Earnings Surprise")
        sur = _read("earnings_surprise.csv")
        ch = Chart(f, width=800, height=380)
        ch.pack(padx=8, pady=10)
        ch.title("Post-results drift: BEAT vs MISS quarters",
                 "5-day market-adjusted return by YoY profit direction")
        if sur is not None and not sur.empty:
            cats, vals, cols = [], [], []
            for _, r in sur.sort_values("event_date").iterrows():
                cats.append(pd.to_datetime(r["event_date"]).strftime("%b-%y"))
                vals.append(r["post5_ar_%"])
                cols.append(POS if r["label"] == "BEAT" else NEG)
            ch.bars(cats, vals, colors=cols, yfmt="{:+.1f}%")
            beat = sur[sur["label"] == "BEAT"]["post5_ar_%"].mean()
            miss = sur[sur["label"] == "MISS"]["post5_ar_%"].mean()
            msg = (f"Average 5-day drift  –  BEAT: {beat:+.2f}%   "
                   f"MISS: {miss:+.2f}%   spread: {beat - miss:+.2f}%")
        else:
            msg = "No earnings-surprise data – run earnings_surprise.py first."
        tk.Label(f, text=msg, bg=BG, fg=INK,
                 font=("Consolas", 11)).pack(padx=12, anchor="w")
        tk.Label(f, text="Green = profit grew YoY, Red = profit fell. The market keeps "
                 "drifting in the surprise direction after results (Post-Earnings-Announcement "
                 "Drift). ⚠ Sample is tiny (2022–24 only) – illustrative, not conclusive.",
                 bg=BG, fg=MUTE, font=("Segoe UI", 9), wraplength=800,
                 justify="left").pack(padx=12, anchor="w", pady=6)

    # -- BACKTEST -------------------------------------------------------
    def _tab_backtest(self):
        f = self._panel("Backtest")
        bt = _read("backtest_summary.csv")
        top = tk.Frame(f, bg=BG)
        top.pack(fill="x", padx=8, pady=8)

        ch1 = Chart(top, width=400, height=320)
        ch1.pack(side="left", padx=4)
        ch1.title("Risk-adjusted return (Sharpe)", "higher = better; >1 is good")
        ch2 = Chart(top, width=400, height=320)
        ch2.pack(side="left", padx=4)
        ch2.title("Avg return per trade", "net of 0.10% cost")

        if bt is not None and not bt.empty:
            # short labels for the x-axis
            def short(s):
                return {"A": "A run-up", "B": "B fade", "C": "C PEAD"}.get(
                    str(s)[0], "Buy&Hold")
            labels = [short(s) for s in bt["strategy"]]
            # Sharpe: green if it beats buy&hold, else muted
            bh_sh = bt[bt["strategy"] == "buy & hold"]["sharpe"].iloc[0] \
                if (bt["strategy"] == "buy & hold").any() else 0
            cols1 = [POS if v >= bh_sh and l != "Buy&Hold" else BLUE
                     for v, l in zip(bt["sharpe"], labels)]
            ch1.bars(labels, bt["sharpe"].tolist(), colors=cols1, yfmt="{:.2f}")
            # avg/trade only for the strategies (buy&hold has none)
            strat = bt[bt["strategy"] != "buy & hold"]
            slabels = [short(s) for s in strat["strategy"]]
            scols = [POS if t > 2 else (AMBER if t > 1 else NEG) for t in strat["t"]]
            ch2.bars(slabels, strat["avg_%"].tolist(), colors=scols, yfmt="{:+.2f}%")

        tk.Label(f, text="Only Strategy A (buy 5 days before results, sell the day before) "
                 "shows a real edge: +1.24%/trade over 72 events, t≈+2.95 (significant), and "
                 "a Sharpe above buy-&-hold. Strategy B (short after results) is noise; "
                 "Strategy C (profit beat→long) looks strong but has only 5 trades. Green "
                 "bars in the right chart = statistically significant (t>2).",
                 bg=BG, fg=MUTE, font=("Segoe UI", 9), wraplength=810,
                 justify="left").pack(padx=12, anchor="w", pady=8)

    # -- FUNDAMENTALS ---------------------------------------------------
    def _tab_fundamentals(self):
        f = self._panel("Fundamentals")
        fund = _read("fundamentals.csv")
        ch = Chart(f, width=800, height=420)
        ch.pack(padx=8, pady=12)
        ch.title("Consolidated quarterly results (₹ crore)",
                 "from parsed XBRL filings")
        if fund is not None and not fund.empty:
            con = fund[fund["consolidated"] == "Consolidated"].copy()
            con["period_end"] = pd.to_datetime(con["period_end"])
            con = con.dropna(subset=["net_profit_cr"]).sort_values("period_end")
            if not con.empty:
                x = list(range(len(con)))
                xl = [d.strftime("%b-%y") for d in con["period_end"]]
                ch.line_series(x, {"Revenue": con["revenue_cr"].tolist(),
                                   "Net profit": con["net_profit_cr"].tolist()},
                               xlabels=xl, zero=False, yfmt="{:,.0f}")
        tk.Label(f, text="Revenue and net profit per quarter, extracted from NSE XBRL. "
                 "Coverage is currently 2022–24 (older filings need more parser work).",
                 bg=BG, fg=MUTE, font=("Segoe UI", 9), wraplength=800,
                 justify="left").pack(padx=12, anchor="w")


def main():
    if not PROC.exists():
        raise SystemExit("No ./processed folder. Run the pipeline scripts first.")
    app = App()
    if "--selftest" in sys.argv:                 # build + render once, then quit
        app.update_idletasks()
        app.update()
        app.after(150, app.destroy)
    app.mainloop()


if __name__ == "__main__":
    main()
