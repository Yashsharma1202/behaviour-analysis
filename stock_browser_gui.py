"""
stock_browser_gui.py
===============================================================================
NSE STOCK BROWSER — one stock at a time, starting from the first.
-------------------------------------------------------------------------------
WHAT IT DOES
    Walks through EVERY stock that has fundamentals on disk (the pnl / quarterly /
    ratios / balance_sheet / cash_flow folders), one at a time. It opens on the
    FIRST stock (alphabetically) and gives you:

        [<< First]  [< Prev]   <symbol dropdown>   [Next >]  [Last >>]
        + arrow-key navigation (Left / Right)

    For the selected stock it shows, entirely OFFLINE (no internet):
        * snapshot cards  : latest Sales, Net Profit, OPM%, ROCE%
        * annual trend    : Sales vs Net Profit over the years
        * quarterly chart : Net Profit per quarter, green/red by YoY growth

DATA
    Reuses fund_loader.load_stock(symbol) to parse the Screener-style CSVs, so it
    works for ALL ~3,400 stocks, not just Reliance.

RUN
    python stock_browser_gui.py
===============================================================================
"""

from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path
from tkinter import ttk

import pandas as pd

import fund_loader

ROOT = Path(__file__).resolve().parent
STATEMENT_DIRS = ["pnl", "quarterly", "ratios", "balance_sheet", "cash_flow"]

# --- "trading terminal" dark palette (matches dashboard_gui.py) --------------
BG = "#0f131b"
PANEL = "#171d2b"
GRID = "#2a3346"
INK = "#e6edf3"
MUTE = "#8b98ad"
POS = "#3fb950"
NEG = "#f85149"
BLUE = "#58a6ff"
AMBER = "#d29922"
PURPLE = "#bc8cff"
SERIES = [BLUE, AMBER, PURPLE, POS, NEG]


# ===========================================================================
# Tiny Canvas charting toolkit (no external plotting libraries)
# ===========================================================================
class Chart(tk.Canvas):
    """A canvas that can draw simple axes, line series and bars."""

    def __init__(self, master, width=460, height=300, **kw):
        super().__init__(master, width=width, height=height, bg=PANEL,
                         highlightthickness=0, **kw)
        self.W, self.H = width, height
        self.ml, self.mr, self.mt, self.mb = 58, 20, 42, 38

    def _box(self):
        return self.ml, self.mt, self.W - self.mr, self.H - self.mb

    def _sx(self, v, lo, hi):
        x0, _, x1, _ = self._box()
        return x0 + (v - lo) / ((hi - lo) or 1) * (x1 - x0)

    def _sy(self, v, lo, hi):
        _, y0, _, y1 = self._box()
        return y1 - (v - lo) / ((hi - lo) or 1) * (y1 - y0)

    def title(self, text, sub=""):
        self.create_text(self.ml, 14, text=text, fill=INK, anchor="w",
                         font=("Segoe UI Semibold", 11))
        if sub:
            self.create_text(self.W - self.mr, 14, text=sub, fill=MUTE, anchor="e",
                             font=("Segoe UI", 8))

    def empty(self, msg="no data for this stock"):
        self.create_text(self.W / 2, self.H / 2, text=msg, fill=MUTE,
                         font=("Segoe UI", 10))

    def _yaxis(self, lo, hi, fmt="{:,.0f}", ticks=5):
        x0, _, x1, _ = self._box()
        for i in range(ticks + 1):
            v = lo + (hi - lo) * i / ticks
            y = self._sy(v, lo, hi)
            self.create_line(x0, y, x1, y, fill=GRID)
            self.create_text(x0 - 7, y, text=fmt.format(v), fill=MUTE, anchor="e",
                             font=("Consolas", 8))

    def line_series(self, x, series: dict, xlabels=None, zero=False, yfmt="{:,.0f}"):
        allv = [v for ys in series.values() for v in ys if pd.notna(v)]
        if not allv:
            self.empty()
            return
        lo = min(allv + ([0] if zero else []))
        hi = max(allv + ([0] if zero else []))
        pad = (hi - lo) * 0.12 or 1
        lo, hi = lo - pad, hi + pad
        self._yaxis(lo, hi, yfmt)
        xlo, xhi = min(x), max(x)
        for i, xv in enumerate(x):
            lab = xlabels[i] if xlabels else str(xv)
            self.create_text(self._sx(xv, xlo, xhi), self.H - self.mb + 13,
                             text=lab, fill=MUTE, font=("Consolas", 8))
        for idx, (name, ys) in enumerate(series.items()):
            col = SERIES[idx % len(SERIES)]
            pts = []
            for xv, yv in zip(x, ys):
                if pd.isna(yv):
                    continue
                pts += [self._sx(xv, xlo, xhi), self._sy(yv, lo, hi)]
            if len(pts) >= 4:
                self.create_line(*pts, fill=col, width=2)
            for xv, yv in zip(x, ys):
                if pd.isna(yv):
                    continue
                px, py = self._sx(xv, xlo, xhi), self._sy(yv, lo, hi)
                self.create_oval(px - 3, py - 3, px + 3, py + 3, fill=col, outline=PANEL)
            ly = self.mt + 2 + idx * 15
            self.create_line(self.W - self.mr - 118, ly, self.W - self.mr - 98, ly,
                             fill=col, width=3)
            self.create_text(self.W - self.mr - 92, ly, text=name, fill=INK,
                             anchor="w", font=("Segoe UI", 9))

    def bars(self, cats, values, colors=None, yfmt="{:,.0f}", value_tags=True):
        vals = [0 if pd.isna(v) else v for v in values]
        if not vals:
            self.empty()
            return
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
            self.create_text(cx, self.H - self.mb + 13, text=str(c), fill=MUTE,
                             font=("Consolas", 8))
            if value_tags:
                self.create_text(cx, y - 8 if v >= 0 else y + 8,
                                 text=yfmt.format(v), fill=INK, font=("Consolas", 8))


# ===========================================================================
# Data helpers
# ===========================================================================
def discover_symbols() -> list[str]:
    """Union of every <symbol>.csv stem across the five statement folders."""
    syms: set[str] = set()
    for st in STATEMENT_DIRS:
        d = ROOT / st
        if d.exists():
            syms.update(p.stem for p in d.glob("*.csv"))
    # Drop purely-numeric BSE scrip codes (e.g. 500142) — those are old/delisted
    # entries with stale, incomplete data. Keep proper alphabetic NSE symbols.
    return sorted(s for s in syms if not s.isdigit())


def _col(df, name):
    """Find a metric column by exact match, else 'contains', else None."""
    if df is None:
        return None
    for c in df.columns:
        if c.strip() == name:
            return c
    for c in df.columns:
        if name.lower() in c.lower():
            return c
    return None


def _latest(df, name):
    """Latest non-NaN value of a metric, or NaN."""
    col = _col(df, name)
    if col is None:
        return float("nan")
    s = df[col].dropna()
    return s.iloc[-1] if not s.empty else float("nan")


def _fmt(v, suffix="", nd=0):
    if pd.isna(v):
        return "—"
    return f"{v:,.{nd}f}{suffix}"


# ===========================================================================
# The application
# ===========================================================================
class StockBrowser(tk.Tk):
    def __init__(self, symbols: list[str]):
        super().__init__()
        self.symbols = symbols
        self.idx = 0                      # start from the FIRST stock

        self.title("NSE Stock Browser  —  one stock at a time")
        self.configure(bg=BG)
        self.geometry("1000x760")
        self.minsize(880, 680)

        self._init_style()
        self._build_navbar()

        self.body = tk.Frame(self, bg=BG)
        self.body.pack(fill="both", expand=True, padx=14, pady=(4, 12))

        # keyboard navigation
        self.bind("<Left>", lambda e: self.step(-1))
        self.bind("<Right>", lambda e: self.step(1))
        self.bind("<Home>", lambda e: self.goto(0))
        self.bind("<End>", lambda e: self.goto(len(self.symbols) - 1))

        self.render()

    # -- styling --------------------------------------------------------
    def _init_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Nav.TButton", background=PANEL, foreground=INK,
                        borderwidth=0, focuscolor=BG, padding=(12, 6),
                        font=("Segoe UI Semibold", 10))
        style.map("Nav.TButton",
                  background=[("active", "#243044"), ("pressed", "#243044")])
        style.configure("TCombobox", fieldbackground=PANEL, background=PANEL,
                        foreground=INK, arrowcolor=INK, borderwidth=0,
                        padding=6)
        self.option_add("*TCombobox*Listbox.background", PANEL)
        self.option_add("*TCombobox*Listbox.foreground", INK)
        self.option_add("*TCombobox*Listbox.selectBackground", BLUE)
        self.option_add("*TCombobox*Listbox.font", ("Consolas", 10))

    # -- navigation bar -------------------------------------------------
    def _build_navbar(self):
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=14, pady=(12, 6))

        tk.Label(bar, text="NSE Stock Browser", bg=BG, fg=INK,
                 font=("Segoe UI Semibold", 15)).pack(side="left")

        # right-aligned controls
        ctrl = tk.Frame(bar, bg=BG)
        ctrl.pack(side="right")

        ttk.Button(ctrl, text="◀◀ First", style="Nav.TButton",
                   command=lambda: self.goto(0)).pack(side="left", padx=2)
        ttk.Button(ctrl, text="◀ Prev", style="Nav.TButton",
                   command=lambda: self.step(-1)).pack(side="left", padx=2)

        self.combo = ttk.Combobox(ctrl, values=self.symbols, width=16,
                                   font=("Consolas", 11), justify="center")
        self.combo.pack(side="left", padx=6)
        self.combo.bind("<<ComboboxSelected>>", self._on_pick)
        self.combo.bind("<Return>", self._on_pick)

        ttk.Button(ctrl, text="Next ▶", style="Nav.TButton",
                   command=lambda: self.step(1)).pack(side="left", padx=2)
        ttk.Button(ctrl, text="Last ▶▶", style="Nav.TButton",
                   command=lambda: self.goto(len(self.symbols) - 1)).pack(side="left", padx=2)

        self.counter = tk.Label(bar, text="", bg=BG, fg=MUTE,
                                font=("Consolas", 10))
        self.counter.pack(side="right", padx=12)

    # -- navigation logic ----------------------------------------------
    def step(self, delta):
        self.goto(self.idx + delta)

    def goto(self, i):
        if not self.symbols:
            return
        self.idx = max(0, min(i, len(self.symbols) - 1))
        self.render()

    def _on_pick(self, _=None):
        want = self.combo.get().strip().upper()
        if want in self.symbols:
            self.goto(self.symbols.index(want))
        else:
            # jump to the first symbol that starts with the typed text
            for j, s in enumerate(self.symbols):
                if s.startswith(want):
                    self.goto(j)
                    return
            self.render()   # reset combobox to current if no match

    # -- render one stock ----------------------------------------------
    def render(self):
        for w in self.body.winfo_children():
            w.destroy()

        sym = self.symbols[self.idx]
        self.combo.set(sym)
        self.counter.config(text=f"{self.idx + 1:,} / {len(self.symbols):,}")

        data = fund_loader.load_stock(sym)
        pnl = data.get("pnl")
        q = data.get("quarterly")
        rat = data.get("ratios")
        bs = data.get("balance_sheet")

        # ---- header line: symbol + which statements were found ----------
        head = tk.Frame(self.body, bg=BG)
        head.pack(fill="x", pady=(2, 8))
        tk.Label(head, text=sym, bg=BG, fg=INK,
                 font=("Segoe UI Semibold", 22)).pack(side="left")
        have = ", ".join(k for k in STATEMENT_DIRS if k in data) or "no statements"
        tk.Label(head, text=f"   statements: {have}", bg=BG, fg=MUTE,
                 font=("Consolas", 9)).pack(side="left", anchor="s", pady=6)

        if not data:
            tk.Label(self.body, text="No fundamental files found for this symbol.",
                     bg=PANEL, fg=MUTE, font=("Segoe UI", 11),
                     padx=20, pady=30).pack(fill="x", padx=4, pady=20)
            return

        # ---- snapshot cards --------------------------------------------
        sales = _latest(pnl, "Sales")
        sales_g = _latest(pnl, "Sales Growth %")
        npft = _latest(pnl, "Net Profit")
        npft_g = _latest(pnl, "Profit Growth %")
        opm = _latest(pnl, "OPM %")
        roce = _latest(rat, "ROCE %")

        eq = _latest(bs, "Equity Capital")
        res = _latest(bs, "Reserves")
        bor = _latest(bs, "Borrowings")
        de = (bor / (eq + res)) if (pd.notna(bor) and pd.notna(eq)
                                    and pd.notna(res) and (eq + res)) else float("nan")

        cards = tk.Frame(self.body, bg=BG)
        cards.pack(fill="x", pady=(0, 10))
        card_defs = [
            ("Sales (latest FY)", _fmt(sales, " cr"), _fmt(sales_g, "% growth", 1), INK),
            ("Net Profit", _fmt(npft, " cr"), _fmt(npft_g, "% growth", 1), POS if (pd.notna(npft) and npft >= 0) else NEG),
            ("Operating margin", _fmt(opm, "%", 1), "OPM", BLUE),
            ("ROCE", _fmt(roce, "%", 1), f"Debt/Equity {_fmt(de, '', 2)}", AMBER),
        ]
        for i, (t, v, s, col) in enumerate(card_defs):
            self._card(cards, t, v, s, col).grid(row=0, column=i, padx=5, sticky="nsew")
            cards.columnconfigure(i, weight=1)

        # ---- charts row -------------------------------------------------
        charts = tk.Frame(self.body, bg=BG)
        charts.pack(fill="both", expand=True)

        c1 = Chart(charts, width=470, height=330)
        c1.pack(side="left", padx=(0, 6))
        c1.title("Annual trend (₹ crore)", "Sales vs Net Profit")
        if pnl is not None and not pnl.empty:
            sc, nc = _col(pnl, "Sales"), _col(pnl, "Net Profit")
            xl = [d.strftime("%Y") for d in pnl.index]
            series = {}
            if sc:
                series["Sales"] = pnl[sc].tolist()
            if nc:
                series["Net Profit"] = pnl[nc].tolist()
            if series:
                c1.line_series(list(range(len(pnl))), series, xlabels=xl, yfmt="{:,.0f}")
            else:
                c1.empty()
        else:
            c1.empty("no annual P&L")

        c2 = Chart(charts, width=470, height=330)
        c2.pack(side="left", padx=(6, 0))
        c2.title("Quarterly Net Profit (₹ cr)", "green = YoY up, red = down")
        if q is not None and not q.empty:
            npc = _col(q, "Net Profit")
            yoyc = _col(q, "YOY Profit Growth %")
            if npc:
                sub = q.tail(12)
                cats = [d.strftime("%b-%y") for d in sub.index]
                vals = sub[npc].tolist()
                if yoyc:
                    cols = [POS if (pd.notna(g) and g >= 0) else NEG for g in sub[yoyc]]
                else:
                    cols = [BLUE] * len(vals)
                c2.bars(cats, vals, colors=cols, yfmt="{:,.0f}")
            else:
                c2.empty("no quarterly net profit")
        else:
            c2.empty("no quarterly data")

        # ---- footer hint ------------------------------------------------
        tk.Label(self.body,
                 text="← / → arrow keys to move  ·  type a symbol in the box and press Enter to jump  "
                      "·  data is read from the local fundamentals folders (offline).",
                 bg=BG, fg=MUTE, font=("Segoe UI", 9)).pack(anchor="w", pady=(8, 0))

    def _card(self, parent, title, value, sub, color=INK):
        c = tk.Frame(parent, bg=PANEL, padx=16, pady=12)
        tk.Label(c, text=title, bg=PANEL, fg=MUTE,
                 font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(c, text=value, bg=PANEL, fg=color,
                 font=("Segoe UI Semibold", 20)).pack(anchor="w")
        tk.Label(c, text=sub, bg=PANEL, fg=MUTE,
                 font=("Segoe UI", 8)).pack(anchor="w")
        return c


# ===========================================================================
def main():
    symbols = discover_symbols()
    if not symbols:
        raise SystemExit(
            "No stock fundamentals found. Expected per-stock CSVs in the "
            "pnl/ quarterly/ ratios/ balance_sheet/ cash_flow/ folders.")

    print(f"Found {len(symbols):,} stocks. First: {symbols[0]}  Last: {symbols[-1]}")
    app = StockBrowser(symbols)

    if "--selftest" in sys.argv:            # render first stock once, then quit
        app.update_idletasks()
        app.update()
        app.after(200, app.destroy)
    app.mainloop()


if __name__ == "__main__":
    main()
