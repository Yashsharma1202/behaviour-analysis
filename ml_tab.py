"""
ml_tab.py
===============================================================================
Shared helper that ADDS an "🤖 ML Signal" tab to the dashboard page and serves the
per-stock ML prediction — used by both the main dashboard (nifty_dash.py) and the
standalone ML host (ml_host.py).

IMPORTANT: this is purely ADDITIVE. It injects a tab into a COPY of the page and
reads ml_output.json — it never touches the permutation/combination win-rate
engine (event_behaviour / build_behaviour / build_rankings). The statistical
accuracy across all four event types is completely unchanged.

    inject(page)   -> the same page with the ML tab wired in
    payload(sym)   -> {"summary": {...}, "pred": {...}} for /api/ml
===============================================================================
"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent

_EXTRA_CSS = """
.mltbl{width:100%;border-collapse:collapse;font-size:.9rem;margin:8px 0}
.mltbl th,.mltbl td{padding:8px 12px;border-bottom:1px solid #232a35;text-align:left}
.mltbl th{font-size:.7rem;text-transform:uppercase;letter-spacing:.04em;color:#98a2b3}
.mlbig{font-size:2.6rem;font-weight:800;margin:6px 0;letter-spacing:-.02em}
</style>"""

_RENDER_ML = r"""
async function renderML(){
  var mv=$("mlview"), sym=SYMS[idx];
  mv.innerHTML='<div class="enote">Loading ML signal…</div>';
  var d;
  try{ d=await (await fetch("/api/ml?sym="+encodeURIComponent(sym))).json(); }
  catch(e){ mv.innerHTML='<div class="banner">ML data unavailable — run <b>python ml_predict.py</b> first.</div>'; return; }
  if(activeTab!=="ml") return;
  if(d.error){ mv.innerHTML='<div class="banner">'+esc(d.error)+'</div>'; return; }
  var s=d.summary, p=d.pred;
  var edge = s.edge ? '<span style="color:#5ec98a;font-weight:700">beats the baseline</span>'
                    : '<span style="color:#f87171;font-weight:700">no tradeable edge</span>';
  var h='';
  h+='<div class="banner" style="border-left:4px solid #fbbf5b;text-align:left">🤖 <b>ML signal (add-on).</b> '
   + 'A separate model that predicts whether a standard event trade (<b>'+s.trade+'</b>) finishes positive — '
   + 'pooled '+s.n.toLocaleString()+' events, time-validated, no look-ahead. '
   + 'It does <b>not</b> change the win-rate statistics in the other tabs. '
   + 'Result: '+edge+' (best AUC '+s.best_auc.toFixed(3)+', where 0.5 = random).</div>';
  if(p){
    var prob=p.prob, col=prob>=55?'#5ec98a':(prob>=50?'#fbbf5b':'#f87171');
    h+='<div class="panel" style="text-align:center;padding:22px">'
     + '<div class="enote">'+esc(sym)+' — model estimate for the latest '+esc(p.event.replace(/_/g," "))+' ('+p.date+')</div>'
     + '<div class="mlbig" style="color:'+col+'">'+prob.toFixed(0)+'%</div>'
     + '<div class="enote">predicted WIN probability &nbsp;·&nbsp; rank '+p.rank+' of '+s.n_stocks+' stocks</div></div>';
  } else {
    h+='<div class="banner">No ML prediction available for '+esc(sym)+'.</div>';
  }
  h+='<h3 style="margin:18px 0 6px">Model performance (out-of-time test, ≥ '+s.test_from+')</h3>';
  h+='<table class="mltbl"><thead><tr><th>Feature set</th><th>Accuracy</th><th>AUC</th></tr></thead><tbody>'
   + '<tr><td>Baseline (always predict WIN)</td><td>'+s.baseline.toFixed(1)+'%</td><td>—</td></tr>'
   + '<tr><td>Price only</td><td>'+s.price_acc.toFixed(1)+'%</td><td>'+s.price_auc.toFixed(3)+'</td></tr>'
   + '<tr><td>Price + Fundamentals</td><td>'+s.fund_acc.toFixed(1)+'%</td><td>'+s.fund_auc.toFixed(3)+'</td></tr>'
   + '</tbody></table>';
  h+='<div class="enote" style="margin-top:10px">The win-rate statistics (permutation/combination grid search) in the '
   + 'Behaviour, Rankings and event tabs are unaffected by this — ML is an add-on lens only.</div>';
  mv.innerHTML=h;
}
"""


def inject(page: str) -> str:
    """Return `page` with the ML tab, view, showTab wiring and renderML added."""
    p = page
    p = p.replace("</style>", _EXTRA_CSS, 1)
    p = p.replace('<button class="tab" data-tab="behaviour">📈 Behaviour</button>',
                  '<button class="tab" data-tab="behaviour">📈 Behaviour</button>\n'
                  '      <button class="tab" data-tab="ml">🤖 ML Signal</button>')
    p = p.replace('<div id="behaviourview" style="display:none"></div>',
                  '<div id="behaviourview" style="display:none"></div>\n'
                  '    <div id="mlview" style="display:none"></div>')
    p = p.replace('isScr=name==="screener", isCmp=name==="compare";',
                  'isScr=name==="screener", isCmp=name==="compare", isML=name==="ml";')
    p = p.replace('$("compareview").style.display=isCmp?"":"none";',
                  '$("compareview").style.display=isCmp?"":"none";\n'
                  '  $("mlview").style.display=isML?"":"none";')
    p = p.replace('$("eventsview").style.display=(isFund||isBeh||isRank||isScr||isCmp)?"none":"";',
                  '$("eventsview").style.display=(isFund||isBeh||isRank||isScr||isCmp||isML)?"none":"";')
    p = p.replace('else if(isCmp) renderCompare();',
                  'else if(isCmp) renderCompare();\n  else if(isML) renderML();')
    p = p.replace('function showTab(name){', _RENDER_ML + '\nfunction showTab(name){')
    return p


def payload(sym: str) -> dict:
    """The /api/ml response for one symbol, read from ml_output.json."""
    f = ROOT / "ml_output.json"
    if not f.exists():
        return {"error": "ML not built yet — run: python ml_predict.py"}
    o = json.loads(f.read_text(encoding="utf-8"))
    preds = o.get("preds", [])
    summary = {
        "trade": o["trade"], "n": o["n"], "n_stocks": len(preds),
        "baseline": o["baseline"], "edge": o["edge"], "test_from": o["test_from"],
        "price_acc": o["price"]["best"] * 100, "price_auc": max(o["price"]["lr_auc"], o["price"]["gb_auc"]),
        "fund_acc": o["fund"]["best"] * 100, "fund_auc": max(o["fund"]["lr_auc"], o["fund"]["gb_auc"]),
        "best_auc": max(o["price"]["lr_auc"], o["price"]["gb_auc"], o["fund"]["lr_auc"], o["fund"]["gb_auc"]),
    }
    pred = None
    for i, r in enumerate(preds, 1):
        if r["symbol"] == sym:
            pred = dict(r); pred["rank"] = i
            break
    return {"summary": summary, "pred": pred}
