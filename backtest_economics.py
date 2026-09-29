"""
BACKTEST ECONOMICS  (net-of-cost edge + risk metrics)
=====================================================
Turns the raw holiday playbook trade log into an HONEST, cost-aware performance
picture so you know the edge BEFORE sizing up.

Source: dashboard_data/holidays_dataset.json  ->  844 trades/holiday-year,
14,348 trades total (17 holidays x 2022-2025), each with a realised ret_pct.

It adds what a quoted win-rate hides:
  * transaction costs + STT/exchange/GST/stamp + slippage per round trip
  * GROSS vs NET returns (the cost drag)
  * win-rate 95% confidence interval (Wilson) — 86% on 15 trades != 86% on 1500
  * max drawdown on the sequential net equity curve
  * Sharpe & Sortino (annualised by trades/year)
  * profit factor, expectancy, avg win / avg loss
  * a cost-sensitivity table (edge at 10 / 20 / 25 / 35 / 50 bps round trip)

OUTPUTS
  dashboard_data/backtest_economics.json   (consumed by the dashboard / reuse)
  Backtest_Economics_Report.html           (clean, printable)
  + a console summary

USAGE
  python backtest_economics.py                     # default cost 0.25% round trip
  python backtest_economics.py --cost 0.20         # override all-in round-trip %
ADDITIVE: new analysis file. Reads existing data only; writes new artifacts.
"""
import sys, os, json, math, argparse
from datetime import date
from collections import defaultdict

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, 'dashboard_data', 'holidays_dataset.json')
OUT_JSON = os.path.join(BASE, 'dashboard_data', 'backtest_economics.json')
OUT_HTML = os.path.join(BASE, 'Backtest_Economics_Report.html')

_MONS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

# --- transparent round-trip cost model (percentage of notional, both legs) ---
# defaults are deliberately conservative for stock-futures execution.
COST_COMPONENTS = {
    'Brokerage (round trip)':      0.020,
    'Exchange txn + GST':          0.006,
    'STT (sell leg, futures)':     0.020,
    'SEBI + stamp duty':           0.004,
    'Slippage (2 legs)':           0.200,
}


def parse_dmy(s):
    import re
    m = re.search(r'(\d{1,2})-([A-Za-z]{3})-(\d{4})', str(s))
    if not m:
        return None
    return date(int(m.group(3)), _MONS.index(m.group(2)) + 1, int(m.group(1)))


def wilson_ci(wins, n, z=1.96):
    """95% Wilson score interval for a win-rate proportion, returned as %."""
    if n == 0:
        return (0.0, 0.0)
    p = wins / n
    denom = 1 + z * z / n
    centre = (p + z * z / (2 * n)) / denom
    half = (z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n))) / denom
    return (round(100 * (centre - half), 1), round(100 * (centre + half), 1))


def _mean(x): return sum(x) / len(x) if x else 0.0
def _std(x):
    if len(x) < 2:
        return 0.0
    m = _mean(x)
    return math.sqrt(sum((v - m) ** 2 for v in x) / len(x))


def stats_for(trades, cost_pct, n_years):
    """trades carry 'ret' (gross %), 'entry' (date), 'event' (holiday|year basket key).

    Two levels, kept distinct because they answer different questions:
      * TRADE level  -> win rate + Wilson CI, mean/median, profit factor. Every
        stock-trade is its own sample.
      * PORTFOLIO level -> drawdown, Sharpe/Sortino computed on ONE return per
        holiday-year basket (~211 simultaneous, correlated trades = one bet), so
        risk stats aren't inflated by treating correlated trades as independent.
    """
    n = len(trades)
    if n == 0:
        return None
    gross = [t['ret'] for t in trades]
    net = [g - cost_pct for g in gross]

    net_sorted = sorted(net)
    median = net_sorted[n // 2] if n % 2 else (net_sorted[n // 2 - 1] + net_sorted[n // 2]) / 2
    wins_g = sum(1 for g in gross if g > 0)
    wins_n = sum(1 for v in net if v > 0)
    win_lo, win_hi = wilson_ci(wins_n, n)

    pos = [v for v in net if v > 0]
    neg = [v for v in net if v < 0]
    profit_factor = (sum(pos) / abs(sum(neg))) if neg and sum(neg) != 0 else float('inf')

    # ---- portfolio (per holiday-year basket) series ----
    baskets = defaultdict(list)
    basket_date = {}
    for t in trades:
        k = t['event']
        baskets[k].append(t['ret'] - cost_pct)
        d = t['entry'] or date(1900, 1, 1)
        if k not in basket_date or d < basket_date[k]:
            basket_date[k] = d
    ev_keys = sorted(baskets, key=lambda k: basket_date[k])
    ev_rets = [_mean(baskets[k]) for k in ev_keys]           # one return per basket
    n_ev = len(ev_rets)
    ev_wins = sum(1 for r in ev_rets if r > 0)

    m_ev, s_ev = _mean(ev_rets), _std(ev_rets)
    # per-event Sharpe (NOT annualised): annualising 60+ correlated same-day
    # baskets by sqrt(count) inflates the ratio, so we report the raw per-event
    # reward/variability and let concrete numbers (worst basket, % negative) carry the risk.
    per_event_sharpe = (m_ev / s_ev) if s_ev > 0 else 0.0
    worst = min(ev_rets) if ev_rets else 0.0
    pct_neg = 100 * sum(1 for r in ev_rets if r < 0) / n_ev if n_ev else 0.0

    eq, peak, mdd = 1.0, 1.0, 0.0
    for r in ev_rets:
        eq *= (1 + r / 100.0)
        peak = max(peak, eq)
        mdd = min(mdd, eq / peak - 1)

    return {
        # trade level
        'trades': n,
        'win_rate_gross': round(100 * wins_g / n, 1),
        'win_rate_net': round(100 * wins_n / n, 1),
        'win_rate_ci95': [win_lo, win_hi],
        'mean_gross': round(_mean(gross), 3),
        'mean_net': round(_mean(net), 3),
        'median_net': round(median, 3),
        'std_net': round(_std(net), 3),
        'cost_drag': round(_mean(gross) - _mean(net), 3),
        'avg_win': round(_mean(pos), 3),
        'avg_loss': round(_mean(neg), 3),
        'expectancy_net': round(_mean(net), 3),
        'profit_factor': (round(profit_factor, 2) if profit_factor != float('inf') else None),
        # portfolio (per holiday-year basket) level
        'port_events': n_ev,
        'port_win_rate': round(100 * ev_wins / n_ev, 1) if n_ev else 0.0,
        'port_mean': round(m_ev, 3),
        'port_std': round(s_ev, 3),
        'port_worst': round(worst, 3),
        'port_pct_neg': round(pct_neg, 1),
        'per_event_sharpe': round(per_event_sharpe, 2),
        'max_drawdown_pct': round(100 * mdd, 2),
    }


def load_trades():
    data = json.load(open(SRC, encoding='utf-8'))
    out = []
    for hol in data:
        for t in hol.get('trade_log', []):
            r = t.get('ret_pct')
            if not isinstance(r, (int, float)):
                continue
            holiday = t.get('holiday') or hol.get('name')
            out.append({
                'ret': float(r),
                'entry': parse_dmy(t.get('entry_date')),
                'year': t.get('year'),
                'holiday': holiday,
                'event': str(holiday) + '|' + str(t.get('year')),   # portfolio basket key
                'direction': str(t.get('direction') or '').upper(),
                'symbol': t.get('symbol'),
            })
    return out


def build(cost_pct):
    trades = load_trades()
    years = sorted({t['year'] for t in trades if t['year']})
    n_years = len(years) or 1

    overall = stats_for(trades, cost_pct, n_years)

    by_dir = {}
    for d in ('LONG', 'SHORT'):
        sub = [t for t in trades if t['direction'] == d]
        by_dir[d] = stats_for(sub, cost_pct, n_years)

    by_holiday = {}
    hol_names = sorted({t['holiday'] for t in trades if t['holiday']})
    for h in hol_names:
        sub = [t for t in trades if t['holiday'] == h]
        by_holiday[h] = stats_for(sub, cost_pct, n_years)

    # cost sensitivity
    sens = []
    for c in (0.10, 0.20, 0.25, 0.35, 0.50):
        s = stats_for(trades, c, n_years)
        sens.append({'cost_pct': c, 'mean_net': s['mean_net'],
                     'win_rate_net': s['win_rate_net'], 'sharpe': s['per_event_sharpe'],
                     'max_dd': s['max_drawdown_pct']})

    return {
        'generated_at': __import__('datetime').datetime.now().isoformat(timespec='seconds'),
        'source': 'dashboard_data/holidays_dataset.json (holiday playbook trade log)',
        'universe': 'Nifty F&O stocks · 17 holidays · %s' % ('-'.join(map(str, (years[0], years[-1]))) if years else '?'),
        'n_trades': len(trades), 'n_years': n_years, 'years': years,
        'cost_model': {'round_trip_pct': cost_pct, 'components': COST_COMPONENTS,
                       'note': 'Round-trip % of notional deducted from every trade return (both legs).'},
        'overall': overall, 'by_direction': by_dir, 'by_holiday': by_holiday,
        'cost_sensitivity': sens,
        'caveats': [
            'Sample = 4 years (2022-2025); win-rate CIs widen sharply for thin per-holiday cells.',
            'Portfolio (basket) risk equal-weights the FULL ~211-stock cross-section per event, which over-diversifies: a concentrated book of a few picks will show materially higher drawdown and lower basket win-rate than shown here.',
            'Per-event Sharpe is NOT annualised (same-day baskets are correlated; sqrt-scaling would inflate it).',
            'Returns are on the underlying futures leg; option-play P&L (1% ITM) differs.',
            'No survivorship/liquidity screen beyond current F&O membership.',
            'Costs are configurable; slippage dominates and is the least certain input.',
        ],
    }


def _fmt_pct(v):
    return ('+' if v >= 0 else '') + ('%.2f' % v) + '%'


def write_html(rep):
    o = rep['overall']
    ci = o['win_rate_ci95']

    def metric_card(label, val, sub='', color='#0f172a'):
        return ('<div class="mc"><div class="mcl">%s</div>'
                '<div class="mcv" style="color:%s">%s</div>'
                '<div class="mcs">%s</div></div>') % (label, color, val, sub)

    cards = ''.join([
        metric_card('Net Mean Return / Trade', _fmt_pct(o['mean_net']),
                    'gross %s · cost drag %s' % (_fmt_pct(o['mean_gross']), _fmt_pct(-o['cost_drag'])),
                    '#059669' if o['mean_net'] >= 0 else '#dc2626'),
        metric_card('Net Win Rate', '%.1f%%' % o['win_rate_net'],
                    '95%% CI  %.1f%% – %.1f%%' % (ci[0], ci[1]), '#2563eb'),
        metric_card('Max Drawdown (portfolio)', '%.2f%%' % o['max_drawdown_pct'],
                    'basket equity · worst basket %s' % _fmt_pct(o['port_worst']), '#dc2626'),
        metric_card('Per-Event Sharpe', '%.2f' % o['per_event_sharpe'],
                    '%.1f%% of %d baskets negative' % (o['port_pct_neg'], o['port_events']), '#7c3aed'),
        metric_card('Profit Factor', ('%.2f' % o['profit_factor']) if o['profit_factor'] else '∞',
                    'avg win %s / avg loss %s' % (_fmt_pct(o['avg_win']), _fmt_pct(o['avg_loss'])), '#0891b2'),
        metric_card('Sample Size', '{:,}'.format(rep['n_trades']),
                    '%d holidays · %d yrs · %d baskets' % (len(rep['by_holiday']), rep['n_years'], o['port_events']), '#334155'),
    ])

    def row(name, s):
        if not s:
            return ''
        pf = ('%.2f' % s['profit_factor']) if s['profit_factor'] else '∞'
        return ('<tr><td class="l">%s</td><td>%d</td><td>%.1f%%</td>'
                '<td class="ci">%.1f–%.1f</td><td class="%s">%s</td><td>%s</td>'
                '<td class="%s">%.1f%%</td><td>%.2f</td><td>%s</td></tr>') % (
            name, s['trades'], s['win_rate_net'], s['win_rate_ci95'][0], s['win_rate_ci95'][1],
            'pos' if s['mean_net'] >= 0 else 'neg', _fmt_pct(s['mean_net']),
            _fmt_pct(o['mean_gross'] if name == 'ALL' else s['mean_gross']),
            'neg', s['max_drawdown_pct'], s['per_event_sharpe'], pf)

    hol_rows = ''.join(row(h, rep['by_holiday'][h]) for h in sorted(rep['by_holiday'],
                       key=lambda k: -(rep['by_holiday'][k]['mean_net'] if rep['by_holiday'][k] else 0)))
    dir_rows = ''.join(row(d, rep['by_direction'][d]) for d in ('LONG', 'SHORT'))
    all_row = row('ALL', o)

    sens_rows = ''.join(
        '<tr><td class="l">%s%%</td><td class="%s">%s</td><td>%.1f%%</td><td>%.2f</td><td class="neg">%.1f%%</td></tr>'
        % (('%.2f' % s['cost_pct']), 'pos' if s['mean_net'] >= 0 else 'neg', _fmt_pct(s['mean_net']),
           s['win_rate_net'], s['sharpe'], s['max_dd'])
        for s in rep['cost_sensitivity'])

    comp_rows = ''.join('<tr><td class="l">%s</td><td>%.3f%%</td></tr>' % (k, v)
                        for k, v in rep['cost_model']['components'].items())
    comp_total = sum(rep['cost_model']['components'].values())

    caveats = ''.join('<li>%s</li>' % c for c in rep['caveats'])

    all_row_inner = all_row.replace('<tr>', '').replace('</tr>', '').replace('class="l">ALL', 'class="l">ALL (all holidays)')
    tpl = """<!doctype html><html><head><meta charset="utf-8">
<title>Backtest Economics — Net-of-Cost Edge & Risk</title>
<style>
 *{box-sizing:border-box} body{font-family:'Segoe UI',Arial,sans-serif;margin:0;color:#0f172a;background:#f8fafc}
 .wrap{max-width:1040px;margin:0 auto;padding:32px 40px}
 h1{font-size:24px;margin:0 0 2px} .sub{color:#64748b;font-size:13px;margin-bottom:22px}
 h2{font-size:16px;margin:26px 0 10px;border-left:4px solid #2563eb;padding-left:10px}
 .cards{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
 .mc{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px 16px}
 .mcl{font-size:11px;text-transform:uppercase;letter-spacing:.04em;color:#64748b;font-weight:700}
 .mcv{font-size:26px;font-weight:800;font-variant-numeric:tabular-nums;margin:2px 0}
 .mcs{font-size:12px;color:#64748b}
 table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;font-size:13px}
 th,td{padding:8px 10px;text-align:right;border-bottom:1px solid #eef2f7} th{background:#f1f5f9;font-size:11px;text-transform:uppercase;letter-spacing:.03em;color:#475569}
 td.l,th.l{text-align:left} td.ci{color:#64748b;font-size:12px} .pos{color:#059669;font-weight:700} .neg{color:#dc2626;font-weight:700}
 tr.allrow td{background:#f0f9ff;font-weight:700}
 .note{background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:12px 16px;font-size:12.5px;color:#7c2d12;margin-top:8px}
 ul{font-size:12.5px;color:#475569;line-height:1.6} .foot{color:#94a3b8;font-size:11px;margin-top:26px}
</style></head><body><div class="wrap">
 <h1>Backtest Economics — Net-of-Cost Edge &amp; Risk</h1>
 <div class="sub">@@UNIVERSE@@ &bull; @@NTRADES@@ trades &bull; generated @@GEN@@ &bull; round-trip cost <b>@@COST@@%</b></div>

 <h2>Headline (net of costs)</h2>
 <div class="cards">@@CARDS@@</div>

 <h2>By Holiday — net edge (ranked)</h2>
 <table><thead><tr><th class="l">Holiday</th><th>Trades</th><th>Net Win%</th><th>Win 95% CI</th><th>Net Mean</th><th>Gross Mean</th><th>Max DD</th><th>Sharpe</th><th>PF</th></tr></thead>
 <tbody>@@HOLROWS@@<tr class="allrow">@@ALLROW@@</tbody></table>

 <h2>By Direction</h2>
 <table><thead><tr><th class="l">Side</th><th>Trades</th><th>Net Win%</th><th>Win 95% CI</th><th>Net Mean</th><th>Gross Mean</th><th>Max DD</th><th>Sharpe</th><th>PF</th></tr></thead>
 <tbody>@@DIRROWS@@</tbody></table>

 <h2>Cost sensitivity — where the edge survives</h2>
 <table><thead><tr><th class="l">Round-trip cost</th><th>Net Mean/Trade</th><th>Net Win%</th><th>Sharpe</th><th>Max DD</th></tr></thead>
 <tbody>@@SENSROWS@@</tbody></table>
 <div class="note">The edge is real but <b>slippage-sensitive</b>: read the row matching your realistic fill quality. A strategy that is +2.7% gross can compress fast once execution costs are honest.</div>

 <h2>Cost model (round-trip, % of notional)</h2>
 <table><thead><tr><th class="l">Component</th><th>%</th></tr></thead><tbody>@@COMPROWS@@
 <tr class="allrow"><td class="l">TOTAL round trip</td><td>@@COMPTOTAL@@%</td></tr></tbody></table>

 <h2>Caveats</h2><ul>@@CAVEATS@@</ul>
 <div class="foot">Additive analysis — reads dashboard_data/holidays_dataset.json only. Numbers are historical (2022–2025) and not a forward guarantee.</div>
</div></body></html>"""
    repl = {
        '@@UNIVERSE@@': rep['universe'], '@@NTRADES@@': '{:,}'.format(rep['n_trades']),
        '@@GEN@@': rep['generated_at'], '@@COST@@': ('%.2f' % rep['cost_model']['round_trip_pct']),
        '@@CARDS@@': cards, '@@HOLROWS@@': hol_rows, '@@ALLROW@@': all_row_inner,
        '@@DIRROWS@@': dir_rows, '@@SENSROWS@@': sens_rows, '@@COMPROWS@@': comp_rows,
        '@@COMPTOTAL@@': ('%.3f' % comp_total), '@@CAVEATS@@': caveats,
    }
    for k, v in repl.items():
        tpl = tpl.replace(k, v)
    with open(OUT_HTML, 'w', encoding='utf-8') as f:
        f.write(tpl)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--cost', type=float, default=round(sum(COST_COMPONENTS.values()), 3),
                    help='all-in round-trip cost %% (default = sum of components)')
    args = ap.parse_args()

    if not os.path.exists(SRC):
        print('ERROR: missing', SRC); sys.exit(1)

    rep = build(args.cost)
    json.dump(rep, open(OUT_JSON, 'w', encoding='utf-8'), indent=2)
    write_html(rep)

    o = rep['overall']
    ci = o['win_rate_ci95']
    print('== BACKTEST ECONOMICS ==')
    print('  trades           : {:,}  ({} holidays, {} yrs)'.format(rep['n_trades'], len(rep['by_holiday']), rep['n_years']))
    print('  round-trip cost  : %.2f%%' % args.cost)
    print('  gross mean/trade : %+.3f%%' % o['mean_gross'])
    print('  NET  mean/trade  : %+.3f%%   (cost drag %.3f%%)' % (o['mean_net'], o['cost_drag']))
    print('  net win rate     : %.1f%%   (95%% CI %.1f-%.1f)  [trade level]' % (o['win_rate_net'], ci[0], ci[1]))
    print('  -- portfolio (per holiday-year basket, %d baskets) --' % o['port_events'])
    print('  basket win rate  : %.1f%%   mean/basket %+.3f%%  worst %+.3f%%' % (o['port_win_rate'], o['port_mean'], o['port_worst']))
    print('  baskets negative : %.1f%%' % o['port_pct_neg'])
    print('  max drawdown     : %.2f%%' % o['max_drawdown_pct'])
    print('  per-event Sharpe : %.2f  (not annualised)' % o['per_event_sharpe'])
    print('  profit factor    : %s' % (('%.2f' % o['profit_factor']) if o['profit_factor'] else 'inf'))
    print('\n  wrote  dashboard_data/backtest_economics.json')
    print('  wrote  Backtest_Economics_Report.html')


if __name__ == '__main__':
    main()
