"""
LIVE P&L SERVER  (MTM on the position-taking playbook)
======================================================
Serves live mark-to-market P&L for the dashboard's RBI / holiday positions.

MODES
  - MOCK (default): simulates live ticks so the dashboard's Live P&L panel works
    with no credentials. Start it and open the dashboard.
  - LIVE (Angel One SmartAPI): create  smartapi_config.json  (see TEMPLATE below)
    and run  `python live_pnl_server.py --live`.  Uses only `requests` + stdlib
    (TOTP is implemented here; no SmartApi SDK / pyotp needed).

ENDPOINTS  (CORS enabled for the dashboard)
  GET /api/pnl      -> { generated_at, mode, anchor, totals{...}, positions[...] }
  GET /api/health   -> { ok, mode }

ADDITIVE: new local file. Nothing is pushed to the host. Secrets live only in
smartapi_config.json (git-ignored) on your machine.

smartapi_config.json TEMPLATE (for --live, later):
  { "api_key": "...", "client_code": "...", "pin": "1234", "totp_secret": "BASE32SECRET" }
"""
import sys, os, json, time, random, math, base64, hmac, hashlib, struct, argparse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from datetime import datetime, date, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE_DIR, 'dashboard_data', 'event_dashboard_data.json')
FO = os.path.join(BASE_DIR, 'dashboard_data', 'fo_stocks_211.json')
STATE_FILE = os.path.join(BASE_DIR, 'live_pnl_state.json')
CONFIG_FILE = os.path.join(BASE_DIR, 'smartapi_config.json')
PORT = 8787

NSE_2026_HOLIDAYS = {
    # '2026-03-04' removed (not a real NSE holiday -- leftover from Holi's date
    # once being wrongly typed as 04-Mar; real Holi holiday is 03-Mar, kept
    # below). '2026-11-08'/'2026-11-10' (Diwali) added -- were missing entirely.
    '2026-01-26','2026-03-03','2026-03-26','2026-03-31','2026-04-03',
    '2026-04-14','2026-05-01','2026-05-28','2026-06-26','2026-09-14','2026-10-02',
    '2026-10-20','2026-11-08','2026-11-10','2026-11-24','2026-12-25'
}
_MONS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

# ---------------- trading-day helpers (mirror the dashboard) ----------------
def is_trading_day(d):
    return d.weekday() < 5 and d.isoformat() not in NSE_2026_HOLIDAYS

def shift_trading_days(anchor, n):
    d = anchor; step = 1 if n >= 0 else -1; c = abs(n)
    while c > 0:
        d = d + timedelta(days=step)
        if is_trading_day(d):
            c -= 1
    return d

def parse_dmy(s):
    import re
    m = re.search(r'(\d{1,2})-([A-Za-z]{3})-(\d{4})', str(s))
    if not m:
        return None
    return date(int(m.group(3)), _MONS.index(m.group(2)) + 1, int(m.group(1)))

def parse_window(w):
    import re
    m = re.search(r'T-?(\d+)\s*to\s*T\+?(\d+)', str(w), re.I)
    return (int(m.group(1)), int(m.group(2))) if m else (2, 3)

def status_for(entry, exit_, today):
    dE = (entry - today).days
    dX = (exit_ - today).days
    if dE == 0:  return 'ENTER_TODAY'
    if dE == 1:  return 'ENTRY_TOMORROW'
    if dE < 0 <= dX: return 'IN_PROGRESS'
    if dX < 0:   return 'COMPLETED'
    return 'UPCOMING'

# ---------------- load playbook positions ----------------
def load_positions():
    """RBI-policy playbook positions (group='RBI'). Entry/exit derived from the
    MPC decision day + each stock's T-n/T+m window on the NSE-2026 calendar."""
    d = json.load(open(DATA, encoding='utf-8'))
    fo = {f['symbol']: f for f in json.load(open(FO, encoding='utf-8'))}
    rp = (d.get('rbi_policy') or [None])[0]
    if not rp or not rp.get('stocks'):
        return None, []
    anchor = parse_dmy(rp.get('decision_day')) or parse_dmy(rp.get('date')) or date(2026, 10, 9)
    event = 'RBI MPC ' + (anchor.strftime('%d-%b-%Y') if anchor else '')
    out = []
    for s in rp['stocks']:
        lead, hold = parse_window(s.get('window'))
        entry = shift_trading_days(anchor, -lead)
        exit_ = shift_trading_days(anchor, hold)
        f = fo.get(s['symbol'], {})
        out.append({
            'group': 'RBI',
            'event': event,
            'symbol': s['symbol'],
            'name': f.get('name', s['symbol']),
            'sector': s.get('sector', ''),
            'direction': s.get('direction', 'LONG'),
            'window': s.get('window', ''),
            'entry_date': entry.isoformat(),
            'exit_date': exit_.isoformat(),
            'lot': int(f.get('lot_size') or s.get('lot') or 500),
            'ref_price': float(f.get('spot_ltp') or 1000.0),
        })
    return anchor, out


def load_holiday_positions():
    """Holiday playbook positions (group='HOLIDAY'). One block of stock positions
    per upcoming holiday; entry/exit come straight from the dataset's precomputed
    dates (fall back to computing from the window if a date is missing)."""
    d = json.load(open(DATA, encoding='utf-8'))
    fo = {f['symbol']: f for f in json.load(open(FO, encoding='utf-8'))}
    out = []
    for h in (d.get('holidays') or []):
        hname = h.get('name', h.get('id', 'Holiday'))
        anchor = parse_dmy(h.get('date'))
        for s in (h.get('stocks') or []):
            entry = parse_dmy(s.get('entry_date'))
            exit_ = parse_dmy(s.get('exit_date'))
            if (entry is None or exit_ is None) and anchor is not None:
                lead, hold = parse_window(s.get('window'))
                entry = entry or shift_trading_days(anchor, -lead)
                exit_ = exit_ or shift_trading_days(anchor, hold)
            if entry is None or exit_ is None:
                continue
            f = fo.get(s['symbol'], {})
            out.append({
                'group': 'HOLIDAY',
                'event': hname,
                'symbol': s['symbol'],
                'name': f.get('name', s.get('name', s['symbol'])),
                'sector': s.get('sector', ''),
                'direction': s.get('direction', 'LONG'),
                'window': s.get('window', ''),
                'entry_date': entry.isoformat(),
                'exit_date': exit_.isoformat(),
                'lot': int(s.get('lot') or f.get('lot_size') or 500),
                'ref_price': float(f.get('spot_ltp') or 1000.0),
            })
    return out

# ---------------- state (persist captured entry prices + mock walk) ----------------
def load_state():
    try:
        return json.load(open(STATE_FILE, encoding='utf-8'))
    except Exception:
        return {}

def save_state(st):
    try:
        json.dump(st, open(STATE_FILE, 'w', encoding='utf-8'))
    except Exception:
        pass

# ---------------- MOCK price feed ----------------
def mock_ltp(sym, ref, state):
    key = 'mock:' + sym
    last = state.get(key, ref)
    # small random walk, mean-reverting toward ref
    drift = (ref - last) * 0.02
    shock = last * random.uniform(-0.004, 0.004)
    px = max(1.0, last + drift + shock)
    state[key] = px
    return round(px, 2)

# ---------------- ANGEL SmartAPI (LIVE) ----------------
class Angel:
    LOGIN = 'https://apiconnect.angelbroking.com/rest/auth/angelbroking/user/v1/loginByPassword'
    QUOTE = 'https://apiconnect.angelbroking.com/rest/secure/angelbroking/market/v1/quote/'
    SCRIP = 'https://margincalculator.angelbroking.com/OpenAPI_File/files/OpenAPIScripMaster.json'

    def __init__(self, cfg):
        import requests
        self.requests = requests
        self.cfg = cfg
        self.jwt = None
        self.tokens = {}   # symbol -> nse token

    @staticmethod
    def _totp(secret):
        key = base64.b32decode(secret.strip().upper() + '=' * ((8 - len(secret) % 8) % 8))
        msg = struct.pack('>Q', int(time.time()) // 30)
        h = hmac.new(key, msg, hashlib.sha1).digest()
        o = h[19] & 15
        return '%06d' % ((struct.unpack('>I', h[o:o + 4])[0] & 0x7fffffff) % 1000000)

    def _headers(self, auth=False):
        h = {
            'Content-Type': 'application/json', 'Accept': 'application/json',
            'X-UserType': 'USER', 'X-SourceID': 'WEB',
            'X-ClientLocalIP': '127.0.0.1', 'X-ClientPublicIP': '127.0.0.1',
            'X-MACAddress': '00:00:00:00:00:00', 'X-PrivateKey': self.cfg['api_key'],
        }
        if auth and self.jwt:
            h['Authorization'] = 'Bearer ' + self.jwt
        return h

    def login(self):
        body = {'clientcode': self.cfg['client_code'], 'password': str(self.cfg['pin']),
                'totp': self._totp(self.cfg['totp_secret'])}
        r = self.requests.post(self.LOGIN, headers=self._headers(), json=body, timeout=15)
        j = r.json()
        if not j.get('status'):
            raise RuntimeError('Angel login failed: ' + json.dumps(j.get('message') or j))
        self.jwt = j['data']['jwtToken'].replace('Bearer ', '')
        print('[live] Angel login OK for', self.cfg['client_code'], flush=True)

    def load_tokens(self, symbols):
        print('[live] downloading scrip master…', flush=True)
        data = self.requests.get(self.SCRIP, timeout=60).json()
        want = set(symbols)
        for row in data:
            if row.get('exch_seg') == 'NSE' and str(row.get('symbol', '')).endswith('-EQ'):
                base = row['symbol'][:-3]
                if base in want:
                    self.tokens[base] = row['token']
        print('[live] resolved tokens for %d/%d symbols' % (len(self.tokens), len(want)), flush=True)

    def ltps(self, symbols):
        toks = [self.tokens[s] for s in symbols if s in self.tokens]
        if not toks:
            return {}
        body = {'mode': 'LTP', 'exchangeTokens': {'NSE': toks}}
        r = self.requests.post(self.QUOTE, headers=self._headers(auth=True), json=body, timeout=15)
        j = r.json()
        out = {}
        tok2sym = {v: k for k, v in self.tokens.items()}
        for row in (j.get('data', {}) or {}).get('fetched', []) or []:
            sym = tok2sym.get(str(row.get('symbolToken')))
            if sym:
                out[sym] = float(row.get('ltp') or 0)
        return out


# ---------------- P&L computation ----------------
def compute_pnl(mode, angel, positions, state):
    today = date.today()
    symbols = [p['symbol'] for p in positions]
    live = {}
    if mode == 'live' and angel:
        try:
            live = angel.ltps(symbols)
        except Exception as e:
            print('[live] quote error:', e, flush=True)

    rows = []
    tot_pnl = 0.0
    n_entered = 0
    # per-group tallies so RBI and HOLIDAY carry independent MTM
    groups = {}
    for p in positions:
        grp = p.get('group', 'RBI')
        g = groups.setdefault(grp, {'net_pnl': 0.0, 'positions_entered': 0, 'positions_total': 0})
        g['positions_total'] += 1

        entry_d = date.fromisoformat(p['entry_date'])
        exit_d = date.fromisoformat(p['exit_date'])
        st = status_for(entry_d, exit_d, today)
        dir_mul = 1 if p['direction'].upper() == 'LONG' else -1

        # State + price feed are namespaced per POSITION (group + event + symbol) so
        # the SAME symbol tracked in RBI vs HOLIDAY — or across two different holidays
        # — each gets its own entry price and its own tick walk and never clobbers
        # another position's captured entry.
        gkey = grp + '|' + str(p.get('event', '')) + '|' + p['symbol']
        ltp = live.get(p['symbol']) if mode == 'live' else mock_ltp(gkey, p['ref_price'], state)
        if not ltp:
            ltp = mock_ltp(gkey, p['ref_price'], state)  # fallback so UI stays alive

        entered = st in ('IN_PROGRESS', 'COMPLETED') or (st == 'ENTER_TODAY')
        ekey = 'entry:' + gkey
        entry_price = state.get(ekey)
        if entered and entry_price is None:
            entry_price = ltp            # capture entry once, at first tick after entry
            state[ekey] = entry_price
        if not entered:
            state.pop(ekey, None)
            entry_price = None

        pnl = pnl_pct = None
        if entry_price:
            pnl = round((ltp - entry_price) * p['lot'] * dir_mul, 2)
            pnl_pct = round((ltp - entry_price) / entry_price * 100 * dir_mul, 2)
            tot_pnl += pnl
            n_entered += 1
            g['net_pnl'] += pnl
            g['positions_entered'] += 1

        rows.append({**{k: p[k] for k in ('group', 'event', 'symbol', 'name', 'sector',
                                          'direction', 'window', 'entry_date', 'exit_date', 'lot')},
                     'status': st, 'ltp': ltp, 'entry_price': entry_price,
                     'pnl': pnl, 'pnl_pct': pnl_pct})
    save_state(state)
    rows.sort(key=lambda r: (r['pnl'] is None, -(r['pnl'] or 0)))
    for g in groups.values():
        g['net_pnl'] = round(g['net_pnl'], 2)
    return {
        'generated_at': datetime.now().isoformat(timespec='seconds'),
        'mode': mode,
        'groups': groups,
        'totals': {'net_pnl': round(tot_pnl, 2), 'positions_entered': n_entered,
                   'positions_total': len(positions)},
        'positions': rows,
    }


# ---------------- HTTP server ----------------
class Handler(BaseHTTPRequestHandler):
    def _send(self, obj, code=200):
        body = json.dumps(obj).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path.startswith('/api/health'):
            return self._send({'ok': True, 'mode': self.server.mode})
        if self.path.startswith('/api/pnl'):
            try:
                data = compute_pnl(self.server.mode, self.server.angel,
                                   self.server.positions, self.server.state)
                return self._send(data)
            except Exception as e:
                return self._send({'error': str(e)}, 500)
        return self._send({'error': 'not found'}, 404)

    def log_message(self, *a):
        pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--live', action='store_true', help='use Angel SmartAPI (needs smartapi_config.json)')
    ap.add_argument('--test', action='store_true', help='verify Angel login + print a few live LTPs, then exit')
    ap.add_argument('--port', type=int, default=PORT)
    args = ap.parse_args()

    anchor, rbi_positions = load_positions()
    holiday_positions = load_holiday_positions()
    positions = (rbi_positions or []) + (holiday_positions or [])
    if not positions:
        print('No positions found in', DATA); sys.exit(1)
    print('Loaded %d RBI + %d Holiday = %d playbook positions (MPC anchor %s)'
          % (len(rbi_positions or []), len(holiday_positions or []), len(positions), anchor))

    if args.test:
        if not os.path.exists(CONFIG_FILE):
            print('ERROR: smartapi_config.json not found — create it first (see the file header).'); sys.exit(1)
        cfg = json.load(open(CONFIG_FILE, encoding='utf-8'))
        missing = [k for k in ('api_key', 'client_code', 'pin', 'totp_secret') if not cfg.get(k)]
        if missing:
            print('ERROR: smartapi_config.json is missing keys:', missing); sys.exit(1)
        a = Angel(cfg)
        try:
            a.login()
            syms = [p['symbol'] for p in positions]
            a.load_tokens(syms)
            sample = syms[:5]
            q = a.ltps(sample)
            print('\n=== LIVE QUOTE TEST ===')
            for s in sample:
                print('  %-12s LTP = %s' % (s, q.get(s, 'NO DATA (market closed or token not found?)')))
            print('\n✅ Angel SmartAPI connection works. Run:  python live_pnl_server.py --live')
        except Exception as e:
            print('\n❌ Angel connection FAILED:', e)
            print('   Check api_key / client_code / MPIN / totp_secret, and that TOTP is enabled.')
        return

    mode = 'mock'
    angel = None
    if args.live:
        if not os.path.exists(CONFIG_FILE):
            print('smartapi_config.json not found — staying in MOCK mode.')
        else:
            cfg = json.load(open(CONFIG_FILE, encoding='utf-8'))
            angel = Angel(cfg)
            try:
                angel.login()
                angel.load_tokens([p['symbol'] for p in positions])
                mode = 'live'
            except Exception as e:
                print('[live] setup failed, falling back to MOCK:', e)
                angel = None

    srv = ThreadingHTTPServer(('0.0.0.0', args.port), Handler)
    srv.mode = mode
    srv.angel = angel
    srv.positions = positions
    srv.state = load_state()
    print('LIVE P&L SERVER (%s mode) -> http://localhost:%d/api/pnl' % (mode.upper(), args.port))
    print('Press Ctrl+C to stop.')
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print('\nStopped.')


if __name__ == '__main__':
    main()
