"""
무매v4.0 by ryanp0w
- 구글 시트 백엔드 + URL 사용자 인증 (?user=xxx)
- 4탭 (포트폴리오/매매이력/정산이력/슬롯 관리)
- 슬롯 없을 때 SOXL 종가 + 환율 표시
"""
import streamlit as st
import pandas as pd
import yfinance as yf
import requests
from datetime import datetime, date, timedelta
import json
import uuid
import gspread
from google.oauth2 import service_account
import extra_streamlit_components as stx

st.set_page_config(
    page_title="무매v4.0 by ryanp0w",
    page_icon="🚀",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ==========================================
# URL 사용자 인증 (+ 쿠키 자동 기억)
# ==========================================
cookie_manager = stx.CookieManager(key='mumae_cookie_manager')

query_params = st.query_params
user = query_params.get('user', None)

# 대소문자 통일 (모두 소문자로)
if user:
    user_lower = user.lower()
    if user != user_lower:
        st.query_params['user'] = user_lower
        st.rerun()
    user = user_lower

# URL에 user 있으면 → 쿠키에 저장 (다음 PWA/접속 시 자동 사용)
if user:
    cookie_manager.set('mumae_user', user, key='save_url_cookie')

# URL 없으면 → 쿠키에서 가져와 redirect 시도
if not user:
    saved = cookie_manager.get('mumae_user')
    if saved and isinstance(saved, str) and 4 <= len(saved) <= 50 and saved.replace('_', '').isalnum():
        st.query_params['user'] = saved.lower()
        st.rerun()

# 그래도 user 없으면 로그인 폼
if not user:
    st.markdown("# 🚀 무매v4.0 by ryanp0w")
    st.markdown("### 🔐 사용자 코드 입력")
    with st.form("login_form"):
        code = st.text_input("사용자 코드", placeholder="예: 23xfkdvc")
        submitted = st.form_submit_button("접속", width='stretch')
        if submitted:
            code = code.strip().lower()  # 대소문자 통일
            if code and code.replace('_', '').isalnum() and 4 <= len(code) <= 50:
                cookie_manager.set('mumae_user', code, key='save_login_cookie')
                st.query_params['user'] = code
                st.rerun()
            else:
                st.error("4-50자 영문/숫자/_ 만 사용 가능")
    st.caption("💡 한 번 접속하면 쿠키에 저장돼서 다음부터 자동 인증")
    st.stop()

if not user.replace('_', '').isalnum() or not (4 <= len(user) <= 50):
    st.title("🔒 잘못된 사용자 코드")
    st.markdown("사용자 코드가 올바르지 않습니다.")
    st.stop()

st.session_state.user = user

# ==========================================
# 구글 시트 연결
# ==========================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SLOT_HEADERS = ['id', 'name', 'ticker', 'capital', 'split', 'target_profit',
                'T', 'mode', 'reverse_entered_at', 'manual_reverse_star',
                'status', 'created_at', 'completed_at',
                'extra_buy_levels', 'extra_buy_pct_step']
TX_HEADERS = ['slot_id', 'date', 'type', 'price', 'qty', 'mode']

@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    creds = service_account.Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=SCOPES,
    )
    client = gspread.authorize(creds)
    return client.open_by_key(st.secrets["sheet"]["sheet_id"])

def get_or_create_ws(name, headers):
    ss = get_spreadsheet()
    # case-insensitive 검색: 대소문자 다른 같은 이름 시트가 있으면 그걸 사용
    for ws in ss.worksheets():
        if ws.title.lower() == name.lower():
            return ws
    # 없으면 새로 생성
    ws = ss.add_worksheet(title=name, rows=1000, cols=len(headers))
    ws.update(values=[headers], range_name='A1')
    return ws

def load_user_data(u):
    slots_ws = get_or_create_ws(f"{u}_slots", SLOT_HEADERS)
    txs_ws = get_or_create_ws(f"{u}_txs", TX_HEADERS)
    slots_rows = slots_ws.get_all_records()
    txs_rows = txs_ws.get_all_records()

    slots = {}
    for r in slots_rows:
        if not r.get('id'):
            continue
        sid = str(r['id'])
        slots[sid] = {
            'id': sid,
            'name': str(r.get('name', '')),
            'ticker': str(r.get('ticker', '')).upper(),
            'capital': float(r.get('capital') or 0),
            'split': int(r.get('split') or 40),
            'target_profit': float(r.get('target_profit') or 10),
            'T': float(r.get('T') or 0),
            'mode': str(r.get('mode') or 'normal'),
            'reverse_entered_at': (str(r['reverse_entered_at']) if r.get('reverse_entered_at') else None),
            'manual_reverse_star': (float(r['manual_reverse_star']) if r.get('manual_reverse_star') else None),
            'status': str(r.get('status', 'active')),
            'created_at': str(r.get('created_at', '')),
            'completed_at': (str(r['completed_at']) if r.get('completed_at') else None),
            'extra_buy_levels': int(r.get('extra_buy_levels') or 3),
            'extra_buy_pct_step': float(r.get('extra_buy_pct_step') or 5.0),
            'transactions': [],
        }
    for r in txs_rows:
        sid = str(r.get('slot_id', ''))
        if sid in slots:
            try:
                slots[sid]['transactions'].append({
                    'date': str(r['date']),
                    'type': str(r['type']),
                    'price': float(r['price']),
                    'qty': int(r['qty']),
                    'mode': str(r.get('mode') or 'normal'),
                })
            except Exception:
                continue
    return slots

def save_user_data(u, slots):
    slots_ws = get_or_create_ws(f"{u}_slots", SLOT_HEADERS)
    txs_ws = get_or_create_ws(f"{u}_txs", TX_HEADERS)

    slot_rows = [SLOT_HEADERS]
    for s in slots.values():
        slot_rows.append([
            s['id'], s['name'], s['ticker'],
            s['capital'], s['split'], s['target_profit'],
            s['T'], s.get('mode', 'normal'),
            s.get('reverse_entered_at') or '',
            s.get('manual_reverse_star') or '',
            s['status'], s['created_at'],
            s.get('completed_at') or '',
            s.get('extra_buy_levels', 3),
            s.get('extra_buy_pct_step', 5.0),
        ])
    slots_ws.clear()
    slots_ws.update(values=slot_rows, range_name='A1')

    tx_rows = [TX_HEADERS]
    for s in slots.values():
        for t in s['transactions']:
            tx_rows.append([
                s['id'], t['date'], t['type'],
                t['price'], t['qty'], t.get('mode', 'normal'),
            ])
    txs_ws.clear()
    txs_ws.update(values=tx_rows, range_name='A1')

def persist():
    try:
        save_user_data(st.session_state.user, st.session_state.slots)
    except Exception as e:
        st.error(f"시트 저장 실패: {e}")

# ==========================================
# 초기 로드
# ==========================================
if 'slots' not in st.session_state:
    with st.spinner("데이터 불러오는 중..."):
        try:
            st.session_state.slots = load_user_data(user)
        except Exception as e:
            st.error(f"데이터 로드 실패: {e}")
            st.session_state.slots = {}

if 'selected_slot_id' not in st.session_state:
    st.session_state.selected_slot_id = None

# ==========================================
# 계산
# ==========================================
def calc_holdings(transactions, capital):
    h, a, used, rev = 0, 0.0, 0.0, 0.0
    for t in transactions:
        if t['type'] == 'buy':
            new_total = a * h + t['price'] * t['qty']
            h += t['qty']
            a = new_total / h if h > 0 else 0
            used += t['price'] * t['qty']
        else:
            h -= t['qty']
            rev += t['price'] * t['qty']
            if h <= 0:
                h, a = 0, 0
    cash = capital - used + rev
    return h, a, cash, used, rev

def calc_realized(transactions):
    h, a, realized = 0, 0.0, 0.0
    for t in transactions:
        if t['type'] == 'buy':
            new_total = a * h + t['price'] * t['qty']
            h += t['qty']
            a = new_total / h if h > 0 else 0
        else:
            realized += (t['price'] - a) * t['qty']
            h -= t['qty']
            if h <= 0:
                h, a = 0, 0
    return realized

def calc_star_pct(T, ticker, split):
    # 별% = base × (1 − 2T/N)  (base: TQQQ=15, SOXL/기타=20)
    # 검증: N=20 SOXL→20−2T, N=40 SOXL→20−T, N=20 TQQQ→15−1.5T, N=40 TQQQ→15−0.75T
    # 임의 분할수(예: 30)도 동일 공식으로 일반화 처리
    base = 15 if ticker == 'TQQQ' else 20
    return base * (1 - 2 * T / split) if split else 0

def calc_one_buy(cash, T, split):
    d = split - T
    return cash / d if d > 0 else 0

def calc_extra_buy_tiers(base_price, target_budget, existing_cum_qty, num_tiers, pct_step):
    """
    별지점/평단 등 기존 매수 지점 아래로 추가 LOC 매수 지점을 계산.
    각 지점까지 가격이 떨어졌을 때, 그 지점까지의 누적 매수금액이
    target_budget(1회매수액)에 최대한 가깝도록 수량을 누적 계산.

    예) 1회매수액 600, 평단 200, 별지점 220 (전반전, 각 300씩 배정)
        기존: 별지점 220에서 1주(300/220=1.36→1), 평단 200에서 1주(300/200=1.5→1)
        기존 누적 수량 = 2주, 기존 누적 사용액 = 220+200 = 420 (실제 체결가 기준으로는 더 적을 수 있음)
        추가 지점(평단 아래 -5%=190): 600/190=3.15→3주 누적 목표. 기존 2주 빼면 +1주
        추가 지점(-10%=180): 600/180=3.33→3주. 기존 3주면 +0주
        추가 지점(-15%=170): 600/170=3.52→3주. 추가 없음
        → 실제 종가 150이면 위 지점들 다 체결되어 총 3~4주 매수, 300+300=600 예산 최대한 사용
    """
    tiers = []
    if base_price <= 0 or target_budget <= 0:
        return tiers
    cum_qty = existing_cum_qty
    price = base_price
    for _ in range(num_tiers):
        price = round(price * (1 - pct_step / 100), 2)
        if price <= 0:
            break
        total_qty_at_price = int(target_budget // price)
        qty_here = total_qty_at_price - cum_qty
        if qty_here < 0:
            qty_here = 0
        if qty_here > 0:
            tiers.append({'price': price, 'qty': qty_here})
            cum_qty += qty_here
    return tiers

def resolve_top_buy_price(star_buy_price, close_price, big_pct=12.0, reject_pct=20.0):
    """매수점 표의 최상단(★or큰수) 가격을 결정.
    별지점(−0.01 적용가)이 현재가 대비 약 reject_pct%를 초과해 위에 있으면
    증권사가 현재가 ±20% 초과 LOC 주문을 거부하는 것을 회피하기 위해
    큰수매수(현재가 ×(1+big_pct%))로 대체한다(PDF1 5-6쪽 헤더 "★or큰수", 8쪽 큰수매수).
    수량 배정은 호출부에서 배정액//사용가격으로 유지.
    현재가 조회 실패 시 별지점을 그대로 유지(안전 폴백).
    반환: (사용가격, 라벨)
    """
    if close_price and star_buy_price > close_price * (1 + reject_pct / 100):
        return round(close_price * (1 + big_pct / 100), 2), '큰수매수'
    return star_buy_price, '★ 별지점'

def get_phase(T, split):
    # 소진: T > N−1 (T=N−1은 아직 후반전). 리버스 전환조건(new_T > split−1)과 일관
    if T > split - 1:
        return 'sojin'
    elif T >= split / 2:
        return 'late'
    return 'early'

def phase_label(p):
    return {'early': '전반전', 'late': '후반전', 'sojin': '소진'}[p]

# ==========================================
# 리버스 모드 헬퍼
# ==========================================
@st.cache_data(ttl=3600, show_spinner=False)
def fetch_reverse_star(ticker):
    """별지점 = 직전 5거래일 종가 평균"""
    try:
        t = yf.Ticker(ticker)
        d = t.history(period="10d")
        if not d.empty and len(d) >= 5:
            return float(d['Close'].iloc[-5:].mean())
    except Exception:
        pass
    return None

def calc_reverse_sell_qty(holdings, split):
    """리버스 매도수량 = 직전보유 ÷ (N/2), 내림 (20분할→1/10, 40분할→1/20, 30분할→1/15)"""
    divisor = split // 2 if split else 1
    return holdings // divisor if divisor else 0

def is_first_reverse_sell(slot):
    """최근 리버스 진입 이후 첫 매도가 아직 없는지"""
    entered_at = slot.get('reverse_entered_at')
    # 진입 시점 없으면 (구버전 데이터 호환) 전체 리버스 매도 이력으로 판단
    if not entered_at:
        for t in slot['transactions']:
            if t.get('mode') == 'reverse' and t['type'] == 'sell':
                return False
        return True
    # 진입 시점 이후의 리버스 매도만 카운트
    for t in slot['transactions']:
        if t.get('mode') == 'reverse' and t['type'] == 'sell':
            if t['date'] >= entered_at:
                return False
    return True

def check_reverse_exit(slot, close_price):
    """종료조건: 종가 > 평단 × (1 - X%/100), X=15(TQQQ) or 20(SOXL)"""
    if not close_price:
        return False, None
    h, a, _, _, _ = calc_holdings(slot['transactions'], slot['capital'])
    if not a:
        return False, None
    pct = 15 if slot['ticker'] == 'TQQQ' else 20
    threshold = a * (1 - pct / 100)
    return close_price > threshold, threshold

def calc_reverse_T_buy(T, split):
    """리버스 매수시 T값: T + (split-T)×0.25"""
    return T + (split - T) * 0.25

def calc_reverse_T_sell(T, split):
    """리버스 매도시 T값: T × (1 − 2/N) (20분할→0.9, 40분할→0.95, 30분할→0.9333)"""
    return T * (1 - 2 / split) if split else T

# ==========================================
# yfinance
# ==========================================
@st.cache_data(ttl=600, show_spinner=False)
def fetch_close(ticker):
    """현재가/종가: yfinance fast_info → history → stooq 실시간 → stooq 일별 순"""
    # 1. yfinance fast_info (실시간 우선)
    try:
        t = yf.Ticker(ticker)
        try:
            fi = t.fast_info
            p = fi.get('last_price') if hasattr(fi, 'get') else getattr(fi, 'last_price', None)
            if p and p > 0:
                return float(p)
        except Exception:
            pass
        # 2. yfinance history
        d = t.history(period="5d")
        if not d.empty:
            return float(d['Close'].iloc[-1])
    except Exception:
        pass
    # 3. stooq 실시간 CSV
    try:
        url = f"https://stooq.com/q/l/?s={ticker.lower()}.us&f=sd2t2ohlcv&h&e=csv"
        r = requests.get(url, timeout=5)
        if r.status_code == 200:
            lines = r.text.strip().split('\n')
            if len(lines) >= 2:
                parts = lines[1].split(',')
                if len(parts) >= 7 and parts[6] not in ('N/D', ''):
                    try:
                        val = float(parts[6])
                        if val > 0:
                            return val
                    except ValueError:
                        pass
    except Exception:
        pass
    # 4. stooq 일별 CSV (가장 안정적인 백업)
    try:
        url2 = f"https://stooq.com/q/d/l/?s={ticker.lower()}.us&i=d"
        r = requests.get(url2, timeout=5)
        if r.status_code == 200:
            lines = r.text.strip().split('\n')
            if len(lines) >= 2:
                last_line = lines[-1].strip()
                parts = last_line.split(',')
                # Date,Open,High,Low,Close,Volume
                if len(parts) >= 5 and parts[4]:
                    try:
                        val = float(parts[4])
                        if val > 0:
                            return val
                    except ValueError:
                        pass
    except Exception:
        pass
    return None

@st.cache_data(ttl=3600, show_spinner=False)
def fetch_rate():
    """환율: yfinance 우선, 실패시 frankfurter API"""
    # 1. yfinance USDKRW=X
    try:
        d = yf.Ticker("USDKRW=X").history(period="2d")
        if not d.empty:
            return float(d['Close'].iloc[-1])
    except Exception:
        pass
    # 2. yfinance KRW=X
    try:
        d = yf.Ticker("KRW=X").history(period="2d")
        if not d.empty:
            return float(d['Close'].iloc[-1])
    except Exception:
        pass
    # 3. frankfurter.app (무료 환율 API, 키 불필요)
    try:
        r = requests.get("https://api.frankfurter.app/latest?from=USD&to=KRW", timeout=5)
        if r.status_code == 200:
            return float(r.json()['rates']['KRW'])
    except Exception:
        pass
    return None

# ==========================================
# 슬롯 헬퍼
# ==========================================
def create_slot(name, ticker, capital, split, target_profit):
    sid = str(uuid.uuid4())[:8]
    st.session_state.slots[sid] = {
        'id': sid, 'name': name, 'ticker': ticker,
        'capital': float(capital), 'split': int(split),
        'target_profit': float(target_profit),
        'T': 0.0, 'transactions': [],
        'mode': 'normal',
        'reverse_entered_at': None,
        'manual_reverse_star': None,
        'status': 'active',
        'created_at': str(date.today()),
        'completed_at': None,
        'extra_buy_levels': 3,
        'extra_buy_pct_step': 5.0,
    }
    return sid

def active_slots():
    return [s for s in st.session_state.slots.values() if s.get('status') == 'active']

def completed_slots():
    return [s for s in st.session_state.slots.values() if s.get('status') == 'completed']

def current_slot():
    sid = st.session_state.selected_slot_id
    if sid and sid in st.session_state.slots and st.session_state.slots[sid].get('status') == 'active':
        return st.session_state.slots[sid]
    acts = active_slots()
    if acts:
        st.session_state.selected_slot_id = acts[0]['id']
        return acts[0]
    return None

def big_number(label, value_str, color):
    st.markdown(
        f"<div style='text-align: center; padding: 4px 0;'>"
        f"<div style='font-size: 13px; color: #888; margin-bottom: 4px;'>{label}</div>"
        f"<div style='font-size: 40px; font-weight: bold; color: {color}; line-height: 1.1;'>{value_str}</div>"
        f"</div>",
        unsafe_allow_html=True
    )

def big_number_row(items):
    """items: [(label, value_str, color), ...] - 항상 가로로 강제 배치 (HTML flex, Streamlit columns 아님)"""
    cells = ""
    for label, value_str, color in items:
        cells += (
            f"<div style='flex: 1; text-align: center; padding: 4px 2px; min-width: 0;'>"
            f"<div style='font-size: 13px; color: #888; margin-bottom: 4px;'>{label}</div>"
            f"<div style='font-size: 34px; font-weight: bold; color: {color}; line-height: 1.15; "
            f"white-space: nowrap;'>{value_str}</div>"
            f"</div>"
        )
    st.markdown(
        f"<div style='display: flex; width: 100%;'>{cells}</div>",
        unsafe_allow_html=True
    )

def metric_grid(items, columns=2):
    """items: [(label, value_str), ...] - CSS grid로 화면 폭 상관없이 고정 열 유지"""
    cells = ""
    for label, value_str in items:
        cells += (
            f"<div style='padding: 6px 4px; min-width: 0;'>"
            f"<div style='font-size: 13px; color: #888; margin-bottom: 2px;'>{label}</div>"
            f"<div style='font-size: 21px; font-weight: 600; overflow-wrap: break-word;'>{value_str}</div>"
            f"</div>"
        )
    st.markdown(
        f"<div style='display: grid; grid-template-columns: repeat({columns}, 1fr); "
        f"column-gap: 8px; row-gap: 4px;'>{cells}</div>",
        unsafe_allow_html=True
    )

def trans_df(transactions):
    if not transactions:
        return pd.DataFrame()
    rows = []
    h, a = 0, 0.0
    for t in sorted(transactions, key=lambda x: x['date']):
        if t['type'] == 'buy':
            new_total = a * h + t['price'] * t['qty']
            h += t['qty']
            a = new_total / h if h > 0 else 0
        else:
            h -= t['qty']
            if h <= 0:
                h, a = 0, 0
        rows.append({
            '날짜': t['date'],
            '구분': '매수' if t['type'] == 'buy' else '매도',
            '수량': t['qty'],
            '가격': round(t['price'], 2),
            '평단': round(a, 2) if a else 0,
            '금액': round(t['price'] * t['qty'], 2),
        })
    return pd.DataFrame(rows).iloc[::-1].reset_index(drop=True)

# ==========================================
# 사이드바: 사용자 정보 + 환율 + 안내
# ==========================================
def render_sidebar():
    with st.sidebar:
        st.markdown(f"👤 **사용자**: `{user}`")
        st.divider()
        st.markdown("**📊 시세**")
        soxl = fetch_close("SOXL")
        tqqq = fetch_close("TQQQ")
        rate = fetch_rate()
        st.markdown(f"SOXL: **${soxl:.2f}**" if soxl else "SOXL: 조회 실패")
        st.markdown(f"TQQQ: **${tqqq:.2f}**" if tqqq else "TQQQ: 조회 실패")
        st.markdown(f"환율: **{rate:,.2f}원**" if rate else "환율: 조회 실패")
        if st.button("🔄 시세 새로고침", width='stretch', key='refresh_cache'):
            st.cache_data.clear()
            st.rerun()
        st.divider()
        st.caption("슬롯 추가·관리·백업은 메인의 **⚙️ 슬롯 관리** 탭에서")

# ==========================================
# 헤더
# ==========================================
def render_header():
    st.markdown("# 🚀 무매v4.0 by ryanp0w")
    acts = active_slots()

    if not acts:
        # 슬롯이 없을 때: SOXL 종가 + 환율 표시 + 안내
        st.info("아직 슬롯이 없어요. 아래 **⚙️ 슬롯 관리** 탭에서 슬롯을 추가하세요.")
        soxl_close = fetch_close("SOXL")
        rate = fetch_rate()
        with st.container(border=True):
            cc = st.columns(2)
            cc[0].metric(
                "📊 SOXL 현재가",
                f"${soxl_close:.2f}" if soxl_close else "조회 실패"
            )
            cc[1].metric(
                "💱 USD/KRW 환율",
                f"{rate:,.2f}원" if rate else "조회 실패"
            )
        return None

    # 슬롯이 여러 개면 selector
    if len(acts) > 1:
        opts = [s['name'] for s in acts]
        sids = [s['id'] for s in acts]
        if st.session_state.selected_slot_id not in sids:
            st.session_state.selected_slot_id = sids[0]
        ci = sids.index(st.session_state.selected_slot_id)
        i = st.selectbox("슬롯", range(len(opts)),
                         format_func=lambda x: opts[x], index=ci)
        if sids[i] != st.session_state.selected_slot_id:
            st.session_state.selected_slot_id = sids[i]
            st.rerun()
    return current_slot()

# ==========================================
# 탭 1: 포트폴리오
# ==========================================
def tab_portfolio(s):
    # === 계산 ===
    h, a, cash, used, rev = calc_holdings(s['transactions'], s['capital'])
    cls = fetch_close(s['ticker'])
    realized = calc_realized(s['transactions'])
    unreal = (cls - a) * h if (cls and a) else 0
    total_profit = realized + unreal

    val = h * cls if cls else 0
    total = cash + val
    # 수익률: 실제 투입한 금액 (used = 누적 매수 금액) 대비
    total_pct = (total_profit / used * 100) if used else 0

    T, spl = s['T'], s['split']
    one_buy = calc_one_buy(cash, T, spl)
    phase = get_phase(T, spl)
    mode = s.get('mode', 'normal')
    star_pct = calc_star_pct(T, s['ticker'], spl)
    star_price = a * (1 + star_pct / 100) if a else 0

    # === 1. 오늘의 주문 (최상단) ===
    if s['ticker'] not in ('TQQQ', 'SOXL'):
        st.caption(f"ℹ️ 방법론은 TQQQ/SOXL 기준 — 이 티커({s['ticker']})는 SOXL(20%) 규칙으로 계산됨")
    if mode == 'reverse':
        st.subheader("🔄 리버스 모드 — 오늘의 주문")
        render_reverse_orders(s, h, a, cash, cls)
    else:
        st.subheader("🌙 오늘의 주문")
        if h == 0:
            st.markdown("**🚀 처음 매수**")
            if cls:
                big = round(cls * 1.12, 2)
                qty = int(one_buy // big) if big > 0 else 0
                first_rows = [{"구분": "큰수매수 LOC (+12%)", "가격": f"${big:.2f}", "수량": f"{qty}주"}]
                # 첫날 급락 대비: 큰수 아래로 하향 LOC 사다리(1회매수액 최대 소진)
                extra_n = s.get('extra_buy_levels', 3)
                extra_step = s.get('extra_buy_pct_step', 5.0)
                for t in calc_extra_buy_tiers(big, one_buy, qty, extra_n, extra_step):
                    first_rows.append({"구분": "하향 LOC 매수", "가격": f"${t['price']}", "수량": f"{t['qty']}주"})
                st.dataframe(pd.DataFrame(first_rows), hide_index=True, width='stretch')
                st.caption(f"종가 ${cls:.2f} 기준 · 1회매수액 ${one_buy:.2f}")
            else:
                st.warning("종가 로드 실패. 수동으로 +10~15% 위에 LOC + 아래로 하향 LOC 사다리")
        else:
            bp = round(star_price - 0.01, 2)
            spct = 15 if s['ticker'] == 'TQQQ' else 20
            target_sell = round(a * (1 + spct / 100), 2)
            quarter = h // 4
            rest = h - quarter

            st.markdown(f"**🛒 매수** · {phase_label(phase)}")
            extra_n = s.get('extra_buy_levels', 3)
            extra_step = s.get('extra_buy_pct_step', 5.0)
            buy_rows = []
            # 별지점이 현재가 대비 20% 초과로 위에 있으면 큰수매수로 대체(증권사 주문거부 회피).
            # 이 경우에도 사용자가 원래 별지점 기준 매수표를 확인할 수 있도록 토글을 제공한다.
            resolved_top_price, resolved_top_label = resolve_top_buy_price(bp, cls)
            is_big_buy = resolved_top_label == '큰수매수'
            if is_big_buy:
                show_big_buy = st.toggle(
                    "큰수매수 표 보기",
                    value=True,
                    key=f"show_big_buy_{s['id']}",
                    help="끔: 원래 별지점 기준 매수표를 표시합니다."
                )
                if show_big_buy:
                    top_price, top_label = resolved_top_price, resolved_top_label
                    st.caption(
                        f"주문거부 방지 큰수매수 ${resolved_top_price:.2f} 적용 중 "
                        f"· 원래 별지점 매수 ${bp:.2f}"
                    )
                else:
                    top_price, top_label = bp, '★ 별지점'
                    st.caption(
                        f"원래 별지점 매수표 표시 중 (${bp:.2f}) "
                        f"· 큰수매수 기준 ${resolved_top_price:.2f}"
                    )
            else:
                top_price, top_label = resolved_top_price, resolved_top_label
            if phase == 'early':
                half = one_buy / 2
                qs = int(half // top_price) if top_price > 0 else 0
                qa = int(half // a) if a > 0 else 0
                buy_rows.append({"구분": top_label, "가격": f"${top_price}", "수량": f"{qs}주"})
                buy_rows.append({"구분": "평단", "가격": f"${a:.2f}", "수량": f"{qa}주"})
                base_for_extra = min(top_price, a) if a > 0 else top_price
                existing_cum = qs + qa
            elif phase == 'late':
                qs = int(one_buy // top_price) if top_price > 0 else 0
                buy_rows.append({"구분": top_label, "가격": f"${top_price}", "수량": f"{qs}주"})
                base_for_extra = top_price
                existing_cum = qs
            else:
                base_for_extra = None
                existing_cum = 0

            if phase in ('early', 'late') and base_for_extra:
                extra_tiers = calc_extra_buy_tiers(
                    base_for_extra, one_buy, existing_cum, extra_n, extra_step
                )
                for t in extra_tiers:
                    buy_rows.append({"구분": "추가 LOC 매수", "가격": f"${t['price']}", "수량": f"{t['qty']}주"})

            if phase == 'sojin':
                st.error("⚠️ 소진 도달 · 다음 매수 저장 시 자동 리버스 전환")
            else:
                st.dataframe(pd.DataFrame(buy_rows), hide_index=True, width='stretch')
                st.caption(f"1회매수액 ${one_buy:.2f}")

            st.markdown("**💰 매도**")
            sell_df = pd.DataFrame([
                {"구분": "★ 별지점", "가격": f"${star_price:.2f}", "수량": f"{quarter}주"},
                {"구분": f"{spct}% 지정가", "가격": f"${target_sell}", "수량": f"{rest}주"},
            ])
            st.dataframe(sell_df, hide_index=True, width='stretch')

    st.divider()

    # === 2. 투자 수익률 / 손익 (핵심 지표, 강제 2열) ===
    pcolor = '#ef4444' if total_pct >= 0 else '#3b82f6'
    big_number_row([
        ("투자 수익률", f"{total_pct:+.2f}%", pcolor),
        ("투자 손익", f"{total_profit:+.2f} USD", pcolor),
    ])
    st.markdown("<br>", unsafe_allow_html=True)

    # === 3. 나머지 현황 (강제 2열 그리드) ===
    st.subheader(f"📌 {s['name']}")
    metric_grid([
        ("1회 매수금", f"${one_buy:.0f}"),
        ("🔄 모드" if mode == 'reverse' else "단계",
         "리버스" if mode == 'reverse' else phase_label(phase)),
        ("진행도", f"{T:.2f}/{spl}T"),
        ("보유", f"{h}주"),
        ("평단", f"${a:.2f}" if a else "-"),
        ("현재가", f"${cls:.2f}" if cls else "-"),
        ("잔금", f"${cash:,.0f}"),
        ("운용 종목수", f"{len(active_slots())}개"),
    ], columns=3)

def render_reverse_orders(s, h, a, cash, cls):
    """리버스 모드 주문 가이드"""
    # 별지점: 수동 입력값 우선, 없으면 yfinance 자동
    star = s.get('manual_reverse_star')
    star_source = "수동 입력"
    if not star:
        star = fetch_reverse_star(s['ticker'])
        star_source = "직전 5거래일 평균 (자동)"

    is_first = is_first_reverse_sell(s)
    exit_ok, threshold = check_reverse_exit(s, cls)
    pct = 15 if s['ticker'] == 'TQQQ' else 20

    # 상태
    status_rows = [{"항목": "별지점", "값": f"${star:.2f} ({star_source})" if star else "조회 실패"}]
    if a and threshold:
        status_rows.append({"항목": "종료조건", "값": f"종가 > ${threshold:.2f} (평단 -{pct}%)"})
    st.dataframe(pd.DataFrame(status_rows), hide_index=True, width='stretch')
    if exit_ok:
        st.success("✅ 종료조건 충족! 슬롯 관리 탭에서 일반모드 복귀 가능")

    # 매도
    sell_qty = calc_reverse_sell_qty(h, s['split'])
    divisor = s['split'] // 2 if s['split'] else 1
    if is_first:
        st.markdown("**💰 처음 매도** · MOC 무조건 체결")
        df = pd.DataFrame([{"구분": "MOC 매도", "가격": "시장가", "수량": f"{sell_qty}주"}])
        st.dataframe(df, hide_index=True, width='stretch')
        st.caption(f"보유 {h}주 × 1/{divisor}")
    else:
        st.markdown("**💰 매도** · 별지점 위 LOC")
        df = pd.DataFrame([{"구분": "★ 별지점 위", "가격": f"${star:.2f}" if star else "-", "수량": f"{sell_qty}주"}])
        st.dataframe(df, hide_index=True, width='stretch')
        st.caption(f"직전보유 × 1/{divisor}")

    # 매수 (첫날 제외)
    if not is_first:
        qbuy = cash / 4
        st.markdown("**🛒 쿼터매수** · 별지점 아래 LOC")
        df = pd.DataFrame([{"구분": "쿼터매수", "가격": f"${star:.2f} 아래" if star else "재량", "수량": f"${qbuy:.2f}"}])
        st.dataframe(df, hide_index=True, width='stretch')
        st.caption("잔금 ÷ 4, 가격대는 재량으로 분산")

# ==========================================
# 탭 2: 매매이력
# ==========================================
# ==========================================
# 매도 결과 모달
# ==========================================
@st.dialog("🏁 매도 결과")
def sell_result_dialog():
    info = st.session_state.get('pending_sell_dialog')
    if not info:
        return

    slot = st.session_state.slots.get(info['slot_id'])
    if not slot:
        return

    sell_type = info['sell_type']  # 'quarter', 'fixed', 'reverse', 'full'

    h, a, cash, used, rev = calc_holdings(slot['transactions'], slot['capital'])
    realized = calc_realized(slot['transactions'])
    pct = (realized / used * 100) if used else 0

    st.markdown(f"### {slot['name']} ({slot['ticker']})")

    # 이미지 분기
    if realized < 0:
        img_file = "images/sad.jpg"
    elif sell_type == 'full':
        img_file = "images/happy2.jpg"
    else:
        img_file = "images/happy1.jpg"

    try:
        st.image(img_file, width='stretch')
    except Exception:
        pass

    color = '#ef4444' if realized >= 0 else '#3b82f6'
    st.markdown(
        f"<div style='text-align: center; padding: 10px 0;'>"
        f"<div style='font-size: 14px; color: #888;'>실현 수익</div>"
        f"<div style='font-size: 36px; font-weight: bold; color: {color};'>"
        f"{realized:+.2f} USD</div>"
        f"<div style='font-size: 20px; color: {color}; margin-top: 5px;'>"
        f"({pct:+.2f}%)</div>"
        f"</div>",
        unsafe_allow_html=True
    )

    if st.button("확인", width='stretch', key="dialog_confirm"):
        del st.session_state['pending_sell_dialog']
        st.rerun()


def tab_history(s):
    cls = fetch_close(s['ticker'])
    mode = s.get('mode', 'normal')

    if mode == 'reverse':
        st.info("🔄 리버스 모드 매매 입력")

    st.subheader("📝 체결 입력")
    btab, stab = st.tabs(["🛒 매수", "💰 매도"])

    with btab:
        with st.form(f"buy_{s['id']}", clear_on_submit=True):
            fd = st.date_input("날짜", value=date.today(), key=f"bd_{s['id']}")
            cc = st.columns(2)
            fq = cc[0].number_input("수량", min_value=1, value=1, step=1, key=f"bq_{s['id']}")
            fp = cc[1].number_input("가격", min_value=0.01,
                                     value=float(cls) if cls else 50.0, step=0.01, key=f"bp_{s['id']}")

            if mode == 'reverse':
                # 리버스 매수: T + (split-T)×0.25 (한 옵션만)
                new_T = calc_reverse_T_buy(s['T'], s['split'])
                st.caption(f"리버스 쿼터매수: T {s['T']:.3f} → **{new_T:.3f}**")
                tx_mode = 'reverse'
            else:
                tc = st.radio("T값 변화", ["+1 (1회 매수)", "+0.5 (절반 매수)"], key=f"btc_{s['id']}")
                td = 1.0 if "+1" in tc else 0.5
                new_T = s['T'] + td
                st.caption(f"T: {s['T']:.3f} → **{new_T:.3f}**")
                tx_mode = 'normal'

            if st.form_submit_button("💾 매수 저장", width='stretch'):
                s['transactions'].append({
                    'date': str(fd), 'type': 'buy',
                    'price': float(fp), 'qty': int(fq),
                    'mode': tx_mode,
                })
                s['T'] = new_T
                # 일반→리버스 자동 전환
                if mode == 'normal' and new_T > s['split'] - 1:
                    s['mode'] = 'reverse'
                    s['reverse_entered_at'] = str(date.today())
                    s['manual_reverse_star'] = None  # 새 리버스 진입이니 별지점 초기화
                    persist()
                    st.warning(f"🔄 T값이 {s['split']-1}을 초과해 **리버스 모드로 자동 전환**됨")
                else:
                    persist()
                    st.success("매수 저장됨")
                st.rerun()

    with stab:
        h_cur, _, _, _, _ = calc_holdings(s['transactions'], s['capital'])
        with st.form(f"sell_{s['id']}", clear_on_submit=True):
            fd = st.date_input("날짜", value=date.today(), key=f"sd_{s['id']}")
            cc = st.columns(2)
            fq = cc[0].number_input("수량", min_value=1, value=1, step=1, key=f"sq_{s['id']}")
            fp = cc[1].number_input("가격", min_value=0.01,
                                     value=float(cls) if cls else 50.0, step=0.01, key=f"sp_{s['id']}")

            if mode == 'reverse':
                # 리버스 매도: T×0.9 or T×0.95 (한 옵션만)
                stp = st.radio("매도 종류",
                               ["일반 리버스 매도",
                                f"🏁 전액매도 ({h_cur}주, 슬롯 종료)"], key=f"sstp_{s['id']}")
                if "전액매도" in stp:
                    new_T = s['T']
                    st.warning(f"⚠️ 보유 전량 {h_cur}주를 매도하고 슬롯 종료됩니다.")
                else:
                    new_T = calc_reverse_T_sell(s['T'], s['split'])
                    st.caption(f"리버스 매도: T {s['T']:.3f} → **{new_T:.3f}** (×{0.9 if s['split']==20 else 0.95})")
                tx_mode = 'reverse'
            else:
                stp = st.radio("T값 변화",
                               ["쿼터매도 (×0.75)",
                                "지정가 매도 (×0.25)",
                                f"🏁 전액매도 ({h_cur}주, 슬롯 종료)"], key=f"sstp_{s['id']}")
                if "0.75" in stp:
                    new_T = s['T'] * 0.75
                elif "0.25" in stp:
                    new_T = s['T'] * 0.25
                else:
                    new_T = s['T']
                    st.warning(f"⚠️ 보유 전량 {h_cur}주를 매도하고 슬롯 종료됩니다. 수량 입력값은 무시됨.")
                st.caption(f"T: {s['T']:.3f} → **{new_T:.3f}**")
                st.caption("💡 별지점 LOC 매수가 같이 체결됐다면 매수 폼에서 따로 입력")
                tx_mode = 'normal'

            if st.form_submit_button("💾 매도 저장", width='stretch'):
                if "전액매도" in stp:
                    if h_cur <= 0:
                        st.error("보유 수량이 0이라 전액매도 불가")
                    else:
                        slot_id_to_show = s['id']
                        s['transactions'].append({
                            'date': str(fd), 'type': 'sell',
                            'price': float(fp), 'qty': int(h_cur),
                            'mode': tx_mode,
                        })
                        s['status'] = 'completed'
                        s['completed_at'] = str(date.today())
                        st.session_state.selected_slot_id = None
                        persist()
                        st.session_state['pending_sell_dialog'] = {
                            'slot_id': slot_id_to_show,
                            'sell_type': 'full',
                        }
                        st.rerun()
                else:
                    s['transactions'].append({
                        'date': str(fd), 'type': 'sell',
                        'price': float(fp), 'qty': int(fq),
                        'mode': tx_mode,
                    })
                    s['T'] = new_T
                    persist()
                    # 매도 종류 분류
                    if tx_mode == 'reverse':
                        sell_type_label = 'reverse'
                    elif "0.75" in stp:
                        sell_type_label = 'quarter'
                    else:
                        sell_type_label = 'fixed'
                    st.session_state['pending_sell_dialog'] = {
                        'slot_id': s['id'],
                        'sell_type': sell_type_label,
                    }
                    st.rerun()

    st.divider()
    st.subheader("📜 매매 이력")
    df = trans_df(s['transactions'])
    if df.empty:
        st.info("거래 없음")
    else:
        def style_row(row):
            color = '#ef444433' if row['구분'] == '매수' else '#3b82f633'
            return [f'background-color: {color}'] * len(row)
        st.dataframe(df.style.apply(style_row, axis=1), hide_index=True, width='stretch')

        tbq = sum(t['qty'] for t in s['transactions'] if t['type'] == 'buy')
        tsq = sum(t['qty'] for t in s['transactions'] if t['type'] == 'sell')
        tba = sum(t['price']*t['qty'] for t in s['transactions'] if t['type'] == 'buy')
        tsa = sum(t['price']*t['qty'] for t in s['transactions'] if t['type'] == 'sell')
        cc = st.columns(2)
        cc[0].metric("총 매수", f"{tbq}주 / ${tba:,.0f}")
        cc[1].metric("총 매도", f"{tsq}주 / ${tsa:,.0f}")

        with st.expander("🛠 수정/삭제"):
            edit_df = pd.DataFrame(s['transactions'])
            if not edit_df.empty:
                edited = st.data_editor(
                    edit_df, width='stretch', num_rows="dynamic",
                    column_config={
                        'date': st.column_config.TextColumn('날짜'),
                        'type': st.column_config.SelectboxColumn('구분', options=['buy', 'sell']),
                        'price': st.column_config.NumberColumn('가격', format="%.2f"),
                        'qty': st.column_config.NumberColumn('수량'),
                    },
                    key=f"ed_{s['id']}",
                )
                if st.button("💾 변경 저장", key=f"se_{s['id']}", width='stretch'):
                    nt = []
                    for _, r in edited.iterrows():
                        try:
                            nt.append({
                                'date': str(r['date']),
                                'type': r['type'],
                                'price': float(r['price']),
                                'qty': int(r['qty']),
                            })
                        except Exception:
                            continue
                    s['transactions'] = nt
                    persist()
                    st.warning("저장됨. T값은 자동 복원 안 됨")
                    st.rerun()

# ==========================================
# 탭 3: 정산이력
# ==========================================
def tab_settlement(s=None):
    """모든 슬롯(active + completed)의 매도 거래 통합 정산"""
    st.subheader("💰 정산 이력 (전체 슬롯)")
    cc = st.columns(2)
    start = cc[0].date_input("시작일", value=date.today() - timedelta(days=90), key="settlement_start")
    end = cc[1].date_input("종료일", value=date.today(), key="settlement_end")

    realized_list = []
    # 모든 슬롯 순회
    for slot in st.session_state.slots.values():
        h, a = 0, 0.0
        for t in sorted(slot['transactions'], key=lambda x: x['date']):
            if t['type'] == 'buy':
                new_total = a * h + t['price'] * t['qty']
                h += t['qty']
                a = new_total / h if h > 0 else 0
            else:
                r = (t['price'] - a) * t['qty']
                realized_list.append({
                    '날짜': t['date'],
                    '슬롯': slot['name'],
                    '종목': slot['ticker'],
                    '매도가': round(t['price'], 2),
                    '수량': t['qty'],
                    '당시평단': round(a, 2),
                    '실현수익': round(r, 2),
                })
                h -= t['qty']
                if h <= 0:
                    h, a = 0, 0

    try:
        filtered = [x for x in realized_list
                    if start <= datetime.fromisoformat(x['날짜']).date() <= end]
    except Exception:
        filtered = []

    total = sum(x['실현수익'] for x in filtered)
    color = '#ef4444' if total >= 0 else '#3b82f6'
    big_number("기간 실현 수익", f"{total:+.2f} USD", color)
    st.markdown("<br>", unsafe_allow_html=True)

    if not filtered:
        st.info(f"{start} ~ {end} 정산 내역 없음")
        return

    # 최신순 정렬
    df = pd.DataFrame(filtered)
    df = df.sort_values('날짜', ascending=False).reset_index(drop=True)
    st.dataframe(df, hide_index=True, width='stretch')

# ==========================================
# 탭 4: 슬롯 관리
# ==========================================
def tab_slot_management():
    # 새 슬롯 추가
    st.subheader("➕ 새 슬롯 추가")
    with st.form("new_slot", clear_on_submit=True):
        name = st.text_input("슬롯 이름", placeholder="SOXL 1회차")
        tk = st.text_input("종목 코드", "SOXL").upper()
        cc = st.columns(2)
        cap = cc[0].number_input("원금 ($)", min_value=100.0, value=5000.0, step=100.0)
        spl = cc[1].selectbox("분할 횟수", [20, 30, 40], index=2)
        tp = st.number_input("목표 수익률 (%)", min_value=1.0, value=20.0, step=1.0)
        if st.form_submit_button("✅ 슬롯 생성", width='stretch'):
            if name and tk:
                sid = create_slot(name, tk, cap, spl, tp)
                st.session_state.selected_slot_id = sid
                persist()
                st.success("슬롯 생성됨")
                st.rerun()
            else:
                st.error("이름과 종목 코드 입력 필요")

    # 현재 슬롯 관리
    cur = current_slot()
    if cur:
        st.divider()
        st.subheader("📌 현재 슬롯 설정")
        mode = cur.get('mode', 'normal')
        if mode == 'reverse':
            st.markdown(f"`{cur['name']}` ({cur['ticker']}) · 🔄 **리버스 모드**")
        else:
            st.markdown(f"`{cur['name']}` ({cur['ticker']}) · 일반 모드")

        # 모드 수동 전환 (리버스 → 일반)
        if mode == 'reverse':
            cls = fetch_close(cur['ticker'])
            exit_ok, threshold = check_reverse_exit(cur, cls)
            if exit_ok:
                st.success(f"✅ 종료조건 충족 (종가 ${cls:.2f} > 임계값 ${threshold:.2f})")
            elif threshold:
                cls_str = f"${cls:.2f}" if cls else "조회 실패"
                st.caption(f"종료조건 미충족 (종가 {cls_str} ≤ 임계값 ${threshold:.2f})")

            # 별지점 수동 입력
            st.markdown("**🌙 별지점 수동 조정**")
            auto_star = fetch_reverse_star(cur['ticker'])
            current_star = cur.get('manual_reverse_star') or auto_star or 0.0
            cc = st.columns([3, 1])
            new_star = cc[0].number_input(
                "별지점 ($)",
                min_value=0.0,
                value=float(current_star),
                step=0.01,
                format="%.2f",
                key=f"star_{cur['id']}",
                help=f"자동 계산값: ${auto_star:.2f}" if auto_star else "자동 계산 실패"
            )
            if cc[1].button("🔄 자동", width='stretch', key=f"star_auto_{cur['id']}"):
                cur['manual_reverse_star'] = None
                persist()
                st.rerun()
            if st.button("💾 별지점 저장", width='stretch', key=f"star_save_{cur['id']}"):
                cur['manual_reverse_star'] = float(new_star) if new_star > 0 else None
                persist()
                st.success("별지점 저장됨")
                st.rerun()

            if st.button("🔄 일반모드 복귀", width='stretch', key=f"normal_{cur['id']}"):
                cur['mode'] = 'normal'
                cur['reverse_entered_at'] = None
                cur['manual_reverse_star'] = None
                persist()
                st.success("일반모드로 복귀")
                st.rerun()
            st.markdown("---")

        with st.form(f"sl_{cur['id']}"):
            nn = st.text_input("이름", value=cur['name'])
            ntk = st.text_input("종목", value=cur['ticker']).upper()
            cc = st.columns(2)
            nc = cc[0].number_input("원금 ($)", min_value=100.0, value=float(cur['capital']), step=100.0)
            nspl = cc[1].selectbox("분할", [20, 30, 40],
                                 index=[20,30,40].index(cur['split']) if cur['split'] in [20,30,40] else 2)
            ntp = st.number_input("목표 수익률 (%)", min_value=1.0, value=float(cur['target_profit']), step=1.0)
            st.markdown("**🪜 추가 LOC 매수 지점 (급락 대비)**")
            cc2 = st.columns(2)
            n_extra = cc2[0].number_input("지점 개수", min_value=0, max_value=10,
                                           value=int(cur.get('extra_buy_levels', 3)), step=1)
            pct_extra = cc2[1].number_input("지점 간격 (%)", min_value=1.0, max_value=30.0,
                                             value=float(cur.get('extra_buy_pct_step', 5.0)), step=0.5)
            st.caption("별지점/평단 매수는 그대로 두고, 그 아래로 추가 LOC 매수 지점을 둬서 급락 시 1회매수액을 최대한 소진해요.")
            if st.form_submit_button("💾 설정 저장", width='stretch'):
                cur.update({'name': nn, 'ticker': ntk, 'capital': nc,
                            'split': nspl, 'target_profit': ntp,
                            'extra_buy_levels': int(n_extra),
                            'extra_buy_pct_step': float(pct_extra)})
                persist()
                st.success("저장됨")
                st.rerun()

        st.markdown("**🔧 T값 수동 조정**")
        nT = st.number_input("T 값", value=float(cur['T']), step=0.1, format="%.4f", key=f"tm_{cur['id']}")
        if st.button("T 적용", width='stretch', key=f"tmb_{cur['id']}"):
            cur['T'] = nT
            persist()
            st.success("T값 변경됨")
            st.rerun()

        st.divider()
        st.markdown("**⚠️ 위험 구역**")
        h, _, _, _, _ = calc_holdings(cur['transactions'], cur['capital'])
        cls = fetch_close(cur['ticker'])
        if h > 0 and cls:
            if st.button(f"🏁 전량매도 (${cls:.2f}) & 슬롯 종료", width='stretch', key=f"end_{cur['id']}"):
                slot_id_to_show = cur['id']
                cur['transactions'].append({
                    'date': str(date.today()), 'type': 'sell',
                    'price': float(cls), 'qty': int(h),
                    'mode': cur.get('mode', 'normal'),
                })
                cur['status'] = 'completed'
                cur['completed_at'] = str(date.today())
                st.session_state.selected_slot_id = None
                persist()
                st.session_state['pending_sell_dialog'] = {
                    'slot_id': slot_id_to_show,
                    'sell_type': 'full',
                }
                st.rerun()
        elif h == 0:
            if st.button("🏁 슬롯 종료 (보유 0)", width='stretch', key=f"end0_{cur['id']}"):
                cur['status'] = 'completed'
                cur['completed_at'] = str(date.today())
                st.session_state.selected_slot_id = None
                persist()
                st.rerun()

        with st.expander("🗑 슬롯 영구 삭제"):
            st.error("복구 불가능. 모든 거래 이력이 사라집니다.")
            if st.button("⚠️ 삭제 확정", width='stretch', key=f"del_{cur['id']}"):
                del st.session_state.slots[cur['id']]
                st.session_state.selected_slot_id = None
                persist()
                st.rerun()

    # 완료된 슬롯
    comp = completed_slots()
    if comp:
        st.divider()
        st.subheader(f"🏁 완료된 슬롯 ({len(comp)}개)")
        for s in comp:
            rz = calc_realized(s['transactions'])
            _, _, _, used_c, _ = calc_holdings(s['transactions'], s['capital'])
            pct = (rz / used_c * 100) if used_c else 0
            with st.expander(f"{s['name']} ({s['ticker']}) · ${rz:.2f} ({pct:+.2f}%)"):
                st.caption(f"기간: {s['created_at']} → {s.get('completed_at', '-')}")
                cc = st.columns(2)
                if cc[0].button("🔄 복구 (진행 전환)", key=f"rc_{s['id']}", width='stretch'):
                    s['status'] = 'active'
                    s['completed_at'] = None
                    persist()
                    st.rerun()
                if cc[1].button("🗑 삭제", key=f"dc_{s['id']}", width='stretch'):
                    del st.session_state.slots[s['id']]
                    persist()
                    st.rerun()

    # 백업/복원
    st.divider()
    st.subheader("💾 백업 / 복원")
    if st.session_state.slots:
        export = {'slots': st.session_state.slots}
        st.download_button(
            "백업 파일 다운로드",
            json.dumps(export, ensure_ascii=False, indent=2),
            file_name=f"mumae_{user}_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            mime="application/json",
            width='stretch'
        )
    up = st.file_uploader("백업 파일로 복원", type='json')
    if up:
        try:
            d = json.load(up)
            st.session_state.slots = d['slots']
            persist()
            st.success("복원 완료")
            st.rerun()
        except Exception as e:
            st.error(f"오류: {e}")

# ==========================================
# 라우팅
# ==========================================
render_sidebar()
cur = render_header()

if not active_slots():
    # 슬롯 없을 때: 슬롯 관리 탭만 표시
    st.divider()
    tab_slot_management()
else:
    # 슬롯 있을 때: 4개 탭
    t1, t2, t3, t4 = st.tabs(["📊 포트폴리오", "📝 매매이력", "💰 정산이력", "⚙️ 슬롯 관리"])
    with t1:
        tab_portfolio(cur)
    with t2:
        tab_history(cur)
    with t3:
        tab_settlement(cur)
    with t4:
        tab_slot_management()

# 매도 결과 모달 (페이지 최상위에서 호출)
if st.session_state.get('pending_sell_dialog'):
    sell_result_dialog()
