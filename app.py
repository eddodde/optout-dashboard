import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit.components.v1 as components
import pathlib

# ── 기본 설정 ──────────────────────────────────────────────
st.set_page_config(
    page_title="VIP 도달·이탈 진단 대시보드",
    page_icon="📉",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 등급 순서(최상위 → 하위) & 그룹 정의
GRADE_ORDER = ["SP", "PT", "GD", "SV", "BK", "PP", "RD"]
GROUPS = {"VIP": ["SP", "PT", "GD", "SV", "BK"], "일반": ["PP", "RD"]}
GRADE2GROUP = {g: grp for grp, gs in GROUPS.items() for g in gs}
GROUP_COLOR = {"VIP": "#4C72B0", "일반": "#B0B0B0"}
CHANNELS = ["PUSH", "SMS", "EMAIL"]
CH_COLOR = {"PUSH": "#4C72B0", "SMS": "#DD8452", "EMAIL": "#55A868"}

# Plotly 한글 폰트 — Noto Sans KR 웹폰트를 강제 로드해 SVG 텍스트에 적용
KFONT = "'Noto Sans KR','Malgun Gothic','Apple SD Gothic Neo',sans-serif"

# ── 커스텀 CSS ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap');
    .metric-card {
        background: #f8f9fa; border-radius: 10px; padding: 14px 18px;
        border-left: 4px solid #4C72B0; margin-bottom: 10px;
    }
    .metric-label { font-size: 13px; color: #666; margin-bottom: 4px; }
    .metric-value { font-size: 24px; font-weight: 700; color: #1a1a2e; }
    .metric-sub  { font-size: 12px; color: #888; margin-top: 2px; }
    .section-title {
        font-size: 18px; font-weight: 700; color: #1a1a2e;
        margin: 26px 0 12px 0; padding-bottom: 6px;
        border-bottom: 2px solid #e9ecef;
        scroll-margin-top: 70px;
    }
    .hint { font-size: 12px; color: #999; margin: -4px 0 10px 0; }
    /* 사이드바 분석 메뉴 */
    a.navlink {
        display: block; padding: 8px 12px; margin: 4px 0; border-radius: 8px;
        background: #f2f5fa; color: #2E68B0; text-decoration: none;
        font-size: 14px; font-weight: 600; border: 1px solid #e3e9f2;
    }
    a.navlink:hover { background: #e3ecf8; color: #163E78; }
    details.navacc { margin: 4px 0; }
    summary.navgroup { cursor: pointer; font-size: 12px; font-weight: 700; color: #55606f;
        letter-spacing: .02em; padding: 7px 8px; border-radius: 6px; list-style: none;
        user-select: none; background: #eef1f6; }
    summary.navgroup::-webkit-details-marker { display: none; }
    summary.navgroup:hover { background: #e3e9f2; color: #1a1a2e; }
    summary.navgroup::before { content: "▸ "; color: #99a3b3; }
    details[open] summary.navgroup::before { content: "▾ "; }
    a.navlink.navsub { margin: 3px 0 3px 10px; font-size: 13px; padding: 6px 10px; }
    .insight {
        background: #eef4fb; border-left: 4px solid #4C72B0; border-radius: 8px;
        padding: 12px 16px; margin: 6px 0 14px 0; font-size: 14px; line-height: 1.6;
    }
    .insight.warn { background: #fdeeee; border-left-color: #C44E52; }
    .insight.ok   { background: #eef7f0; border-left-color: #55A868; }
    .insight b { color: #1a1a2e; }
    .insight ul { margin: 0; padding-left: 20px; }
    .insight li { margin: 5px 0; }
    .insight .cap { font-size: 12px; font-weight: 700; color: #666; display: block; margin-bottom: 6px; }
    /* 진단 논리(소거법) 사다리 */
    .logic-step { display: flex; gap: 12px; margin: 0 0 8px; }
    .logic-num { flex: 0 0 32px; height: 32px; border-radius: 50%; background: #4C72B0;
        color: #fff; font-weight: 700; display: flex; align-items: center; justify-content: center; }
    .logic-card { flex: 1; background: #f8f9fa; border-radius: 8px; padding: 10px 14px;
        font-size: 14px; line-height: 1.65; border-left: 3px solid #4C72B0; }
    .logic-card.ok { background: #eef7f0; border-left-color: #55A868; }
    .logic-step.final .logic-num { background: #55A868; }
    .logic .tag { background: #e3e9f2; color: #55606f; font-size: 11px; padding: 1px 6px;
        border-radius: 4px; margin-left: 4px; }
    .logic .verdict { margin-top: 6px; font-size: 12.5px; }
    .logic .verdict.warn { color: #C44E52; font-weight: 600; }
    /* 독립 검토 의견(외부 관점) */
    .opinion { background: #fffaf0; border: 1px solid #f0d9a8; border-left: 4px solid #E8A33D;
        border-radius: 8px; padding: 14px 18px; margin: 8px 0 14px; font-size: 14px; line-height: 1.7; }
    .opinion .ohead { font-weight: 700; color: #8a5a00; margin-bottom: 8px; }
    .opinion ol { margin: 0; padding-left: 20px; }
    .opinion li { margin: 8px 0; }
    .opinion .chk { color: #8a6d3b; font-size: 12.5px; }
    .opinion b { color: #1a1a2e; }
</style>
""", unsafe_allow_html=True)


# ── 데이터 로드 / 정규화 ────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(file):
    """업로드한 엑셀(.xls/.xlsx)을 (wide: 날짜×등급), (long: 날짜×등급×채널)으로 정규화."""
    raw = pd.read_excel(file, sheet_name=0, header=0)
    raw.columns = [str(c).strip() for c in raw.columns]
    raw = raw.loc[:, [c for c in raw.columns if c and not c.startswith("Unnamed")]]

    needed = {"STD_DD", "GRADE_CD"}
    if not needed.issubset(raw.columns):
        raise ValueError(f"필수 컬럼 누락: {needed - set(raw.columns)}")

    raw = raw.dropna(subset=["STD_DD", "GRADE_CD"])
    raw["STD_DD"] = pd.to_datetime(raw["STD_DD"].astype(str).str.strip(),
                                   format="%Y%m%d", errors="coerce")
    raw = raw.dropna(subset=["STD_DD"])
    raw["GRADE_CD"] = raw["GRADE_CD"].astype(str).str.strip()
    raw = raw[raw["GRADE_CD"] != "TOTAL"]  # 합계행 제외

    for c in raw.columns:
        if c not in ("STD_DD", "NEW_GBN", "GRADE_CD"):
            raw[c] = pd.to_numeric(raw[c], errors="coerce").fillna(0)

    # wide: 날짜 × 등급 (도달률·앱 미보유/삭제 중심)
    w = pd.DataFrame({
        "date": raw["STD_DD"],
        "grade": raw["GRADE_CD"],
        "act": raw.get("ACT_MEM", 0),
        "act_push": raw.get("ACT_PUSH_MEM", 0),
        "tot_push": raw.get("TOT_PUSH_MEM", 0),
        "tot_sms": raw.get("TOT_SMS_MEM", 0),
        "tot_email": raw.get("TOT_EMAIL_MEM", 0),
    })
    for ch in CHANNELS:
        w[f"new_{ch.lower()}"] = raw.get(f"NEW_{ch}_MEM", 0)
        w[f"out_{ch.lower()}"] = raw.get(f"OUT_{ch}_MEM", 0)
    w["group"] = w["grade"].map(GRADE2GROUP).fillna("일반")
    # 앱 미보유/삭제 = 앱푸시 수신자(ACT_PUSH) − 실제 발송가능(TOT_PUSH)
    w["unreach_push"] = (w["act_push"] - w["tot_push"]).clip(lower=0)
    w["reach_push"] = np.where(w["act_push"] > 0, w["tot_push"] / w["act_push"] * 100, 0)
    w["out_all"] = w[[f"out_{c.lower()}" for c in CHANNELS]].sum(axis=1)
    w["new_all"] = w[[f"new_{c.lower()}" for c in CHANNELS]].sum(axis=1)
    w["net_all"] = w["new_all"] - w["out_all"]

    # long: 날짜 × 등급 × 채널 (수신거부 분석)
    recs = []
    for _, r in w.iterrows():
        for ch in CHANNELS:
            cl = ch.lower()
            recs.append({
                "date": r["date"], "grade": r["grade"], "group": r["group"], "channel": ch,
                "tot": r[f"tot_{cl}"], "new": r[f"new_{cl}"], "out": r[f"out_{cl}"],
            })
    L = pd.DataFrame(recs)
    L["net"] = L["new"] - L["out"]
    L["out_rate"] = np.where(L["tot"] > 0, L["out"] / L["tot"] * 100, 0)
    return w, L


@st.cache_data(show_spinner=False)
def load_longterm():
    """장기 추세(전체·등급무관) CSV. 없으면 None."""
    p = pathlib.Path(__file__).parent / "data" / "longterm.csv"
    if not p.exists():
        return None
    lt = pd.read_csv(p)
    lt["date"] = pd.to_datetime(lt["date"])
    lt["metric"] = lt["metric"].astype(str).str.replace(" ", "", regex=False)
    lt["value"] = pd.to_numeric(lt["value"], errors="coerce")
    return lt


def lt_series(lt, source, metric, segment="Total"):
    d = lt[(lt["source"] == source) & (lt["segment"] == segment) & (lt["metric"] == metric)]
    return d.set_index("date")["value"].sort_index()


# 등급명 매핑 (DAU 파일 → 코드)
GRADE_NAME_MAP = {"S-Platinum": "SP", "Platinum": "PT", "Gold": "GD", "Silver": "SV",
                  "Black": "BK", "Purple": "PP", "Red": "RD", "Total": "TOTAL", "*TOTAL": "TOTAL"}


@st.cache_data(show_spinner=False)
def load_dau_monthly(file):
    """등급별 월 파일 → long [date(월), metric(MAU/DAU), grade, value].
    두 포맷 자동 인식: ① 구형(MAU/DAU 블록, LFMS 포함) ② 신형 UV(B2B 제외, 2020~, UV→MAU)."""
    raw = pd.read_excel(file, sheet_name=0, header=None)

    # ③ 간이 포맷: A열 '구분' + DAU/MAU 행, 월 컬럼 (VIP 합계, B2B 제외)
    if (raw.iloc[:2, 0].astype(str).str.strip() == "구분").any():
        years = raw.iloc[0].ffill()
        months = raw.iloc[1]
        col_dates = {}
        for c in range(1, raw.shape[1]):
            s = str(months.iloc[c]).replace("월", "").strip()
            try:
                yi = int(float(str(years.iloc[c]).replace("년", "").strip()))
            except (ValueError, TypeError):
                continue
            if s.isdigit():
                col_dates[c] = pd.Timestamp(yi, int(s), 1)
        recs = []
        for r in range(2, raw.shape[0]):
            met = str(raw.iat[r, 0]).strip().upper()
            if met not in ("DAU", "MAU"):
                continue
            for c, dt in col_dates.items():
                v = pd.to_numeric(raw.iat[r, c], errors="coerce")
                if pd.notna(v):
                    recs.append({"date": dt, "metric": met, "grade": "VIP", "value": v})
        return pd.DataFrame(recs)

    is_uv = raw.iloc[:3].astype(str).apply(lambda s: s.str.contains("회원구분상세")).any().any()

    if is_uv:  # 신형: 1행=연도('2020년'), 2행=월 라벨, 등급=4열째, 지표=UV(→MAU)
        years = raw.iloc[0].astype(str).str.replace("년", "", regex=False).ffill()
        months = raw.iloc[1]
        col_dates = {}
        for c in range(4, raw.shape[1]):
            s = str(months.iloc[c]).replace("월", "").strip()
            ys = str(years.iloc[c]).replace(".0", "").strip()
            if s.isdigit() and ys.isdigit():
                col_dates[c] = pd.Timestamp(int(ys), int(s), 1)
        recs = []
        for r in range(2, raw.shape[0]):
            grade = GRADE_NAME_MAP.get(str(raw.iat[r, 3]).strip(), str(raw.iat[r, 3]).strip())
            for c, dt in col_dates.items():
                v = pd.to_numeric(raw.iat[r, c], errors="coerce")
                if pd.notna(v):
                    recs.append({"date": dt, "metric": "MAU", "grade": grade, "value": v})
        return pd.DataFrame(recs)

    # 구형
    years = raw.iloc[1].ffill()
    months = raw.iloc[2]
    col_dates = {}
    for c in range(2, raw.shape[1]):
        y, ml = years.iloc[c], months.iloc[c]
        s = str(ml).replace("월", "").strip()
        if pd.isna(y) or not s.isdigit():
            continue
        try:
            col_dates[c] = pd.Timestamp(int(float(y)), int(s), 1)
        except Exception:
            continue
    recs, cur = [], None
    for r in range(3, raw.shape[0]):
        g0 = raw.iat[r, 0]
        if pd.notna(g0) and str(g0).strip():
            cur = str(g0).strip()          # MAU / DAU
        grade = GRADE_NAME_MAP.get(str(raw.iat[r, 1]).strip(), str(raw.iat[r, 1]).strip())
        for c, dt in col_dates.items():
            v = pd.to_numeric(raw.iat[r, c], errors="coerce")
            if pd.notna(v):
                recs.append({"date": dt, "metric": cur, "grade": grade, "value": v})
    return pd.DataFrame(recs)


@st.cache_data(show_spinner=False)
def load_dau_channel(file):
    """채널별 일 DAU 파일 → long [date, channel, value]. (col3=2025-01-01 연속일)"""
    raw = pd.read_excel(file, sheet_name=0, header=None)
    base = pd.Timestamp(2025, 1, 1)
    recs = []
    for r in range(2, raw.shape[0]):
        ch = str(raw.iat[r, 1]).strip().replace("*", "")
        if not ch or ch.lower() == "nan":
            continue
        for c in range(2, raw.shape[1]):
            v = pd.to_numeric(raw.iat[r, c], errors="coerce")
            if pd.notna(v):
                recs.append({"date": base + pd.Timedelta(days=c - 2), "channel": ch, "value": v})
    return pd.DataFrame(recs)


def fnum(x):
    return f"{int(round(x)):,}"


def fsigned(x):
    """내부 표기 규칙: 음수 = 빨강 △숫자, 양수 = 일반 숫자."""
    v = int(round(x))
    if v < 0:
        return f'<span style="color:#C44E52">△{abs(v):,}</span>'
    return f"{v:,}"


def metric_card(label, value, sub="", color="#4C72B0"):
    st.markdown(
        f'<div class="metric-card" style="border-left-color:{color}">'
        f'<div class="metric-label">{label}</div>'
        f'<div class="metric-value">{value}</div>'
        f'<div class="metric-sub">{sub}</div></div>',
        unsafe_allow_html=True)


def section(title, hint="", anchor=None):
    aid = f' id="{anchor}"' if anchor else ""
    st.markdown(f'<div class="section-title"{aid}>{title}</div>', unsafe_allow_html=True)
    if hint:
        st.markdown(f'<div class="hint">{hint}</div>', unsafe_allow_html=True)


# 사이드바 분석 메뉴 — 논리 흐름별 그룹핑 (그룹라벨, [(anchor, 라벨), ...])
MENU = [
    ("① 문제 정의", [
        ("sec-dau", "DAU 문제 진단 (업로드 시)"),
        ("sec-core", "VIP 도달 스냅샷"),
    ]),
    ("② 현황 진단 — 도달·이탈(도달력 축)", [
        ("sec-reach", "앱푸시 도달 진단"),
        ("sec-optout", "수신거부 분석"),
        ("sec-within", "그룹 내 등급별"),
        ("sec-group", "그룹 비교 (일반 참고)"),
        ("sec-trend", "장기 추세 (전체)"),
    ]),
    ("③ 결론 — 논리 & 방향", [
        ("sec-logic", "진단 논리 (소거→레버)"),
        ("sec-action", "인사이트 · 시사점 · 액션"),
    ]),
    ("④ 부록", [
        ("sec-table", "상세 데이터"),
    ]),
]


def insight(bullets, kind="", cap="💡 시사점"):
    """인사이트를 불릿 목록으로 렌더(시사점 위주). bullets=문자열 또는 리스트."""
    if isinstance(bullets, str):
        bullets = [bullets]
    bullets = [b for b in bullets if b]
    if not bullets:
        return
    items = "".join(f"<li>{b}</li>" for b in bullets)
    head = f'<span class="cap">{cap}</span>' if cap else ""
    st.markdown(f'<div class="insight {kind}">{head}<ul>{items}</ul></div>', unsafe_allow_html=True)


def plot(fig, title=None):
    """차트 제목은 Streamlit 텍스트로 렌더(폰트 보장) + 차트엔 한글 폰트 직접 적용."""
    if title:
        st.markdown(f'<div style="font-weight:700;font-size:15px;margin:10px 0 -6px">{title}</div>',
                    unsafe_allow_html=True)
    fig.update_layout(font=dict(family=KFONT))
    # 축 숫자 k/M 금지 → #,##0(천단위 콤마). tickformat이 지정된 축은 그대로 존중.
    fig.update_yaxes(exponentformat="none", separatethousands=True)
    st.plotly_chart(fig, use_container_width=True)


def drop_outliers(s, window=21, k=4.0):
    """롤링 중앙값 대비 |편차| > k*MAD 인 이상치 날짜를 NaN 처리(차트에서 제외)."""
    s = s.sort_index()
    med = s.rolling(window, center=True, min_periods=5).median()
    mad = (s - med).abs().rolling(window, center=True, min_periods=5).median()
    thr = (k * mad).replace(0, np.nan)
    mask = (s - med).abs() > thr
    out = s.copy()
    out[mask.fillna(False)] = np.nan
    return out


def drop_incomplete_months(s, cutoff):
    """월초(MS) 인덱스의 시계열/프레임에서 '월말 > cutoff'인 집계 중 부분월을 제거.
    (예: 7/1~7/8까지만 있는 7월은 MAU가 덜 쌓여 DAU/MAU 비율이 왜곡되므로 헤드라인·비율에서 제외)"""
    if s is None or len(s) == 0 or cutoff is None:
        return s
    ends = s.index + pd.offsets.MonthEnd(0)
    return s[ends <= pd.Timestamp(cutoff)]


def trend_word(daily_series):
    """일별 시계열의 추세를 첫/끝 3일 평균 비교로 판정."""
    s = daily_series.dropna()
    if len(s) < 2:
        return "—", 0.0
    k = max(1, min(3, len(s) // 2))
    first, last = s.iloc[:k].mean(), s.iloc[-k:].mean()
    if first == 0:
        return ("증가" if last > 0 else "유지"), 0.0
    chg = (last - first) / first * 100
    word = "증가" if chg > 5 else ("감소" if chg < -5 else "유지")
    return word, chg


# ── 헤더 ───────────────────────────────────────────────────
st.title("📉 VIP 도달·이탈 진단 대시보드")
st.caption("VIP DAU 역신장 가설 추적 — 채널 수신거부(이탈) + 앱 미보유/삭제로 도달 가능 모수가 "
           "얼마나 줄어드는지 등급·그룹별로 진단합니다.")

# ── 사이드바: 업로드 ───────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 설정")
    up = st.file_uploader("도달 데이터 (.xls / .xlsx)", type=["xls", "xlsx"])
    st.caption("STD_DD, GRADE_CD, ACT/TOT/NEW/OUT_채널_MEM 컬럼 포함 export 파일")
    with st.expander("📉 DAU 데이터 (선택)"):
        up_dau = st.file_uploader("등급별 월 MAU(UV) 또는 DAU/MAU", type=["xls", "xlsx"], key="dau_m")
        up_chdau = st.file_uploader("채널별 일 DAU", type=["xls", "xlsx"], key="dau_ch")
        st.caption("업로드형(레포 미저장). 올리면 'DAU 문제 진단' 섹션이 생성됩니다.")

if up is None:
    st.info("👈 사이드바에서 데이터 파일을 업로드하면 진단이 시작됩니다.")
    with st.expander("📋 지표 정의"):
        st.markdown("""
- **퍼널**: 전체유효회원 → **수신동의율**(`ACT_PUSH/ACT_MEM`) → **타겟팅가능율=도달률**(`TOT_PUSH/ACT_PUSH`)
- **앱 미보유/삭제** = `ACT_PUSH_MEM − TOT_PUSH_MEM` (수신동의했으나 앱 미보유/삭제로 발송 불가)
- **증감** = 신규추가(`NEW`) − 기존이탈(`OUT`)  ·  **수신거부율** = `OUT / TOT`
- **그룹** — VIP: SP·PT·GD·SV·BK / 일반: PP·RD
""")
    st.stop()

try:
    W, L = load_data(up)
except Exception as e:
    st.error(f"데이터를 읽는 중 오류가 발생했습니다: {e}")
    st.stop()

# ── 사이드바 필터 ───────────────────────────────────────────
with st.sidebar:
    st.divider()
    st.markdown("**📂 분석 메뉴**")
    _nav = ""
    for _glabel, _items in MENU:
        _open = " open" if _glabel.startswith("①") else ""
        _links = "".join(f'<a href="#{a}" class="navlink navsub">{lbl}</a>' for a, lbl in _items)
        _nav += (f'<details class="navacc" name="navmenu"{_open}>'
                 f'<summary class="navgroup">{_glabel}</summary>{_links}</details>')
    st.markdown(_nav, unsafe_allow_html=True)
    st.divider()
    st.caption("🔎 필터")
    dmin, dmax = W["date"].min().date(), W["date"].max().date()
    dr = st.date_input("기간", value=(dmin, dmax), min_value=dmin, max_value=dmax)
    d0, d1 = dr if isinstance(dr, tuple) and len(dr) == 2 else (dmin, dmax)
    st.caption("상세 섹션 기본값은 VIP입니다. 일반/전체 비교는 '그룹 비교' 섹션에서 항상 함께 보입니다.")
    sel_groups = st.multiselect("그룹", list(GROUPS.keys()), default=["VIP"])
    grade_pool = [g for g in GRADE_ORDER if GRADE2GROUP.get(g) in sel_groups]
    sel_grades = st.multiselect("등급", grade_pool, default=grade_pool)
    sel_channels = st.multiselect("채널 (수신거부 분석)", CHANNELS, default=CHANNELS)

wmask = (W["date"].dt.date >= d0) & (W["date"].dt.date <= d1) & (W["grade"].isin(sel_grades))
fw = W[wmask].copy()
lmask = ((L["date"].dt.date >= d0) & (L["date"].dt.date <= d1) &
         (L["grade"].isin(sel_grades)) & (L["channel"].isin(sel_channels)))
fl = L[lmask].copy()

if fw.empty:
    st.warning("선택한 조건에 데이터가 없습니다. 필터를 조정해 주세요.")
    st.stop()

n_days = fw["date"].nunique()
last_day = fw["date"].max()
fw_last = fw[fw["date"] == last_day]          # 최근일 스냅샷(도달률·모수)
grade_order_sel = [g for g in GRADE_ORDER if g in fw["grade"].unique()]

# 핵심진단(VIP)·그룹비교용: 등급 필터 무관, 기간/채널만 적용 (전체 등급 유지)
fw_d = W[(W["date"].dt.date >= d0) & (W["date"].dt.date <= d1)].copy()
fl_d = L[(L["date"].dt.date >= d0) & (L["date"].dt.date <= d1) & (L["channel"].isin(sel_channels))].copy()
fw_d_last = fw_d[fw_d["date"] == fw_d["date"].max()] if not fw_d.empty else fw_d


def group_snapshot(grp, snap, period_long, period_wide):
    """그룹 단위 집계 (스냅샷=도달률, 기간합=거부/순증감)."""
    s = snap[snap["group"] == grp]
    pl = period_long[period_long["group"] == grp]
    pw = period_wide[period_wide["group"] == grp]
    act = s["act"].sum()
    act_push = s["act_push"].sum()
    tot_push = s["tot_push"].sum()
    # 채널별 증감(=신규추가−기존이탈, 기간 누적) / 채널별 기존이탈
    chnet = {ch: pl[pl["channel"] == ch]["new"].sum() - pl[pl["channel"] == ch]["out"].sum()
             for ch in CHANNELS}
    chout = {ch: pl[pl["channel"] == ch]["out"].sum() for ch in CHANNELS}
    return {
        "act": act,
        "act_push": act_push,
        "tot_push": tot_push,
        "unreach": s["unreach_push"].sum(),
        "consent": (act_push / act * 100) if act else 0,   # 수신동의율
        "reach": (tot_push / act_push * 100) if act_push else 0,   # 타겟팅가능율(도달률)
        "out": pl["out"].sum(),
        "new": pl["new"].sum(),
        "net": pl["new"].sum() - pl["out"].sum(),
        "chnet": chnet,
        "chout": chout,
        "out_trend": trend_word(pw.groupby("date")["out_all"].sum())[0],
    }


# ════════════════════════════════════════════════════════════
# 0.7 DAU 문제 진단 (업로드 시) — 왜 '도달'이 레버인가 (소거법)
# ════════════════════════════════════════════════════════════
dau_sum = {}   # 결론 섹션에서 재사용할 DAU 요약(있으면 채워짐)
if up_dau is not None or up_chdau is not None:
    section("DAU 문제 진단 — 왜 '도달'이 레버인가",
            "DAU 하락(빈도 문제) → 자발 회복·콘텐츠 개선 제약 → ∴ 통제 레버 = 발송의 반응 유도력 제고(행동 트리거 D-1→실시간 × 도달력)", anchor="sec-dau")
    VIPG = GROUPS["VIP"]
    try:
        # ── (0) 채널별 일 DAU 먼저 파싱 (B2B 제외 = 진성 VIP, 내부 관리지표와 동일 기준)
        mon = None
        anom_txt = ""
        cutoff = pd.Timestamp.today().normalize()   # 부분월 판정 기준(데이터 최신일)
        if up_chdau is not None:
            dc = load_dau_channel(up_chdau)
            cutoff = dc["date"].max()   # 일 데이터가 있으면 실제 최신일을 기준으로
            # 단기 급증 구간(대형행사 등) 제외: TOTAL이 91일 롤링 중앙값의 1.6배 초과인 날
            # (예: 25.9.23~10.30 차수제 대형행사 — 평상시 추세·YoY 비교를 위해 제외, 양 연도 동일 규칙 적용)
            tot_d = dc[dc["channel"] == "TOTAL"].set_index("date")["value"].sort_index()
            med_d = tot_d.rolling(91, center=True, min_periods=30).median()
            bad_days = set(tot_d.index[(tot_d / med_d > 1.6).fillna(False)])
            if bad_days:
                dc = dc[~dc["date"].isin(bad_days)]
                anom_txt = (f"※ 대형행사 등 급증일 {len(bad_days)}일 제외 "
                            f"({min(bad_days).date()} ~ {max(bad_days).date()}) — 평상시 추세 비교 기준")
            mon = (dc.groupby([pd.Grouper(key="date", freq="MS"), "channel"])["value"].mean()
                   .unstack().sort_index())
            mon = drop_incomplete_months(mon, cutoff)   # 집계 중 부분월(예: 7/1~8) 제거
            if "TOTAL" in mon.columns and len(mon) >= 13 and mon["TOTAL"].iloc[-13]:
                dau_sum["dau_yoy"] = (mon["TOTAL"].iloc[-1] / mon["TOTAL"].iloc[-13] - 1) * 100
                dau_sum["dau_now"] = mon["TOTAL"].iloc[-1]
                dau_sum["yoy_basis"] = "B2B 제외"
                dau_sum["yoy_total_now"] = mon["TOTAL"].iloc[-1]
                dau_sum["yoy_total_prev"] = mon["TOTAL"].iloc[-13]
                if "PUSH" in mon.columns:
                    dau_sum["yoy_push_now"] = mon["PUSH"].iloc[-1]
                    dau_sum["yoy_push_prev"] = mon["PUSH"].iloc[-13]

        # ── (1) 등급별 월 MAU(±DAU): VIP 추세 + 빈도(DAU/MAU)
        if up_dau is not None:
            dm = load_dau_monthly(up_dau)   # 간이(VIP합계)/UV(등급별 MAU)/구형(MAU+DAU) 자동 인식

            def vip_series(metric):
                d = dm[dm["metric"] == metric]
                g = d[d["grade"].isin(VIPG)]         # 등급별 파일이면 SP~BK 합산
                if g.empty:
                    g = d[d["grade"] == "VIP"]        # VIP 합계 행만 있는 간이 포맷
                return g.groupby("date")["value"].sum().sort_index()

            mau_s = vip_series("MAU")
            dau_in_file = vip_series("DAU")
            # 집계 중인 말단 부분월 제거 — 달력 기준(월말 > 데이터 최신일이면 제외).
            # MAU는 순방문자가 월초에 몰려 빨리 포화(sublinear)하므로 '하락 비율' 감지는 못 잡음 → 반드시 달력으로 판정.
            mau_s = drop_incomplete_months(mau_s, cutoff)
            dau_in_file = drop_incomplete_months(dau_in_file, cutoff)
            # 채널 파일(이상일 제외 후)이 있으면 우선 — 월별 파일 DAU는 25.10 중복집계 오염 포함
            if mon is not None and "TOTAL" in mon.columns:
                dau_s, dau_basis = mon["TOTAL"], "채널 파일 DAU(B2B·이상일 제외)"
            elif len(dau_in_file):
                dau_s, dau_basis = dau_in_file, "월별 파일 DAU(※25.9~10 중복집계 오염 가능)"
                if "dau_yoy" not in dau_sum and len(dau_in_file) >= 13 and dau_in_file.iloc[-13]:
                    dau_sum["dau_yoy"] = (dau_in_file.iloc[-1] / dau_in_file.iloc[-13] - 1) * 100
                    dau_sum["yoy_basis"] = "월별 파일"
            else:
                dau_s, dau_basis = None, ""

            if len(mau_s) > 1:
                both = pd.concat([mau_s.rename("MAU")] +
                                 ([dau_s.rename("DAU")] if dau_s is not None else []), axis=1)
                ov = both.dropna()
                if dau_s is not None and len(ov) > 1:
                    ov = ov.assign(ratio=np.where(ov["MAU"] > 0, ov["DAU"] / ov["MAU"] * 100, np.nan))
                    d0m, d1m = ov.iloc[0], ov.iloc[-1]
                    dau_sum.update(stick0=d0m["ratio"], stick1=d1m["ratio"])
                    if "dau_now" not in dau_sum:
                        dau_sum["dau_now"] = d1m["DAU"]
                    mau_yoy = ((mau_s.iloc[-1] / mau_s.iloc[-13] - 1) * 100
                               if len(mau_s) >= 13 and mau_s.iloc[-13] else None)
                    dau_sum["mau_growth"] = mau_yoy if mau_yoy is not None else (mau_s.iloc[-1] / mau_s.iloc[0] - 1) * 100
                    k1, k2, k3 = st.columns(3)
                    with k1: metric_card("VIP DAU (최근월)", fnum(d1m["DAU"]),
                                         (f"전년비 {dau_sum['dau_yoy']:+.1f}% ({dau_sum.get('yoy_basis','')})"
                                          if "dau_yoy" in dau_sum else dau_basis))
                    with k2: metric_card("VIP MAU (최근월)", fnum(mau_s.iloc[-1]),
                                         (f"전년비 {mau_yoy:+.1f}%" if mau_yoy is not None else "B2B 제외"))
                    with k3: metric_card("DAU/MAU (스티키니스)", f"{d1m['ratio']:.1f}%",
                                         f"{d0m['ratio']:.1f}% → {d1m['ratio']:.1f}% "
                                         f"(월평균 방문일수 ≈ {d0m['ratio']*30.4/100:.1f}일 → {d1m['ratio']*30.4/100:.1f}일)")
                    _basis_m = pd.Timestamp(d1m.name).strftime("%Y-%m")
                    cL, cR = st.columns(2)
                    with cL:
                        figd = go.Figure()
                        figd.add_bar(x=mau_s.index, y=mau_s.values, name="VIP MAU", marker_color="#c9d6e8", opacity=0.7)
                        figd.add_scatter(x=ov.index, y=ov["ratio"], name="DAU/MAU(%)", yaxis="y2",
                                         mode="lines+markers", line=dict(color="#C44E52", width=3))
                        figd.update_layout(height=320, margin=dict(t=10, b=10), hovermode="x unified",
                                           legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                                           legend_title_text="", yaxis=dict(title="VIP MAU(명)"),
                                           yaxis2=dict(title="DAU/MAU(%)", overlaying="y", side="right",
                                                       showgrid=False, tickformat=".1f"))
                        plot(figd, "VIP MAU는 성장하는데 스티키니스(DAU/MAU)는 하락")
                        st.caption(f"MAU=월별 파일(B2B 제외) · DAU={dau_basis} · 스티키니스는 두 시계열이 겹치는 구간만 표시 · "
                                   f"헤드라인·전년비·스티키니스는 마지막 완료월({_basis_m}) 기준(집계 중 부분월 자동 제외)")
                    with cR:
                        # 전년 동월 오버레이: 같은 달끼리 포개서 연도 간 갭 확인
                        # (라디오는 차트 아래 배치 — 좌우 차트 시작 높이를 맞추기 위함)
                        ymet = st.session_state.get("yoy_overlay_metric", "스티키니스(DAU/MAU)")
                        if ymet.startswith("스티"):
                            yser, ytitle, yfmt = ov["ratio"], "DAU/MAU(%)", ".1f"
                        elif ymet == "DAU":
                            yser, ytitle, yfmt = ov["DAU"], "VIP DAU(명)", ",.0f"
                        else:
                            yser, ytitle, yfmt = mau_s, "VIP MAU(명)", ",.0f"
                        yov = yser.dropna().to_frame("v")
                        yov["year"], yov["m"] = yov.index.year, yov.index.month
                        yrs = sorted(yov["year"].unique())
                        grays = ["#b9c6d8", "#8aa2c0", "#6b87ab"]   # 과거 연도(옅은→진한 회청)
                        figy = go.Figure()
                        for i, yr in enumerate(yrs):
                            d = yov[yov["year"] == yr].sort_values("m")
                            last = (yr == yrs[-1])
                            figy.add_scatter(x=d["m"], y=d["v"], name=str(yr), mode="lines+markers",
                                             line=dict(color="#C44E52" if last else grays[i % len(grays)],
                                                       width=3 if last else 2))
                        figy.update_layout(height=320, margin=dict(t=10, b=10), hovermode="x unified",
                                           legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
                                           legend_title_text="", yaxis=dict(title=ytitle, tickformat=yfmt))
                        figy.update_xaxes(tickmode="array", tickvals=list(range(1, 13)),
                                          ticktext=[f"{m}월" for m in range(1, 13)], title=None)
                        plot(figy, f"전년 동월 비교 — {ymet}")
                        st.radio("비교 지표", ["스티키니스(DAU/MAU)", "DAU", "MAU"],
                                 horizontal=True, key="yoy_overlay_metric")
                        st.caption("빨강=최근 연도 · 회청=과거 연도 · 같은 달끼리 세로로 비교")

                    insight([
                        f"VIP <b>MAU는 유지·증가</b>인데 DAU가 빠짐 → <b>스티키니스(DAU/MAU)가 {d0m['ratio']:.0f}%→{d1m['ratio']:.0f}%</b>로 하락"
                        f"(월평균 방문일수 ≈ {d0m['ratio']*30.4/100:.1f}일 → {d1m['ratio']*30.4/100:.1f}일). "
                        "'앱 쓰는 사람이 줄어서'가 아니라 <b>덜 자주 와서</b> — 즉 빈도 문제.",
                    ], "warn")
                else:
                    figd = px.line(mau_s.reset_index(), x="date", y="value",
                                   labels={"value": "VIP MAU(명)", "date": "월"})
                    figd.update_layout(height=300, margin=dict(t=10, b=10), hovermode="x unified")
                    plot(figd, "VIP MAU 장기 추세 (B2B 제외)")
                    st.caption("스티키니스(DAU/MAU) 계산엔 채널별 DAU 파일도 함께 업로드하세요.")

        # ── (2) 채널별 일 DAU: 앱푸시 DAU 하락
        if mon is not None:
            if "PUSH" in mon.columns and "TOTAL" in mon.columns and len(mon) > 1:
                mon["push_share"] = np.where(mon["TOTAL"] > 0, mon["PUSH"] / mon["TOTAL"] * 100, np.nan)
                p0, p1 = mon["PUSH"].iloc[0], mon["PUSH"].iloc[-1]
                dau_sum.update(push_chg=(p1 / p0 - 1) * 100 if p0 else 0,
                               push_share=mon["push_share"].iloc[-1])
                cga, cgb = st.columns(2)
                with cga:
                    _gmode = st.radio("묶음", ["Owned vs Paid", "채널별"], horizontal=True, key="dau_gmode")
                with cgb:
                    _scale = st.radio("표시", ["지수(시작월=100)", "절대값(명)"], horizontal=True, key="pushdau_scale")
                _idx = _scale.startswith("지수")
                OWNED = ["직접", "PUSH", "EP", "미디어커머스"]; PAID = ["광고", "브랜드광고", "제휴"]
                if _gmode.startswith("Owned"):
                    ow = mon[[c for c in OWNED if c in mon.columns]].sum(axis=1)
                    pdd = mon[[c for c in PAID if c in mon.columns]].sum(axis=1)
                    _series = [("VIP 전체 DAU", mon["TOTAL"], "#4C72B0"),
                               ("Owned (자발·앱푸시·직접 등)", ow, "#DD8452"),
                               ("Paid (광고·제휴)", pdd, "#8172B3")]
                else:
                    _cols = [("TOTAL", "VIP 전체", "#4C72B0"), ("PUSH", "앱푸시(owned)", "#DD8452"),
                             ("직접", "직접(자발)", "#55A868"), ("광고", "광고(유료)", "#8172B3")]
                    _series = [(nm, mon[cc], col) for cc, nm, col in _cols if cc in mon.columns]
                figp = go.Figure()
                for _nm, _s, _col in _series:
                    _y = _s / _s.iloc[0] * 100 if _idx else _s
                    figp.add_scatter(x=mon.index, y=_y, name=_nm, mode="lines+markers", line=dict(color=_col, width=2.5))
                if _idx:
                    figp.add_hline(y=100, line=dict(color="#bbb", width=1, dash="dot"))
                figp.update_layout(height=340, margin=dict(t=10, b=10), hovermode="x unified",
                                   legend_title_text="", yaxis=dict(title="지수 (시작월=100)" if _idx else "DAU(명)"))
                plot(figp, "VIP DAU 유입 추세 — Owned vs Paid (B2B 제외)" if _gmode.startswith("Owned")
                     else "VIP DAU 채널별 유입 추세 (B2B 제외)")
                if _gmode.startswith("Owned"):
                    _ow = mon[[c for c in OWNED if c in mon.columns]].sum(axis=1)
                    _pd = mon[[c for c in PAID if c in mon.columns]].sum(axis=1)
                    _ps0 = _pd.iloc[0] / (_ow.iloc[0] + _pd.iloc[0]) * 100
                    _ps1 = _pd.iloc[-1] / (_ow.iloc[-1] + _pd.iloc[-1]) * 100
                    st.caption(f"자발(Owned) {(_ow.iloc[-1]/_ow.iloc[0]-1)*100:+.0f}% · 유료(Paid) {(_pd.iloc[-1]/_pd.iloc[0]-1)*100:+.0f}% "
                               f"→ **유료 의존도 {_ps0:.0f}% → {_ps1:.0f}%**. 빠지는 자발 방문을 유료로 방어하는 구조.")
                st.caption("※ 채널 DAU는 중복 집계(한 방문이 복수 채널에 잡힘)라 합=전체 아님 — '유입 경로별 추세'로 해석.")
                if anom_txt:
                    st.caption(anom_txt + " — 데이터팀 확인 권장")
                yoy_line = ""
                if "dau_yoy" in dau_sum:
                    yoy_line = (f"<b>VIP(B2B 제외) DAU 전년비 {dau_sum['dau_yoy']:+.1f}%</b> — "
                                "내부 관리지표와 동일 기준.")
                insight([
                    yoy_line,
                    f"<b>앱푸시 DAU {fnum(p0)} → {fnum(p1)} ({(p1/p0-1)*100:+.0f}%)</b>로 owned 채널 중 최대 하락. "
                    "직접(자발)도 하락하는데 <b>광고(유료)만 상승</b> → 자발 방문을 유료로 방어 중(로열티 착시·비용 리스크).",
                    "푸시 도달(타겟팅가능)은 정체인데 푸시 DAU는 감소 → 문제는 <b>도달 부족이 아니라 도달 후 반응 저하</b>(1건당 반응률↓). 콘텐츠(전관행사)로도 회복되지 않음.",
                    "자발 회복·콘텐츠 개선이 제약된 상황 → 통제 레버는 <b>발송의 반응 유도력 제고</b>: ① 행동 트리거(D-1→실시간) ② 도달력. "
                    "유료 방어를 owned 정밀 발송으로 전환. (하단 '진단 논리' 참조)",
                ])
    except Exception as e:
        st.warning(f"DAU 데이터 파싱 중 문제: {e} — 파일 형식을 확인해 주세요.")

# ════════════════════════════════════════════════════════════
# 0. 핵심 진단
# ════════════════════════════════════════════════════════════
section("VIP 도달 스냅샷 — 앱푸시(도달력 축)",
        f"기간 {d0} ~ {d1} ({n_days}일) · 도달률=최근일({last_day.date()}) 스냅샷 · 모든 수치 VIP 전용",
        anchor="sec-core")

vip = group_snapshot("VIP", fw_d_last, fl_d, fw_d)   # 등급 필터와 무관하게 항상 VIP 전체
if vip["act_push"]:
    share = vip["unreach"] / vip["act_push"] * 100
    push_net = vip["chnet"].get("PUSH", 0)
    push_out = vip["chout"].get("PUSH", 0)
    kind = "warn" if (vip["reach"] < 50 or push_net < 0) else ""
    push_year = push_net / n_days * 365
    year_pct = abs(push_year) / vip["tot_push"] * 100 if vip["tot_push"] else 0
    proj = (f"현 순감 속도(일평균 {fsigned(push_net / n_days)}명)면 <b>연 △{fnum(abs(push_year))}명</b>"
            f"(현 타겟팅가능의 {year_pct:.0f}%)이 추가로 미도달 — 다만 즉효 레버는 신규 순감보다 "
            f"<b>이미 미도달인 {fnum(vip['unreach'])}명 스톡</b>."
            if push_net < 0 else "")
    insight([
        f"수신동의는 <b>{vip['consent']:.1f}%</b>로 이미 충분 — 병목은 동의가 아니라 <b>앱 보유</b>(타겟팅가능 {vip['reach']:.1f}%). 동의 확보형 캠페인은 효과 한계.",
        f"<b>{fnum(vip['unreach'])}명(동의자의 {share:.0f}%)</b>이 앱 미보유/삭제로 푸시 도달 불가 → 이 풀의 <b>재설치 전환</b>이 VIP DAU 회복의 최대 레버.",
        f"PUSH만 {'순감' if push_net < 0 else '정체'}({fsigned(push_net)})이고 SMS({fsigned(vip['chnet'].get('SMS',0))})·EMAIL({fsigned(vip['chnet'].get('EMAIL',0))})은 순증 → 앱 채널만 약화. 푸시 도달 불가 VIP에는 <b>알림톡/카카오 대체 도달</b> 병행.",
        proj,
    ], kind)

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: metric_card("VIP 푸시 도달률", f"{vip['reach']:.1f}%", f"수신동의율 {vip['consent']:.1f}% → 타겟팅가능")
    with c2: metric_card("VIP 앱 미보유/삭제", fnum(vip["unreach"]), f"수신동의의 {share:.0f}% · 푸시 발송 불가")
    with c3: metric_card("VIP 타겟팅가능_PUSH", fnum(vip["tot_push"]), f"수신동의 {fnum(vip['act_push'])}")
    with c4: metric_card("VIP PUSH 이탈(기간 누적)", fnum(push_out),
                         f"{n_days}일 합계 · 일평균 {fnum(push_out / n_days)}명")
    with c5: metric_card("VIP PUSH 증감(기간 누적)", fsigned(push_net),
                         f"신규−이탈 · 일평균 {fsigned(push_net / n_days)}명 "
                         f"{'순증' if push_net >= 0 else '순감'}")
else:
    st.info("VIP 데이터가 없습니다. 사이드바 그룹 필터에 VIP를 포함해 주세요.")

# ════════════════════════════════════════════════════════════
# 2. 앱푸시 도달률 & 앱 미보유/삭제 (DAU 핵심)
# ════════════════════════════════════════════════════════════
section(f"앱푸시 도달 진단 — 등급별 ({last_day.date()} 기준)",
        "막대=타겟팅가능 vs 앱 미보유/삭제, 라인=도달률(타겟팅가능/수신동의)", anchor="sec-reach")
gp = (fw_last.groupby("grade").agg(act_push=("act_push", "sum"), tot_push=("tot_push", "sum"),
                                   unreach=("unreach_push", "sum")).reindex(grade_order_sel))
gp["reach"] = np.where(gp["act_push"] > 0, gp["tot_push"] / gp["act_push"] * 100, 0)
fig = go.Figure()
fig.add_bar(x=gp.index, y=gp["tot_push"], name="타겟팅가능(앱보유)", marker_color="#55A868")
fig.add_bar(x=gp.index, y=gp["unreach"], name="앱 미보유/삭제", marker_color="#C44E52")
fig.add_scatter(x=gp.index, y=gp["reach"], name="도달률(%)", yaxis="y2",
                mode="lines+markers+text", text=[f"{v:.0f}%" for v in gp["reach"]],
                textposition="top center", line=dict(color="#1a1a2e", width=2))
fig.update_layout(barmode="stack", height=400, margin=dict(t=20, b=10),
                  yaxis=dict(title="수신동의 모수"), legend_title_text="",
                  yaxis2=dict(title="도달률(%)", overlaying="y", side="right", range=[0, 100], showgrid=False),
                  xaxis=dict(categoryorder="array", categoryarray=grade_order_sel))
plot(fig)

gpv = gp.dropna(subset=["reach", "unreach"])
top_vol = gpv["unreach"].idxmax()      # 절대 미보유 규모 최대 → 실질 우선순위
best = gpv["reach"].idxmax()            # 도달률 가장 높은 등급
insight([
    f"재설치 <b>1순위는 {top_vol}</b> — 미보유/삭제 절대 규모가 <b>{fnum(gpv.loc[top_vol,'unreach'])}명</b>으로 최대(도달률 {gpv.loc[top_vol,'reach']:.0f}%). 규모·효율 모두 부합.",
    f"주목할 점: 도달률이 가장 높은 등급 <b>{best}</b>조차 <b>{gpv.loc[best,'reach']:.0f}%</b> — 상위 등급도 절반 가까이 미도달. <b>등급 불문 앱 보유가 공통 병목</b>(단순히 하위 등급 문제가 아님).",
    "발송 타겟에 '<b>앱 보유</b>' 필터 적용 → 미보유 등급은 대체 채널로 전환해 '발송해도 도달하지 않는' 낭비 제거.",
])

# VIP 앱 미보유/삭제 일별 추세 (이탈과 별개 누수) — 등급 필터 무관 VIP 전체
vip_daily = (fw_d[fw_d["group"] == "VIP"].groupby("date")
             .agg(unreach=("unreach_push", "sum"), ap=("act_push", "sum")).reset_index())
if len(vip_daily) > 1:
    vip_daily["pct"] = np.where(vip_daily["ap"] > 0, vip_daily["unreach"] / vip_daily["ap"] * 100, 0)
    figv = go.Figure()
    figv.add_bar(x=vip_daily["date"], y=vip_daily["unreach"], name="앱 미보유/삭제(명)",
                 marker_color="#e7c3c3", marker_line_width=0, opacity=0.75)
    figv.add_scatter(x=vip_daily["date"], y=vip_daily["pct"], name="미보유 비율(동의 대비)", yaxis="y2",
                     mode="lines+markers", line=dict(color="#8C3A3A", width=2.5))
    figv.update_layout(height=300, margin=dict(t=10, b=10), hovermode="x unified", legend_title_text="",
                       yaxis=dict(title="앱 미보유/삭제(명)"),
                       yaxis2=dict(title="미보유 비율(%)", overlaying="y", side="right",
                                   showgrid=False, tickformat=".1f"))
    plot(figv, "VIP 앱 미보유/삭제 일별 추세")
    st.caption("앱 미보유/삭제 = 수신동의 − 타겟팅가능(PUSH). 수신거부(이탈)와 다른 누수 — 동의는 유지하지만 "
               "앱 미보유로 도달 불가한 모수. (VIP 등급 데이터는 6/15부터라 추세가 짧음)")

# ════════════════════════════════════════════════════════════
# 3. 수신거부 분석
# ════════════════════════════════════════════════════════════
section("수신거부(이탈) 분석", f"채널: {', '.join(sel_channels)} · 등급 라벨 순서 SP→RD",
        anchor="sec-optout")
cc1, cc2 = st.columns([1.3, 1])
with cc1:
    log_y = st.toggle("로그 스케일", value=True, key="optout_log",
                      help="PUSH가 SMS/EMAIL보다 훨씬 커서, 작은 채널도 보이도록 로그 스케일 권장")
    daily_ch = fl.groupby(["date", "channel"])["out"].sum().reset_index()
    figc = px.line(daily_ch, x="date", y="out", color="channel", markers=True,
                   color_discrete_map=CH_COLOR, labels={"out": "수신거부", "date": "일자", "channel": "채널"})
    figc.update_layout(height=320, margin=dict(t=10, b=10), hovermode="x unified", legend_title_text="")
    if log_y:
        figc.update_yaxes(type="log")
    plot(figc, "일별 채널별 수신거부 추이")
with cc2:
    # 수신거부율 = 기간 거부 합계 / 최근일 수신자수(선택 채널 합) — 채널 간 max 혼용 제거
    lc = fl["date"].max()
    last_tot = fl[fl["date"] == lc].groupby("grade")["tot"].sum()
    grade_out = fl.groupby("grade")["out"].sum().to_frame("out")
    grade_out["tot"] = last_tot
    grade_out = grade_out.reindex(grade_order_sel)
    grade_out["rate"] = np.where(grade_out["tot"] > 0, grade_out["out"] / grade_out["tot"] * 100, 0)
    go_df = grade_out.reset_index()
    figo = px.bar(go_df, x="grade", y="rate", color="rate", color_continuous_scale="OrRd",
                  labels={"rate": "수신거부율(%)", "grade": "등급"},
                  text=go_df["out"].map(lambda v: f"이탈 {int(v):,}"))
    figo.update_traces(textposition="outside", textfont_size=10, cliponaxis=False)
    figo.update_layout(height=350, margin=dict(t=10, b=10), coloraxis_showscale=False,
                       yaxis=dict(tickformat=".1f"),
                       xaxis=dict(categoryorder="array", categoryarray=grade_order_sel))
    plot(figo, "등급별 수신거부율 (전채널)")
    st.caption("수신거부율 = 기간 누적 이탈 ÷ 최근일 수신자수(선택 채널 합) × 100. "
               "※ 모수 작은 상위 등급(SP·PT)은 이탈 몇 건에도 율의 변동성이 큼 — 절대 이탈 건수를 함께 확인(실제 물량은 RD·BK·PP).")

# 등급 × 채널 히트맵
metric_choice = st.radio("히트맵 지표", ["수신거부 건수", "수신거부율(%)", "순증감"],
                         horizontal=True, label_visibility="collapsed")
po = fl.pivot_table(index="grade", columns="channel", values="out", aggfunc="sum", fill_value=0)
pt = fl.pivot_table(index="grade", columns="channel", values="tot", aggfunc="max", fill_value=0)
pn = fl.pivot_table(index="grade", columns="channel", values="new", aggfunc="sum", fill_value=0)
ch_order = [c for c in CHANNELS if c in sel_channels]
po, pt, pn = (p.reindex(index=grade_order_sel, columns=ch_order) for p in (po, pt, pn))
if metric_choice == "수신거부 건수":
    z, fmt, cs = po, ",.0f", "Reds"
elif metric_choice == "수신거부율(%)":
    z, fmt, cs = (po / pt.replace(0, np.nan) * 100).fillna(0), ".1f", "OrRd"
else:
    z, fmt, cs = (pn - po), ",.0f", "RdYlGn"
figh = px.imshow(z, text_auto=fmt, aspect="auto", color_continuous_scale=cs,
                 labels=dict(x="채널", y="등급", color=metric_choice))
figh.update_layout(height=max(300, 42 * len(grade_order_sel) + 80), margin=dict(t=20, b=10))
plot(figh)

ch_o = fl.groupby("channel")["out"].sum()
ch_n = fl.groupby("channel")["new"].sum()
push_net2 = ch_n.get("PUSH", 0) - ch_o.get("PUSH", 0)
insight([
    f"앱 방문(DAU)에 관여하는 채널은 <b>PUSH</b>(증감 {fsigned(push_net2)}) — SMS/EMAIL 순증은 앱 방문 기여가 적어, 수신거부 방어 우선순위는 PUSH.",
    "등급별 '수신거부율'은 모수 작은 상위 등급에서 변동성이 큰 노이즈 → 방어 대상은 율이 아니라 <b>절대 이탈 건수(RD·BK·PP)</b> 기준으로 선정.",
])

# ════════════════════════════════════════════════════════════
# 4. 그룹 내 등급별 인사이트
# ════════════════════════════════════════════════════════════
section("그룹 내 등급별 인사이트", "각 그룹 안에서 등급 간 도달·이탈 비교", anchor="sec-within")
for grp in sel_groups:
    members = [g for g in GROUPS[grp] if g in grade_order_sel]
    if not members:
        continue
    st.markdown(f"**{grp}** &nbsp; <span style='color:#888;font-size:13px'>{', '.join(members)}</span>",
                unsafe_allow_html=True)
    snap = fw_last[fw_last["grade"].isin(members)].groupby("grade").agg(
        act_push=("act_push", "sum"), tot_push=("tot_push", "sum"), unreach=("unreach_push", "sum")
    ).reindex(members)
    snap["reach"] = np.where(snap["act_push"] > 0, snap["tot_push"] / snap["act_push"] * 100, 0)
    outp = fl[fl["grade"].isin(members)].groupby("grade").agg(
        out=("out", "sum"), new=("new", "sum"), tot=("tot", "max")).reindex(members)
    outp["rate"] = np.where(outp["tot"] > 0, outp["out"] / outp["tot"] * 100, 0)
    outp["net"] = outp["new"] - outp["out"]

    cols = st.columns(len(members))
    for col, g in zip(cols, members):
        with col:
            metric_card(g, f"{snap.loc[g,'reach']:.1f}%",
                        f"미보유/삭제 {fnum(snap.loc[g,'unreach'])}<br>"
                        f"거부 {fnum(outp.loc[g,'out'])} · 순증감 {fsigned(outp.loc[g,'net'])}",
                        color=GROUP_COLOR[grp])

    lo_reach = snap["reach"].idxmin()
    hi_out = outp["rate"].idxmax()
    neg_net = outp[outp["net"] < 0].index.tolist()
    bullets = [
        f"{grp} 내 도달률 최저 <b>{lo_reach}</b>({snap.loc[lo_reach,'reach']:.1f}%) → 그룹 안에서도 재설치 우선 타겟.",
        (f"구독 순감 등급: <b>{', '.join(neg_net)}</b> — 푸시 리텐션 점검 필요(신규<이탈)."
         if neg_net else "전 등급 구독 순증 — 현 푸시 유지 흐름은 양호."),
    ]
    insight(bullets, "warn" if (snap["reach"].min() < 40 or neg_net) else "")

# ════════════════════════════════════════════════════════════
# 1. 그룹 비교 (VIP vs 일반)
# ════════════════════════════════════════════════════════════
section("그룹 비교 — VIP (일반은 참고)",
        "본 진단 목표는 VIP DAU · 일반은 범위 밖 참고용(별도 관제 대상)", anchor="sec-group")
gcols = st.columns(2)
gsnap = {}
for col, grp in zip(gcols, ["VIP", "일반"]):
    gs = group_snapshot(grp, fw_d_last, fl_d, fw_d)
    gsnap[grp] = gs
    share = (gs["unreach"] / gs["act_push"] * 100) if gs["act_push"] else 0
    cn = gs["chnet"]
    _ref = ' <span style="font-size:11px;color:#aaa">(참고·범위 밖)</span>' if grp == "일반" else ""
    with col:
        st.markdown(
            f'<div class="metric-card" style="border-left-color:{GROUP_COLOR[grp]}">'
            f'<div class="metric-value" style="font-size:19px">{grp}{_ref} '
            f'<span style="font-size:13px;color:#888">({", ".join(GROUPS[grp])})</span></div>'
            f'<div class="metric-sub">유효회원 {fnum(gs["act"])} → 수신동의 <b>{gs["consent"]:.1f}%</b> '
            f'→ 타겟팅가능 <b>{gs["reach"]:.1f}%</b> ({fnum(gs["tot_push"])})</div>'
            f'<div class="metric-sub">앱 미보유/삭제 <b>{fnum(gs["unreach"])}</b> ({share:.0f}%)</div>'
            f'<div class="metric-sub">증감 — PUSH <b>{fsigned(cn.get("PUSH",0))}</b> · '
            f'SMS <b>{fsigned(cn.get("SMS",0))}</b> · EMAIL <b>{fsigned(cn.get("EMAIL",0))}</b></div>'
            f'</div>', unsafe_allow_html=True)

# 그룹별 도달률 일별 추이
grp_daily = (fw_d.groupby(["date", "group"]).agg(tp=("tot_push", "sum"), ap=("act_push", "sum")).reset_index())
grp_daily["reach"] = np.where(grp_daily["ap"] > 0, grp_daily["tp"] / grp_daily["ap"] * 100, 0)
figr = px.line(grp_daily, x="date", y="reach", color="group", markers=True,
               color_discrete_map=GROUP_COLOR, labels={"reach": "앱푸시 도달률(%)", "date": "일자", "group": "그룹"})
figr.update_layout(height=320, margin=dict(t=20, b=10), hovermode="x unified", legend_title_text="")
plot(figr, "그룹별 앱푸시 도달률 추이")

if "VIP" in gsnap:
    v = gsnap["VIP"]
    vp = v["chnet"].get("PUSH", 0)
    v_share = (v["unreach"] / v["act_push"] * 100) if v["act_push"] else 0
    insight([
        f"VIP는 앱 미보유/삭제 <b>{fnum(v['unreach'])}명(동의자의 {v_share:.0f}%)</b>이 도달을 막음 — "
        f"본 진단의 목표는 <b>VIP DAU</b>이므로 판단·액션은 VIP 기준으로 한정.",
        f"VIP PUSH <b>{fsigned(vp)}</b>({'순감' if vp < 0 else '정체'}) → VIP 전용 재설치·리텐션이 과제.",
    ])

# ════════════════════════════════════════════════════════════
# 0.5 장기 추세 (전체·등급무관) — 별도 히스토리 데이터(25.1.1~)
# ════════════════════════════════════════════════════════════
LT = load_longterm()
if LT is not None:
    section("장기 추세 (전체·등급무관)", "25.1.1~26.6.25 일별 · 전체회원 기준(등급 구분 불가) · VIP는 절대 도달률↑이나 동일 하락 압력",
            anchor="sec-trend")
    # 이상치(특정일 풀 리카운트 등 급등/급락) 제외 — 롤링 중앙값 대비 과대 편차 날짜 NaN 처리
    raw_act = lt_series(LT, "MEMBERSHIP", "전체유효회원")
    raw_tp = lt_series(LT, "MEMBERSHIP", "타겟팅가능")
    act_s = drop_outliers(raw_act)
    tp_s = drop_outliers(raw_tp)
    reach_s = (tp_s / act_s * 100).dropna()
    excl = sorted(set(raw_tp.index[raw_tp.notna() & tp_s.isna()]) |
                  set(raw_act.index[raw_act.notna() & act_s.isna()]))
    excl_txt = ("제외된 날: " + ", ".join(d.strftime("%Y-%m-%d") for d in excl[:5]) +
                (f" 외 {len(excl)-5}일" if len(excl) > 5 else "")) if excl else "제외된 이상치 없음"
    if len(reach_s) > 1:
        r0, r1 = reach_s.iloc[0], reach_s.iloc[-1]
        a0, a1 = act_s.dropna().iloc[0], act_s.dropna().iloc[-1]
        t0, t1 = tp_s.dropna().iloc[0], tp_s.dropna().iloc[-1]
        insight([
            f"회원 <b>+{(a1/a0-1)*100:.0f}%</b> 증가에도 푸시 타겟팅가능은 정체({fnum(t0)}→{fnum(t1)}) → <b>신규 획득이 도달 자산으로 축적되지 않음</b>. 획득 KPI와 별개로 '앱 보유 전환'을 지표화해야.",
            f"도달률 {r0:.1f}%→{r1:.1f}% 하락은 구조적 — 일시 캠페인으론 못 막음 → <b>온보딩·앱 설치 단계의 구조 개선</b>이 본질 레버.",
        ], "warn")

    # 회원 증가 vs 푸시 도달률 (이중축)
    figt = go.Figure()
    # 막대(전체유효회원)는 기본 축(배경), 도달률 선은 오버레이 축(y2)에 둬서 막대 위로 렌더
    figt.add_bar(x=act_s.index, y=act_s.values, name="전체유효회원",
                 marker_color="#dbe3ef", marker_line_width=0, opacity=0.65)
    figt.add_scatter(x=reach_s.index, y=reach_s.values, name="푸시 도달률(%)", yaxis="y2",
                     mode="lines", line=dict(color="#C44E52", width=3), connectgaps=True)
    figt.update_layout(height=340, margin=dict(t=10, b=10), hovermode="x unified", legend_title_text="",
                       yaxis=dict(title="전체유효회원"),
                       yaxis2=dict(title="푸시 도달률(%)", overlaying="y", side="right", showgrid=False))
    plot(figt, "회원수는 ↑, 푸시 도달률은 ↓")
    st.caption(f"※ 풀 리카운트 등 비정상적으로 급등하는 날은 정확도를 위해 자동 제외(롤링 중앙값 대비 과대 편차). {excl_txt}")

    # 앱 미보유/삭제(전체) 추세 = 수신동의 − 타겟팅가능 (이탈과 별개 누수)
    consent_s = drop_outliers(lt_series(LT, "MEMBERSHIP", "수신동의"))
    gap_s = (consent_s - tp_s).dropna()
    gap_pct = (gap_s / consent_s * 100).dropna()
    if len(gap_s) > 1:
        g0, g1 = gap_s.iloc[0], gap_s.iloc[-1]
        p0, p1 = gap_pct.iloc[0], gap_pct.iloc[-1]
        insight([
            f"동의는 유지되나 앱 없는 풀이 <b>+{(g1/g0-1)*100:.0f}%</b>({fnum(g1)})로 누적 → 발송량을 늘려도 도달은 안 늘어남(이미 발송 대상에서 제외됨).",
            f"이탈(수신거부)보다 앱 미보유/삭제가 훨씬 큰 누수(수신동의 대비 {p1:.1f}%) → <b>수신거부 방어보다 앱 재설치·유지가 우선순위</b>.",
        ])
        figg = go.Figure()
        figg.add_bar(x=gap_s.index, y=gap_s.values, name="앱 미보유/삭제(명)",
                     marker_color="#e7c3c3", marker_line_width=0, opacity=0.7)
        figg.add_scatter(x=gap_pct.index, y=gap_pct.values, name="수신동의 대비(%)", yaxis="y2",
                         mode="lines", line=dict(color="#8C3A3A", width=3), connectgaps=True)
        figg.update_layout(height=300, margin=dict(t=10, b=10), hovermode="x unified", legend_title_text="",
                           yaxis=dict(title="앱 미보유/삭제(명)"),
                           yaxis2=dict(title="수신동의 대비(%)", overlaying="y", side="right", showgrid=False))
        plot(figg, "앱 미보유/삭제 추세 (전체) — 수신동의했으나 앱 미보유로 도달 불가한 모수")

    tcol1, tcol2 = st.columns(2)
    with tcol1:
        # 채널별 타겟팅가능 모수 — 시작=100 지수화
        idx_rows = []
        for ch in CHANNELS:
            s = drop_outliers(lt_series(LT, ch, "수신동의"))
            base = s.dropna()
            if len(base):
                idx_rows.append(pd.DataFrame({"date": s.index, "idx": s.values / base.iloc[0] * 100, "채널": ch}))
        if idx_rows:
            idf = pd.concat(idx_rows)
            figi = px.line(idf, x="date", y="idx", color="채널", color_discrete_map=CH_COLOR,
                           labels={"idx": "지수(시작=100)", "date": "일자"})
            figi.update_traces(connectgaps=True)
            figi.update_layout(height=320, margin=dict(t=10, b=10), hovermode="x unified", legend_title_text="")
            plot(figi, "채널별 타겟팅가능 모수 (시작=100)")
    with tcol2:
        # 채널별 월별 증감(신규추가−기존이탈) — 이상치 제외 후 월 합산
        net_rows = []
        for ch in CHANNELS:
            s = drop_outliers(lt_series(LT, ch, "증감"))
            if len(s):
                m = s.resample("MS").sum(min_count=1)
                net_rows.append(pd.DataFrame({"month": m.index, "증감": m.values, "채널": ch}))
        if net_rows:
            ndf = pd.concat(net_rows).dropna(subset=["증감"])
            fign = px.bar(ndf, x="month", y="증감", color="채널", barmode="group",
                          color_discrete_map=CH_COLOR, labels={"month": "월", "증감": "월 증감"})
            fign.update_layout(height=320, margin=dict(t=10, b=10), legend_title_text="")
            plot(fign, "채널별 월 증감 (신규추가−기존이탈)")

# ════════════════════════════════════════════════════════════
# 4.4 진단 논리 — 소거법에서 최종 방향까지
# ════════════════════════════════════════════════════════════
section("진단 논리 — 소거법에서 방향까지",
        "「DAU 개선 방안」 Summ 1.0 → 3.0 의 분석 흐름", anchor="sec-logic")
st.markdown("""
<div class="logic">
  <div class="logic-step">
    <div class="logic-num">1</div>
    <div class="logic-card">
      <b>문제 정의 & 현상</b><span class="tag">Summ 1.0~1.1</span><br>
      MAU는 유지·증가인데 DAU는 지속 역신장 → <b>방문 고객수가 아닌 방문 빈도의 문제</b>.
      빈도가 하향 이동(고빈도 구성비 34.1%→30.5%, 저빈도 46.1%→49.8%).
      행동 근거: 탐색(검색·상품상세)은 유지·증가인데 <b>구매 의도 행동(쇼핑백·주문)은 약화</b>
      (저빈도 쇼핑백 유입 △47%, 주문완료 △66% / 검색→쇼핑백 전환율 하락).
    </div>
  </div>
  <div class="logic-step">
    <div class="logic-num">2</div>
    <div class="logic-card">
      <b>일자 유형 검증</b><span class="tag">Summ 1.2</span><br>
      전관행사·비전관·평일·주말 <b>모두 △9%대 하락</b> → 특정 일자 유형 부진이 아니라 <b>전체 방문 베이스 하락</b>.
      전관행사는 DAU 견인율 <b>△0.1%</b>(유입 확대 X)·거래액 견인 +11.5% → <b>유입이 아닌 구매효율 장치</b>.
      <div class="verdict warn">✕ '무슨 행사·혜택을 더 하나'로는 DAU 안 풀림</div>
    </div>
  </div>
  <div class="logic-step">
    <div class="logic-num">3</div>
    <div class="logic-card">
      <b>종합 진단 — 핵심은 '고빈도의 저빈도화'</b><span class="tag">Summ 2.0</span><br>
      코호트로 보면 <b>작년 고빈도였던 VIP의 38%가 올해 중·저빈도로 이동</b>(중빈도의 49%도 저빈도화) — <b>동일 인물의 열화</b>.
      원래 자주 오던 사람이 덜 오면서 방문·구매가 함께 감소.
      DAU는 복합지표라 단일 원인 특정은 불가 → 원인 규명보다 <b>확인된 약화 구간에 파일럿을 걸어 실측</b>.
    </div>
  </div>
  <div class="logic-step final">
    <div class="logic-num">✓</div>
    <div class="logic-card ok">
      <b>방향 — 통제 가능한 최대 레버 (no-regret)</b><span class="tag">Summ 3.0</span><br>
      <b>전제:</b> 자발적 빈도 회복은 사실상 불가(경쟁 심화·상품/UX는 타부서·즉시개선 난망) → <b>발송(터치)이 실제 반응으로 이어져야</b> 함.
      관점 전환: AS-IS <i>"무엇을 만들까"</i> → TO-BE <b>"재방문 계기를 적시·적합하게 전달하는가"</b>.<br>
      ① <b>행동 기반 트리거</b> — 현재 발송이 <b>D-1(하루 전) 데이터</b> 기반이라 조회·장바구니·위시 등 관심이 살아있는 순간을 놓침
      → <b>처방(개입) 자체가 손발 묶인 상태</b>. 행동 발생 시점 기반으로 풀어야 터치가 먹힘.<br>
      ② <b>채널 접점·도달력</b> — 그 터치가 실제로 닿는 상태 유지 (앞의 도달 진단 = 이 축).<br>
      <span style="font-size:12.5px;color:#3a6b47">※ 'DAU 최대 원인'이 아니라 <b>'통제 가능한 최대 레버'</b> — 원인 규명과 무관하게 당길 가치가 있는 no-regret 선택. 효과는 파일럿으로 실측.</span>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)
st.caption("출처: 「DAU 개선 방안」 Summ 1.0~3.0 · 결론은 '원인'이 아니라 통제 가능한 최대 레버로 제시 — 효과는 CRM 파일럿으로 실측.")

st.markdown("""
<div class="opinion">
  <div class="ohead">🔍 독립 검토 의견 (외부 관점) — 위 진단을 반박·보완</div>
  <ol>
    <li><b>문제 크기 재보정</b> — 전체 DAU도 −10%대 동반 하락(전 등급 기준), 동일 소스 비교상 VIP는 오히려 <b>덜 빠짐</b>.
        하락이 VIP 고유가 아니라 시장·플랫폼 광범위일 가능성 → "VIP 로열티 붕괴"보다 <b>"시장 하락 속 VIP 방어"</b>가 더 정확·유리(임원 방어 근거).</li>
    <li><b>빈도 하락 = 전부 충성층 이탈?</b> — VIP MAU +9.4%인데 하위 등급 MAU는 감소 → <b>등급 승급 유입(믹스)</b> 유력.
        신규 편입 VIP가 저빈도면 빈도는 기계적으로 하락. <span class="chk">→ 신규 VIP vs 기존 VIP 빈도 분리 필요: win-back이냐 신규 온보딩이냐로 액션이 갈림. (현 결론은 암묵적으로 '기존층 이탈'만 가정)</span></li>
    <li><b>인과 방향 점검</b> — '빈도↓→구매의도↓'가 아니라 <b>'구매의도↓(구색·가격 경쟁력)→빈도↓'</b>일 수 있음(직접 −12% / 광고 +8%가 방증).
        MD 영역이라 CRM 통제 밖이나, <b>VIP 이탈 집중 카테고리를 태깅해 MD/기획에 리텐션 근거로 이관</b>은 CRM이 할 수 있는 역할.</li>
    <li><b>지표 재검토</b> — 총 DAU는 광고로 저관여 트래픽 사는 유인이 됨. <b>VIP의 '구매 의도 있는 재방문'</b>을 보조지표로 병행.</li>
    <li><b>실행 순서</b> — D-1→실시간은 <b>적재 요청 시 착수 가능</b>하나 효과는 미검증 → <b>기존 D-1로 타겟팅 정교화 A/B 선검증</b> 후 확대(타이밍이 진짜 병목인지 값싸게 확인).</li>
  </ol>
  <div class="chk" style="margin-top:8px">종합: 방향(행동 트리거 + 도달력)은 CRM 범위 내에서 타당. 단 ① 문제 크기 과대평가 가능 ② 신규/기존 VIP 미분리 ③ 지표·투자순서 리스크 — <b>다음 급소 = 믹스(신규/기존 VIP) 분리.</b> (제3자 관점, 데이터 검증 필요)</div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div style="font-weight:700;font-size:15px;margin:14px 0 2px">🎛 CRM 레버 인벤토리 — 영향도(판단) × 착수 여부</div>',
            unsafe_allow_html=True)
st.markdown("""
<div style="max-width:720px">
<svg width="100%" viewBox="0 0 640 410" role="img" style="height:auto;font-family:'Noto Sans KR','Malgun Gothic',sans-serif;display:block">
<title>CRM 레버 영향도 × 착수여부 매트릭스</title>
<text x="205" y="34" text-anchor="middle" fill="#854F0B" font-size="11" font-weight="600">미착수 · 부분</text>
<text x="486" y="34" text-anchor="middle" fill="#5F5E5A" font-size="11" font-weight="600">운영 중</text>
<text transform="rotate(-90 28 216)" x="28" y="216" text-anchor="middle" fill="#5F5E5A" font-size="11">낮음 ↓        영향도 (판단)        ↑ 높음</text>
<rect x="64" y="44" width="562" height="344" rx="4" fill="none" stroke="#D3D1C7" stroke-width="1"/>
<line x1="345" y1="44" x2="345" y2="388" stroke="#D3D1C7" stroke-width="1"/>
<line x1="64" y1="216" x2="626" y2="216" stroke="#D3D1C7" stroke-width="1"/>
<rect x="76" y="58" width="258" height="50" rx="6" fill="#FAC775" stroke="#854F0B" stroke-width="2"/>
<text x="88" y="80" fill="#412402" font-size="12" font-weight="600">행동 시점 정밀도 (D-1 → 실시간)</text>
<text x="88" y="97" fill="#633806" font-size="10">현재 발송 D-1 기준 (Summ 3.0) · 요청 시 실시간 가능</text>
<rect x="76" y="118" width="258" height="44" rx="6" fill="#FAEEDA" stroke="#BA7517" stroke-width="1"/>
<text x="88" y="137" fill="#633806" font-size="12" font-weight="600">발송 관련성·피로도 재배분</text>
<text x="88" y="152" fill="#854F0B" font-size="10">앱푸시 6회 · 문자 2회/일 (Summ 3.0)</text>
<rect x="357" y="58" width="258" height="44" rx="6" fill="#F1EFE8" stroke="#C9C7BD" stroke-width="1"/>
<text x="369" y="77" fill="#444441" font-size="12" font-weight="600">휴면·이탈 타겟 자동화</text>
<text x="369" y="92" fill="#5F5E5A" font-size="10">30일 미방문 자동 발송 (운영 중)</text>
<rect x="76" y="240" width="258" height="48" rx="6" fill="#FAEEDA" stroke="#BA7517" stroke-width="1"/>
<text x="88" y="261" fill="#633806" font-size="12" font-weight="600">미보유 재설치</text>
<text x="88" y="278" fill="#854F0B" font-size="10">미도달 16.8만 · 10% 전환 시 손익분기 부근</text>
<rect x="357" y="230" width="258" height="40" rx="6" fill="#F1EFE8" stroke="#C9C7BD" stroke-width="1"/>
<text x="369" y="248" fill="#444441" font-size="12" font-weight="600">혜택·프로모션 (전관행사)</text>
<text x="369" y="262" fill="#5F5E5A" font-size="10">DAU 견인 △0.1% (Summ 1.2)</text>
<rect x="357" y="276" width="258" height="40" rx="6" fill="#F1EFE8" stroke="#C9C7BD" stroke-width="1"/>
<text x="369" y="294" fill="#444441" font-size="12" font-weight="600">발송량·빈도</text>
<text x="369" y="308" fill="#5F5E5A" font-size="10">앱푸시 6회/일 · 포화</text>
<rect x="357" y="322" width="258" height="40" rx="6" fill="#F1EFE8" stroke="#C9C7BD" stroke-width="1"/>
<text x="369" y="340" fill="#444441" font-size="12" font-weight="600">채널 도달 (앱푸시)</text>
<text x="369" y="354" fill="#5F5E5A" font-size="10">타겟팅가능 78k ≈ MAU 81k (활성 커버)</text>
</svg>
</div>
""", unsafe_allow_html=True)
st.caption("주황 = 미착수·부분(남은 레버) · 회색 = 운영 중. 세로축 영향도는 판단(데이터 아님). "
           "운영 중 레버의 DAU 효과는 데이터상 제한적(전관행사 견인 △0.1%, 앱푸시 6회/일 포화) → 남은 착수 대상은 행동 시점 정밀도(현 D-1)·발송 관련성.")

st.markdown('<div style="font-weight:700;font-size:15px;margin:16px 0 4px">📉 CRM 레버별 개선 시 커버 (매트릭스 요소 · 커버 큰 순)</div>',
            unsafe_allow_html=True)
st.markdown("""
<table style="width:100%;max-width:760px;border-collapse:collapse;font-size:12.5px;color:#2C2C2A">
<thead><tr style="background:#EDEBE4;color:#444441">
<th style="text-align:left;padding:8px 10px">CRM 레버</th>
<th style="padding:8px 10px;white-space:nowrap">개선 시 커버(상한)</th>
<th style="padding:8px 10px;white-space:nowrap">착수 상태</th>
</tr></thead>
<tbody>
<tr style="border-bottom:1px solid #E0DED6;background:#FEF7E9">
<td style="padding:10px"><b>행동 시점 정밀도 · 발송 관련성</b><div style="font-size:11px;color:#777;margin-top:2px">고객 행동이 하루 뒤(D-1)에야 발송에 반영돼 관심이 식은 뒤 도착. 실시간 전환 시 개선 여지.</div></td>
<td style="padding:10px;text-align:center"><span style="background:#FAC775;color:#633806;padding:2px 10px;border-radius:10px">~45%</span><div style="font-size:11px;color:#888;margin-top:3px">앱푸시 회복 시</div></td>
<td style="padding:10px;text-align:center;color:#854F0B;font-weight:600">미착수<div style="font-size:11px;margin-top:2px">← 다음 액션</div></td>
</tr>
<tr style="border-bottom:1px solid #E0DED6">
<td style="padding:10px"><b>미보유 재설치 · 채널 도달</b><div style="font-size:11px;color:#777;margin-top:2px">앱이 없는 고객은 방문 자체가 불가. 재설치 전환율이 낮아(10% 전환도 손익분기) 회복 폭 제한.</div></td>
<td style="padding:10px;text-align:center"><span style="background:#F1EFE8;color:#5F5E5A;padding:2px 10px;border-radius:10px">~10%</span><div style="font-size:11px;color:#888;margin-top:3px">재설치 1% 가정</div></td>
<td style="padding:10px;text-align:center;color:#5F5E5A">부분</td>
</tr>
<tr>
<td style="padding:10px"><b>전관행사 · 발송량 · 휴면 자동화</b><div style="font-size:11px;color:#777;margin-top:2px">이미 최대로 운영 중이라 추가 여력 없음. 전관행사도 DAU 견인 △0.1%(사실상 0, Summ 1.2).</div></td>
<td style="padding:10px;text-align:center"><span style="background:#F1EFE8;color:#5F5E5A;padding:2px 10px;border-radius:10px">~0%</span><div style="font-size:11px;color:#888;margin-top:3px">이미 소진</div></td>
<td style="padding:10px;text-align:center;color:#5F5E5A">운영 중</td>
</tr>
</tbody></table>
""", unsafe_allow_html=True)
st.caption("커버 = 역신장(=100%) 대비 이 CRM 레버로 되돌릴 수 있는 최대치(채널 기여 기준·상한). 앱푸시=역신장의 ~45% · 재설치는 1% 가정 · 실제 회복폭은 파일럿 실측.")
insight([
    "CRM 레버 중 <b>개선 여지가 남은 건 '행동 시점 정밀도(D-1→실시간)' — 커버 상한 ~45%</b>. 운영 중 레버(전관행사·발송량·휴면·도달)는 이미 소진(~0%), 재설치는 바운드(~10%).",
    "→ CRM 최선은 접점 개선으로 <b>하락 방어(반전 아님)</b>. 재설치 레버 상세는 결론 하단 시뮬레이터.",
])

# ════════════════════════════════════════════════════════════
# 4.5 결론 — 인사이트·시사점·액션 (컨설팅식 종합)
# ════════════════════════════════════════════════════════════
section("인사이트 · 시사점 · 액션",
        "현황 진단 → 핵심 문제 → 시사점 → 권고 액션 (VIP·앱푸시 기준 종합)", anchor="sec-action")

if vip["act_push"]:
    a_share = vip["unreach"] / vip["act_push"] * 100
    a_pnet = vip["chnet"].get("PUSH", 0)
    # 장기(전체) 요약 재계산
    lt_reach_txt = lt_gap_txt = ""
    if LT is not None:
        _a = drop_outliers(lt_series(LT, "MEMBERSHIP", "전체유효회원"))
        _tp = drop_outliers(lt_series(LT, "MEMBERSHIP", "타겟팅가능"))
        _c = drop_outliers(lt_series(LT, "MEMBERSHIP", "수신동의"))
        _reach = (_tp / _a * 100).dropna()
        _gap = (_c - _tp).dropna()
        _gp = (_gap / _c * 100).dropna()
        _ad = _a.dropna()
        if len(_reach) > 1 and len(_ad) > 1:
            _mg = (_ad.iloc[-1] / _ad.iloc[0] - 1) * 100
            lt_reach_txt = (f"전체 푸시 도달률 <b>{_reach.iloc[0]:.1f}%→{_reach.iloc[-1]:.1f}%</b> 하락, "
                            f"회원 +{_mg:.0f}% 증가에도 타겟팅가능 정체")
        if len(_gap) > 1:
            lt_gap_txt = f"앱 미보유/삭제(전체) <b>{fnum(_gap.iloc[-1])}</b>(수신동의의 {_gp.iloc[-1]:.0f}%)로 확대"

    has_dau = bool(dau_sum)
    asis = []
    if has_dau:
        _p = []
        if "dau_yoy" in dau_sum:
            _p.append(f"VIP DAU 전년비 <b>{dau_sum['dau_yoy']:+.0f}%</b>({dau_sum.get('yoy_basis','')} 기준)")
        if "stick1" in dau_sum:
            _p.append(f"스티키니스(DAU/MAU) <b>{dau_sum['stick0']:.0f}%→{dau_sum['stick1']:.0f}%</b>")
        if "push_chg" in dau_sum:
            _p.append(f"앱푸시 DAU <b>{dau_sum['push_chg']:+.0f}%</b>(VIP DAU의 ~{dau_sum.get('push_share',0):.0f}%)")
        if _p:
            asis.append("실측 DAU — " + " · ".join(_p) + ".")
    asis += [
        f"VIP 수신동의는 <b>{vip['consent']:.1f}%</b>로 확보됐으나 실제 앱푸시 <b>타겟팅가능은 {vip['reach']:.1f}%</b>"
        f"({fnum(vip['tot_push'])}/{fnum(vip['act_push'])})에 불과.",
        f"동의자 중 <b>{fnum(vip['unreach'])}명({a_share:.0f}%)</b>이 앱 미보유/삭제 상태, "
        f"VIP PUSH 증감 {fsigned(a_pnet)}(순{'감' if a_pnet < 0 else '증'}).",
        (f"{lt_reach_txt} · {lt_gap_txt}." if lt_reach_txt else ""),
    ]
    insight(asis, cap="📌 현황 진단 (As-Is)")

    if has_dau:
        prob = []
        if "stick1" in dau_sum:
            prob.append(
                f"DAU 하락의 실체는 <b>스티키니스(DAU/MAU) 하락</b>({dau_sum['stick0']:.0f}%→{dau_sum['stick1']:.0f}%) — "
                f"VIP MAU는 {dau_sum.get('mau_growth',0):+.0f}%로 <b>유지·증가</b>. '앱 쓰는 사람이 줄어서'(모수 축소)가 아님.")
        if "push_chg" in dau_sum:
            prob.append(
                f"<b>앱푸시 DAU {dau_sum['push_chg']:+.0f}%</b>로 최대 하락(VIP DAU의 ~{dau_sum.get('push_share',0):.0f}% 견인). "
                "푸시 도달(타겟팅가능)은 flat인데 DAU가 빠짐 → <b>1건당 반응률(재방문 전환) 하락</b>.")
        prob += [
            "콘텐츠 레버 소진 — 전관행사(최대 혜택)로도 회복 안 됨. 반응률·빈도는 단기 통제 어려움.",
            "<span style='color:#888'>※ 이 레버(정밀 트리거·도달력)의 DAU 회복 효과는 <b>파일럿으로 실측 필요</b> — 'DAU 최대 원인'이 아니라 통제 가능한 최대 레버(no-regret)로 제시.</span>",
        ]
    else:
        prob = [
            "병목은 '수신 동의'가 아니라 '<b>앱 보유</b>' — 동의는 98%+지만 실제 도달은 1/3 수준.",
            "<b>앱 미보유/삭제 풀이 구조적으로 확대</b> → 앱 도달 가능 모수가 회원 증가를 못 따라가 정체.",
            "SMS/EMAIL은 순증하는데 <b>PUSH(앱 채널)만 순감</b> → 앱 방문을 유도하는 채널이 선택적으로 약화.",
            "<span style='color:#888'>※ 단, 본 데이터엔 <b>DAU 실측이 없음</b>. '도달↔DAU'는 정합적 <b>가설</b>이며 인과 확증엔 DAU 조인 분석이 별도 필요(사이드바에서 DAU 업로드 시 갱신).</span>",
        ]
    insight(prob, "warn", cap="⚠️ 핵심 문제 (Problem)")

    if has_dau:
        sowhat = [
            "범용 스티키니스(출첵·데일리딜)는 commoditized·콘텐츠는 소진 → 차별화 레버는 <b>정밀 재참여(정밀도 × 도달)</b>.",
            "<b>정밀도</b>: 행동데이터 D-1(하루 전)→실시간으로 신선화해 당일 행동 기반으로 발송. <b>도달</b>: 앱 푸시 도달율이 그 전달 용량(문자·광고 대비 저비용).",
            "지금 DAU는 <b>광고(유료)로 방어</b> 중 → owned 정밀 푸시로 전환해야 지속가능. 총 DAU 대신 <b>직접(자발) DAU·유료 의존도</b>를 건강 지표로.",
        ]
    else:
        sowhat = [
            "마케팅 레버가 '발송량·동의 확보'에서 '<b>앱 보유·재설치</b>'로 이동해야 함 — 동의 확보형 캠페인은 한계.",
            "수신거부 방어보다 <b>앱 미보유/삭제 풀(훨씬 큰 누수)의 재활성화</b>가 도달 개선 ROI가 큼.",
            "신규 획득 KPI와 별개로 '<b>앱 설치·유지 전환</b>'을 독립 지표로 관리해야 도달 자산이 축적됨.",
        ]
    insight(sowhat, cap="💡 인사이트 · 시사점 (So-What)")

    insight([
        "<b>① 정밀도 — 행동 발생 시점 기반 CRM(D-1 → 실시간)</b> — 당일 조회·찜·장바구니 행동을 당일 발송에 반영. "
        "현재는 하루 전 기준이라 관심 시점을 놓침. 이 latency 개선이 정밀 타겟팅의 전제.",
        f"<b>② 도달 — 앱 푸시 도달율 확대</b> — 정밀 발송의 전달 용량. 미도달 <b>{fnum(vip['unreach'])}명</b> 재설치 + 활성이나 도달 불가하던 층 포함. "
        "owned 채널이라 문자·광고 대비 저비용.",
        "<b>③ 유료 의존 축소 — 광고/문자로 산 DAU를 정밀 푸시로 전환</b> — 지속가능성·비용 개선. "
        "<b>직접(자발) DAU·유료 의존도</b>를 건강 지표로 모니터링.",
        "<b>④ 파일럿·검증</b> — 행동 트리거(D-1→실시간) 개선의 DAU 효과는 미검증 → 소규모 파일럿으로 실측 후 확대.",
        "<span style='color:#888'>참고(업계 일반 전술, 특정 사례 아님): 실시간 행동 트리거 푸시(재입고·가격인하·찜), 앱 미설치엔 웹푸시·카카오 대체 도달, "
        "딥링크로 설치 직후 이탈 방지 — 커머스 리인게이지먼트 표준 패턴.</span>",
    ], "ok", cap="✅ 권고 액션 (Action)")

    # ── 액션 임팩트 시뮬레이션 — 재설치 전환 → DAU 리프트 ──
    st.markdown('<div style="font-weight:700;font-size:15px;margin:16px 0 4px">🎛 시뮬레이션 ② '
                '— 미보유 재설치(도달) 레버 개선 시</div>', unsafe_allow_html=True)
    st.caption(f"미도달(앱 미보유/삭제) VIP {fnum(vip['unreach'])}명 중 일부가 재설치해 앱 보유로 전환된다고 가정합니다.")

    sc1, sc2 = st.columns(2)
    with sc1:
        conv = st.slider("재설치 전환율 (%)", 0.5, 10.0, 1.0, 0.5,
                         help="미도달 스톡 중 캠페인으로 앱을 다시 설치·유지하는 비율")
    with sc2:
        stick_default = float(f"{dau_sum['stick1']:.0f}") if "stick1" in dau_sum else 29.0
        stick = st.slider("재설치자 스티키니스 가정 — DAU/MAU (%)", 5.0, 50.0,
                          min(stick_default, 50.0) / 2, 1.0,
                          help="한 번 이탈했던 층이라 VIP 평균(약 "
                               f"{stick_default:.0f}%)보다 낮게 잡는 게 보수적입니다")

    reinstalled = vip["unreach"] * conv / 100
    dau_lift = reinstalled * stick / 100
    base_dau = dau_sum.get("dau_now", 0)
    yearly_loss = None
    if "dau_yoy" in dau_sum and base_dau:
        prev_dau = base_dau / (1 + dau_sum["dau_yoy"] / 100)
        yearly_loss = prev_dau - base_dau   # 전년 대비 감소 절대량(양수=감소)

    new_dau = base_dau + dau_lift if base_dau else None
    yoy_now = dau_sum.get("dau_yoy")            # 현 전년비 (%)
    yoy_after = None
    if yoy_now is not None and base_dau:
        prev_dau_base = base_dau / (1 + yoy_now / 100)
        yoy_after = (new_dau / prev_dau_base - 1) * 100

    m1, m2, m3, m4 = st.columns(4)
    with m1: metric_card("재설치(앱 보유 전환)", fnum(reinstalled), f"미도달 {fnum(vip['unreach'])}명 × {conv:.1f}%")
    with m2:
        if new_dau:
            metric_card("VIP DAU 변화", f"{fnum(base_dau)} → {fnum(new_dau)}",
                        f"+{fnum(dau_lift)}명 (+{dau_lift/base_dau*100:.1f}%)")
        else:
            metric_card("예상 DAU 리프트", f"+{fnum(dau_lift)}", f"스티키니스 {stick:.0f}% 가정")
    with m3:
        if yoy_after is not None:
            metric_card("전년비 변화", f"{yoy_now:+.1f}% → {yoy_after:+.1f}%",
                        "역신장 폭 축소" if yoy_after < 0 else "신장 전환")
        else:
            metric_card("전년비 변화", "—", "DAU 데이터 업로드 시 표시")
    with m4:
        if yearly_loss and yearly_loss > 0:
            metric_card("연간 감소 상쇄", f"{dau_lift/yearly_loss*100:.0f}%",
                        f"전년비 감소 {fnum(yearly_loss)}명 대비")
        else:
            metric_card("연간 감소 상쇄", "—", "DAU 데이터 업로드 시 표시")

    # 역신장 0%(전년 수준 회복)에 필요한 전환율 역산
    breakeven_txt = ""
    if yoy_now is not None and yoy_now < 0 and base_dau and vip["unreach"] and stick > 0:
        need_lift = prev_dau_base - base_dau
        need_conv = need_lift / (vip["unreach"] * stick / 100) * 100
        breakeven_txt = (f"<b>역신장 0%(전년 수준 회복)</b>까지 필요한 리프트는 +{fnum(need_lift)} — "
                         f"스티키니스 {stick:.0f}% 가정 시 재설치 전환율 <b>{need_conv:.1f}%</b>가 손익분기점.")

    insight([
        (f"전환 <b>{conv:.1f}%</b>·스티키니스 <b>{stick:.0f}%</b> 가정 시 VIP DAU <b>{fnum(base_dau)} → {fnum(new_dau)}</b>"
         f"(+{fnum(dau_lift)}), 전년비 <b>{yoy_now:+.1f}% → {yoy_after:+.1f}%</b>."
         if yoy_after is not None else
         f"전환 <b>{conv:.1f}%</b>·스티키니스 <b>{stick:.0f}%</b> 가정 시 DAU <b>+{fnum(dau_lift)}</b>."),
        breakeven_txt,
        "재설치자는 <b>푸시 타겟팅가능에도 편입</b> → 이후 푸시 발송으로 방문 빈도 제고 여지(선순환).",
        "<span style='color:#888'>※ 가정 기반 추정 — 재설치자의 실제 빈도·리텐션은 캠페인 후 실측으로 검증 필요.</span>",
    ])
else:
    st.info("VIP 데이터가 있어야 결론(인사이트·액션) 섹션이 생성됩니다.")

# ════════════════════════════════════════════════════════════
# 5. 상세 테이블
# ════════════════════════════════════════════════════════════
section("상세 데이터", "등급별 종합 (도달=최근일, 거부/순증감=기간 합계)", anchor="sec-table")
det = fw_last.groupby("grade").agg(act=("act", "sum"), act_push=("act_push", "sum"),
                                   tot_push=("tot_push", "sum"), unreach=("unreach_push", "sum")).reindex(grade_order_sel)
det["reach"] = np.where(det["act_push"] > 0, det["tot_push"] / det["act_push"] * 100, 0).round(1)
out_g = fl.groupby("grade").agg(out=("out", "sum"), new=("new", "sum")).reindex(grade_order_sel)
det = det.join(out_g)
det["net"] = det["new"] - det["out"]
det.insert(0, "group", [GRADE2GROUP.get(g, "") for g in det.index])
det = det.reset_index().rename(columns={
    "grade": "등급", "group": "그룹", "act": "유효회원", "act_push": "수신동의_PUSH",
    "tot_push": "타겟팅가능_PUSH", "unreach": "앱미보유/삭제", "reach": "도달률(%)",
    "out": "기존이탈", "new": "신규추가", "net": "증감"})
det = det[["그룹", "등급", "유효회원", "수신동의_PUSH", "타겟팅가능_PUSH", "앱미보유/삭제",
           "도달률(%)", "신규추가", "기존이탈", "증감"]]
numfmt = {c: st.column_config.NumberColumn(format="localized")
          for c in ["유효회원", "수신동의_PUSH", "타겟팅가능_PUSH", "앱미보유/삭제", "신규추가", "기존이탈", "증감"]}
st.dataframe(det, use_container_width=True, hide_index=True, column_config=numfmt)
st.download_button("⬇️ 집계 CSV 다운로드", det.to_csv(index=False).encode("utf-8-sig"),
                   "vip_reach_summary.csv", "text/csv")

st.caption("ⓘ 수신동의율 = ACT_PUSH/ACT_MEM · 타겟팅가능율(도달률) = TOT_PUSH/ACT_PUSH · "
           "앱 미보유/삭제 = ACT_PUSH−TOT_PUSH · 증감 = 신규추가−기존이탈 · VIP=SP·PT·GD·SV·BK / 일반=PP·RD")

# ── 스크롤스파이: 현재 보는 섹션을 사이드바 메뉴에서 강조 (베스트-에포트 JS) ──
components.html("""
<script>
const doc = window.parent.document;
function getScroller(){
  const c = [doc.scrollingElement, doc.querySelector('section.main'),
             doc.querySelector('[data-testid="stMain"]'),
             doc.querySelector('[data-testid="stAppViewContainer"]'),
             doc.documentElement, doc.body];
  for(const e of c){ if(e && e.scrollHeight > e.clientHeight + 5) return e; }
  return doc.scrollingElement || doc.documentElement;
}
function spy(){
  const titles = Array.from(doc.querySelectorAll('.section-title')).filter(t => t.id);
  if(!titles.length) return;
  let active = titles[0].id;
  for(const t of titles){ if(t.getBoundingClientRect().top <= 140) active = t.id; }
  const se = getScroller();
  if(se && se.scrollTop + se.clientHeight >= se.scrollHeight - 8){ active = titles[titles.length-1].id; }
  doc.querySelectorAll('a.navlink').forEach(a=>{
    const on = a.getAttribute('href') === '#'+active;
    a.style.background = on ? '#d6e4f7' : '#f2f5fa';
    a.style.color = on ? '#163E78' : '#2E68B0';
  });
}
window.parent.addEventListener('scroll', spy, true);
setInterval(spy, 400); setTimeout(spy, 300);
</script>
""", height=0)
