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
    .navgroup { font-size: 11px; font-weight: 700; color: #8a94a6; letter-spacing: .03em;
        margin: 12px 0 4px; text-transform: none; }
    a.navlink.navsub { margin-left: 8px; font-size: 13px; padding: 6px 10px; }
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
    ("🎯 종합 요약", [
        ("sec-core", "핵심 진단"),
    ]),
    ("📊 현황 진단", [
        ("sec-group", "그룹 비교 (VIP vs 일반)"),
        ("sec-reach", "앱푸시 도달 진단"),
        ("sec-optout", "수신거부 분석"),
        ("sec-within", "그룹 내 등급별"),
        ("sec-trend", "장기 추세 (전체)"),
    ]),
    ("🧭 결론", [
        ("sec-action", "인사이트 · 시사점 · 액션"),
    ]),
    ("📁 부록", [
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
    fig.update_yaxes(tickformat=",")   # 축 숫자 k/M 금지 → #,##0(천단위 콤마)
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
    up = st.file_uploader("데이터 업로드 (.xls / .xlsx)", type=["xls", "xlsx"])
    st.caption("STD_DD, GRADE_CD, ACT/TOT/NEW/OUT_채널_MEM 컬럼 포함 export 파일")

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
        _nav += f'<div class="navgroup">{_glabel}</div>'
        _nav += "".join(f'<a href="#{a}" class="navlink navsub">{lbl}</a>' for a, lbl in _items)
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
# 0. 핵심 진단
# ════════════════════════════════════════════════════════════
section("핵심 진단 — VIP · 앱푸시(DAU 채널) 기준",
        f"기간 {d0} ~ {d1} ({n_days}일) · 도달률=최근일({last_day.date()}) 스냅샷 · 모든 수치 VIP 전용",
        anchor="sec-core")

vip = group_snapshot("VIP", fw_d_last, fl_d, fw_d)   # 등급 필터와 무관하게 항상 VIP 전체
if vip["act_push"]:
    share = vip["unreach"] / vip["act_push"] * 100
    push_net = vip["chnet"].get("PUSH", 0)
    push_out = vip["chout"].get("PUSH", 0)
    kind = "warn" if (vip["reach"] < 50 or push_net < 0) else ""
    insight([
        f"수신동의는 <b>{vip['consent']:.1f}%</b>로 이미 충분 — 병목은 동의가 아니라 <b>앱 보유</b>(타겟팅가능 {vip['reach']:.1f}%). 동의 확보형 캠페인은 효과 한계.",
        f"<b>{fnum(vip['unreach'])}명(동의자의 {share:.0f}%)</b>이 앱 미보유/삭제로 푸시 도달 불가 → 이 풀의 <b>재설치 전환</b>이 VIP DAU 회복의 최대 레버.",
        f"PUSH만 {'순감' if push_net < 0 else '정체'}({fsigned(push_net)})이고 SMS({fsigned(vip['chnet'].get('SMS',0))})·EMAIL({fsigned(vip['chnet'].get('EMAIL',0))})은 순증 → 앱 채널만 약화. 푸시 못 닿는 VIP엔 <b>알림톡/카카오 대체 도달</b> 병행.",
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
# 1. 그룹 비교 (VIP vs 일반)
# ════════════════════════════════════════════════════════════
section("그룹 비교 — VIP vs 일반", "VIP 중심, 일반은 대조용 · 등급 필터와 무관하게 항상 표시", anchor="sec-group")
gcols = st.columns(2)
gsnap = {}
for col, grp in zip(gcols, ["VIP", "일반"]):
    gs = group_snapshot(grp, fw_d_last, fl_d, fw_d)
    gsnap[grp] = gs
    share = (gs["unreach"] / gs["act_push"] * 100) if gs["act_push"] else 0
    cn = gs["chnet"]
    with col:
        st.markdown(
            f'<div class="metric-card" style="border-left-color:{GROUP_COLOR[grp]}">'
            f'<div class="metric-value" style="font-size:19px">{grp} '
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

if "VIP" in gsnap and "일반" in gsnap:
    v, gen = gsnap["VIP"], gsnap["일반"]
    vp, gpn = v["chnet"].get("PUSH", 0), gen["chnet"].get("PUSH", 0)
    insight([
        f"VIP 도달률({v['reach']:.1f}%)이 일반({gen['reach']:.1f}%)보다 높지만 절대 미보유/삭제가 <b>{fnum(v['unreach'])}명</b> → 규모 자체가 커서 VIP 우선 공략의 실익이 큼.",
        (f"VIP PUSH는 <b>{fsigned(vp)}</b>로 {'순감' if vp < 0 else '정체'}인데 일반은 {fsigned(gpn)} → 같은 푸시인데 VIP만 약화, VIP 전용 리텐션 대응 필요."
         if vp < gpn or vp < 0 else
         f"VIP·일반 PUSH 증감(VIP {fsigned(vp)} / 일반 {fsigned(gpn)}) 모두 점검 — 푸시 도달 모수 방어가 공통 과제."),
    ])

# ════════════════════════════════════════════════════════════
# 2. 앱푸시 도달률 & 앱 미보유/삭제 (DAU 핵심)
# ════════════════════════════════════════════════════════════
section("앱푸시 도달 진단 — 등급별", "최근일 스냅샷 · 막대=타겟팅가능 vs 미보유/삭제, 라인=도달률",
        anchor="sec-reach")
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

worst = gp["reach"].idxmin()
insight([
    f"도달률 최저 <b>{worst}</b>({gp.loc[worst,'reach']:.1f}%)에 미보유/삭제 <b>{fnum(gp.loc[worst,'unreach'])}명</b> 집중 → 재설치 캠페인 <b>1순위 타겟</b>.",
    "막대 빨강(미보유)이 큰 등급은 발송해도 안 닿음 → 발송 타겟 선정 시 '<b>앱 보유</b>' 필터를 적용해 도달 효율 확보.",
])

# VIP 앱 미보유/삭제 일별 추세 (이탈과 별개 누수) — 등급 필터 무관 VIP 전체
vip_daily = (fw_d[fw_d["group"] == "VIP"].groupby("date")
             .agg(unreach=("unreach_push", "sum"), ap=("act_push", "sum")).reset_index())
if len(vip_daily) > 1:
    vip_daily["pct"] = np.where(vip_daily["ap"] > 0, vip_daily["unreach"] / vip_daily["ap"] * 100, 0)
    figv = go.Figure()
    figv.add_bar(x=vip_daily["date"], y=vip_daily["unreach"], name="앱 미보유/삭제(명)",
                 marker_color="#e7c3c3", marker_line_width=0, opacity=0.75)
    figv.add_scatter(x=vip_daily["date"], y=vip_daily["pct"], name="수신동의 대비(%)", yaxis="y2",
                     mode="lines+markers", line=dict(color="#8C3A3A", width=2.5))
    figv.update_layout(height=300, margin=dict(t=10, b=10), hovermode="x unified", legend_title_text="",
                       yaxis=dict(title="앱 미보유/삭제(명)"),
                       yaxis2=dict(title="수신동의 대비(%)", overlaying="y", side="right", showgrid=False))
    plot(figv, "VIP 앱 미보유/삭제 일별 추세")
    st.caption("앱 미보유/삭제 = 수신동의 − 타겟팅가능(PUSH). 수신거부(이탈)와 다른 누수 — 동의는 유지하지만 "
               "앱이 없어 못 닿는 모수. (VIP 등급 데이터는 6/15부터라 추세가 짧음)")

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
                       xaxis=dict(categoryorder="array", categoryarray=grade_order_sel))
    plot(figo, "등급별 수신거부율 (전채널)")
    st.caption("※ 모수가 작은 상위 등급(SP·PT)은 이탈 몇 건만으로 율이 크게 튑니다. "
               "막대 위 절대 이탈 건수를 함께 보세요 — 실제 이탈 물량은 RD·BK·PP가 큽니다.")

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
    z, fmt, cs = (po / pt.replace(0, np.nan) * 100).fillna(0), ".3f", "OrRd"
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
    f"DAU와 직결되는 건 <b>PUSH 증감({fsigned(push_net2)})</b> — SMS/EMAIL 순증은 앱 방문 기여가 적어 수신거부 방어 우선순위는 PUSH.",
    "등급별 '수신거부율'은 모수 작은 상위 등급에서 튀는 노이즈 → 방어 대상은 율이 아니라 <b>절대 이탈 건수(RD·BK·PP)</b> 기준으로 선정.",
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
    st.caption(f"※ 풀 리카운트 등 비정상적으로 튀는 날은 정확도를 위해 자동 제외(롤링 중앙값 대비 과대 편차). {excl_txt}")

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
        plot(figg, "앱 미보유/삭제 추세 (전체) — 동의했지만 앱이 없어 못 닿는 모수")

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

    insight([
        f"VIP 수신동의는 <b>{vip['consent']:.1f}%</b>로 확보됐으나 실제 앱푸시 <b>타겟팅가능은 {vip['reach']:.1f}%</b>"
        f"({fnum(vip['tot_push'])}/{fnum(vip['act_push'])})에 불과.",
        f"동의자 중 <b>{fnum(vip['unreach'])}명({a_share:.0f}%)</b>이 앱 미보유/삭제 상태, "
        f"VIP PUSH 증감 {fsigned(a_pnet)}(순{'감' if a_pnet < 0 else '증'}).",
        (f"{lt_reach_txt} · {lt_gap_txt}." if lt_reach_txt else ""),
    ], cap="📌 현황 진단 (As-Is)")

    insight([
        "병목은 '수신 동의'가 아니라 '<b>앱 보유</b>' — 동의는 98%+지만 실제 도달은 1/3 수준.",
        "<b>앱 미보유/삭제 풀이 구조적으로 확대</b> → 발송량을 늘려도 도달 모수가 안 늘어 DAU 역신장으로 직결.",
        "SMS/EMAIL은 순증하는데 <b>PUSH(앱 채널)만 순감</b> → DAU를 끄는 채널이 선택적으로 약화.",
    ], "warn", cap="⚠️ 핵심 문제 (Problem)")

    insight([
        "마케팅 레버가 '발송량·동의 확보'에서 '<b>앱 보유·재설치</b>'로 이동해야 함 — 동의 확보형 캠페인은 한계.",
        "수신거부 방어보다 <b>앱 미보유/삭제 풀(훨씬 큰 누수)의 재활성화</b>가 도달·DAU 개선 ROI가 큼.",
        "신규 획득 KPI와 별개로 '<b>앱 설치·유지 전환</b>'을 독립 지표로 관리해야 도달 자산이 축적됨.",
    ], cap="💡 인사이트 · 시사점 (So-What)")

    insight([
        f"<b>① 재설치 캠페인</b> — 앱 미보유 VIP <b>{fnum(vip['unreach'])}명</b> 대상 인센티브+딥링크 재설치 유도(도달률 최저 등급 우선).",
        "<b>② 발송 최적화</b> — 발송 타겟에 '앱 보유' 필터 적용, 미보유 VIP엔 알림톡/카카오 등 대체 도달 병행.",
        "<b>③ 구조 개선</b> — 신규 유입 → 앱 설치 전환율 제고(온보딩·설치 유도 강화).",
        "<b>④ 지표화·모니터링</b> — 앱 보유 전환율·푸시 도달률을 주간 KPI로 추적, 하락 시 조기 대응.",
    ], "ok", cap="✅ 권고 액션 (Action)")
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
