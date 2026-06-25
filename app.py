import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit.components.v1 as components

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

# ── 커스텀 CSS ─────────────────────────────────────────────
st.markdown("""
<style>
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
    .insight {
        background: #eef4fb; border-left: 4px solid #4C72B0; border-radius: 8px;
        padding: 12px 16px; margin: 6px 0 14px 0; font-size: 14px; line-height: 1.6;
    }
    .insight.warn { background: #fdeeee; border-left-color: #C44E52; }
    .insight.ok   { background: #eef7f0; border-left-color: #55A868; }
    .insight b { color: #1a1a2e; }
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


# 사이드바 분석 메뉴 항목 (anchor, 라벨)
MENU = [
    ("sec-core", "🔑 핵심 진단"),
    ("sec-group", "👥 그룹 비교 (VIP vs 일반)"),
    ("sec-reach", "📲 앱푸시 도달 진단"),
    ("sec-optout", "🚫 수신거부 분석"),
    ("sec-within", "🏅 그룹 내 등급별 인사이트"),
    ("sec-table", "📋 상세 데이터"),
]


def insight(html, kind=""):
    st.markdown(f'<div class="insight {kind}">{html}</div>', unsafe_allow_html=True)


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
- **앱 미보유/삭제** = `ACT_PUSH_MEM − TOT_PUSH_MEM` (앱푸시 동의했지만 실제 발송 불가)
- **푸시 도달률** = `TOT_PUSH_MEM / ACT_PUSH_MEM` (앱 보유·푸시 발송 가능 비율)
- **수신거부율** = `OUT / TOT`  ·  **순증감** = `NEW − OUT`
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
    st.markdown("".join(f'<a href="#{a}" class="navlink">{lbl}</a>' for a, lbl in MENU),
                unsafe_allow_html=True)
    st.divider()
    st.caption("🔎 필터")
    dmin, dmax = W["date"].min().date(), W["date"].max().date()
    dr = st.date_input("기간", value=(dmin, dmax), min_value=dmin, max_value=dmax)
    d0, d1 = dr if isinstance(dr, tuple) and len(dr) == 2 else (dmin, dmax)
    sel_groups = st.multiselect("그룹", list(GROUPS.keys()), default=list(GROUPS.keys()))
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


def group_snapshot(grp, snap, period_long, period_wide):
    """그룹 단위 집계 (스냅샷=도달률, 기간합=거부/순증감)."""
    s = snap[snap["group"] == grp]
    pl = period_long[period_long["group"] == grp]
    pw = period_wide[period_wide["group"] == grp]
    act_push = s["act_push"].sum()
    tot_push = s["tot_push"].sum()
    return {
        "act": s["act"].sum(),
        "act_push": act_push,
        "tot_push": tot_push,
        "unreach": s["unreach_push"].sum(),
        "reach": (tot_push / act_push * 100) if act_push else 0,
        "out": pl["out"].sum(),
        "new": pl["new"].sum(),
        "net": pl["new"].sum() - pl["out"].sum(),
        "out_trend": trend_word(pw.groupby("date")["out_all"].sum())[0],
    }


# ════════════════════════════════════════════════════════════
# 0. 핵심 진단
# ════════════════════════════════════════════════════════════
section("핵심 진단", f"기간 {d0} ~ {d1} ({n_days}일) · 도달률은 최근일({last_day.date()}) 스냅샷 기준",
        anchor="sec-core")

if "VIP" in sel_groups:
    vip = group_snapshot("VIP", fw_last, fl, fw)
    share = (vip["unreach"] / vip["act_push"] * 100) if vip["act_push"] else 0
    kind = "warn" if vip["reach"] < 50 else ""
    insight(
        f"VIP 앱푸시 도달률은 <b>{vip['reach']:.1f}%</b> — 앱푸시 동의 {fnum(vip['act_push'])}명 중 "
        f"<b>{fnum(vip['unreach'])}명({share:.0f}%)</b>이 <b>앱 미보유/삭제</b>로 실제 발송 불가입니다. "
        f"기간 중 VIP 수신거부는 총 <b>{fnum(vip['out'])}건</b>(추세 {vip['out_trend']}), "
        f"순증감 <b>{fsigned(vip['net'])}</b>. "
        f"도달 가능 모수가 동의 모수의 1/3 수준이라, 발송량을 늘려도 DAU 기여 한계가 큽니다.",
        kind)

c1, c2, c3, c4, c5 = st.columns(5)
push_reach_all = (fw_last["tot_push"].sum() / fw_last["act_push"].sum() * 100) if fw_last["act_push"].sum() else 0
with c1: metric_card("선택분 앱푸시 도달률", f"{push_reach_all:.1f}%", "TOT_PUSH / ACT_PUSH")
with c2: metric_card("앱 미보유/삭제", fnum(fw_last["unreach_push"].sum()), "푸시 발송 불가 모수")
with c3: metric_card("앱푸시 발송 가능", fnum(fw_last["tot_push"].sum()), f"동의 {fnum(fw_last['act_push'].sum())}")
with c4: metric_card("기간 수신거부(전채널)", fnum(fl["out"].sum()), f"일평균 {fnum(fl['out'].sum()/n_days)}")
net_all = fl["new"].sum() - fl["out"].sum()
with c5: metric_card("순증감(신규−거부)", fsigned(net_all),
                     "구독자 순증가" if net_all >= 0 else "구독자 순감소")

# ════════════════════════════════════════════════════════════
# 1. 그룹 비교 (VIP vs 일반)
# ════════════════════════════════════════════════════════════
section("그룹 비교 — VIP vs 일반", "도달률=최근일 스냅샷 · 거부/순증감=기간 합계", anchor="sec-group")
gcols = st.columns(len(sel_groups))
for col, grp in zip(gcols, sel_groups):
    gs = group_snapshot(grp, fw_last, fl, fw)
    share = (gs["unreach"] / gs["act_push"] * 100) if gs["act_push"] else 0
    with col:
        st.markdown(
            f'<div class="metric-card" style="border-left-color:{GROUP_COLOR[grp]}">'
            f'<div class="metric-value" style="font-size:19px">{grp} '
            f'<span style="font-size:13px;color:#888">({", ".join(GROUPS[grp])})</span></div>'
            f'<div class="metric-sub">앱푸시 도달률 <b>{gs["reach"]:.1f}%</b> · '
            f'미보유/삭제 <b>{fnum(gs["unreach"])}</b> ({share:.0f}%)</div>'
            f'<div class="metric-sub">유효회원 {fnum(gs["act"])} · 발송가능 {fnum(gs["tot_push"])}</div>'
            f'<div class="metric-sub">기간 수신거부 <b>{fnum(gs["out"])}</b> (추세 {gs["out_trend"]}) · '
            f'순증감 <b>{fsigned(gs["net"])}</b></div>'
            f'</div>', unsafe_allow_html=True)

# 그룹별 도달률 일별 추이
grp_daily = (fw.groupby(["date", "group"]).agg(tp=("tot_push", "sum"), ap=("act_push", "sum")).reset_index())
grp_daily["reach"] = np.where(grp_daily["ap"] > 0, grp_daily["tp"] / grp_daily["ap"] * 100, 0)
figr = px.line(grp_daily, x="date", y="reach", color="group", markers=True,
               color_discrete_map=GROUP_COLOR, labels={"reach": "앱푸시 도달률(%)", "date": "일자", "group": "그룹"})
figr.update_layout(height=320, margin=dict(t=20, b=10), hovermode="x unified", legend_title_text="")
st.plotly_chart(figr, use_container_width=True)

# ════════════════════════════════════════════════════════════
# 2. 앱푸시 도달률 & 앱 미보유/삭제 (DAU 핵심)
# ════════════════════════════════════════════════════════════
section("앱푸시 도달 진단 — 등급별", "최근일 스냅샷 · 막대=발송가능 vs 미보유/삭제, 라인=도달률",
        anchor="sec-reach")
gp = (fw_last.groupby("grade").agg(act_push=("act_push", "sum"), tot_push=("tot_push", "sum"),
                                   unreach=("unreach_push", "sum")).reindex(grade_order_sel))
gp["reach"] = np.where(gp["act_push"] > 0, gp["tot_push"] / gp["act_push"] * 100, 0)
fig = go.Figure()
fig.add_bar(x=gp.index, y=gp["tot_push"], name="발송가능(앱보유)", marker_color="#55A868")
fig.add_bar(x=gp.index, y=gp["unreach"], name="앱 미보유/삭제", marker_color="#C44E52")
fig.add_scatter(x=gp.index, y=gp["reach"], name="도달률(%)", yaxis="y2",
                mode="lines+markers+text", text=[f"{v:.0f}%" for v in gp["reach"]],
                textposition="top center", line=dict(color="#1a1a2e", width=2))
fig.update_layout(barmode="stack", height=400, margin=dict(t=20, b=10),
                  yaxis=dict(title="앱푸시 동의 모수"), legend_title_text="",
                  yaxis2=dict(title="도달률(%)", overlaying="y", side="right", range=[0, 100], showgrid=False),
                  xaxis=dict(categoryorder="array", categoryarray=grade_order_sel))
st.plotly_chart(fig, use_container_width=True)

worst = gp["reach"].idxmin()
insight(
    f"도달률이 가장 낮은 등급은 <b>{worst}</b> (<b>{gp.loc[worst,'reach']:.1f}%</b>, 미보유/삭제 "
    f"{fnum(gp.loc[worst,'unreach'])}명). 등급 라벨 순서는 상위(SP)→하위(RD)이며, 막대의 빨강 영역이 클수록 "
    f"앱푸시로 닿지 못하는 모수가 큽니다 — DAU 회복을 위해 우선 공략할 구간입니다.")

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
    figc.update_layout(height=320, margin=dict(t=20, b=10), hovermode="x unified",
                       legend_title_text="", title="일별 채널별 수신거부 추이")
    if log_y:
        figc.update_yaxes(type="log")
    st.plotly_chart(figc, use_container_width=True)
with cc2:
    # 수신거부율 = 기간 거부 합계 / 최근일 수신자수(선택 채널 합) — 채널 간 max 혼용 제거
    lc = fl["date"].max()
    last_tot = fl[fl["date"] == lc].groupby("grade")["tot"].sum()
    grade_out = fl.groupby("grade")["out"].sum().to_frame("out")
    grade_out["tot"] = last_tot
    grade_out = grade_out.reindex(grade_order_sel)
    grade_out["rate"] = np.where(grade_out["tot"] > 0, grade_out["out"] / grade_out["tot"] * 100, 0)
    figo = px.bar(grade_out.reset_index(), x="grade", y="rate", color="rate", color_continuous_scale="OrRd",
                  labels={"rate": "수신거부율(%)", "grade": "등급"}, title="등급별 수신거부율")
    figo.update_layout(height=350, margin=dict(t=20, b=10), coloraxis_showscale=False,
                       xaxis=dict(categoryorder="array", categoryarray=grade_order_sel))
    st.plotly_chart(figo, use_container_width=True)

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
st.plotly_chart(figh, use_container_width=True)

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
    msg = (f"{grp} 내 도달률 최저는 <b>{lo_reach}</b>({snap.loc[lo_reach,'reach']:.1f}%), "
           f"수신거부율 최고는 <b>{hi_out}</b>({outp.loc[hi_out,'rate']:.3f}%).")
    if neg_net:
        msg += f" 기간 중 <b>구독자 순감소</b> 등급: {', '.join(neg_net)} — 신규수신보다 이탈이 많습니다."
    else:
        msg += " 모든 등급이 기간 중 구독자 순증가 상태입니다."
    insight(msg, "warn" if (snap["reach"].min() < 40 or neg_net) else "")

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
    "grade": "등급", "group": "그룹", "act": "유효회원수", "act_push": "앱푸시동의",
    "tot_push": "발송가능", "unreach": "앱미보유/삭제", "reach": "도달률(%)",
    "out": "수신거부", "new": "신규수신", "net": "순증감"})
det = det[["그룹", "등급", "유효회원수", "앱푸시동의", "발송가능", "앱미보유/삭제",
           "도달률(%)", "신규수신", "수신거부", "순증감"]]
numfmt = {c: st.column_config.NumberColumn(format="localized")
          for c in ["유효회원수", "앱푸시동의", "발송가능", "앱미보유/삭제", "신규수신", "수신거부", "순증감"]}
st.dataframe(det, use_container_width=True, hide_index=True, column_config=numfmt)
st.download_button("⬇️ 집계 CSV 다운로드", det.to_csv(index=False).encode("utf-8-sig"),
                   "vip_reach_summary.csv", "text/csv")

st.caption("ⓘ 앱 미보유/삭제 = ACT_PUSH_MEM − TOT_PUSH_MEM · 도달률 = TOT_PUSH/ACT_PUSH · "
           "수신거부율 = OUT/TOT · 순증감 = NEW − OUT · VIP=SP·PT·GD·SV·BK / 일반=PP·RD")

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
