import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ── 기본 설정 ──────────────────────────────────────────────
st.set_page_config(
    page_title="등급별·채널별 수신거부 대시보드",
    page_icon="📭",
    layout="wide",
    initial_sidebar_state="expanded",
)

CHANNELS = ["PUSH", "SMS", "EMAIL"]
CH_COLOR = {"PUSH": "#4C72B0", "SMS": "#DD8452", "EMAIL": "#55A868"}

# ── 커스텀 CSS ─────────────────────────────────────────────
st.markdown("""
<style>
    .metric-card {
        background: #f8f9fa; border-radius: 10px; padding: 16px 20px;
        border-left: 4px solid #4C72B0; margin-bottom: 10px;
    }
    .metric-label { font-size: 13px; color: #666; margin-bottom: 4px; }
    .metric-value { font-size: 26px; font-weight: 700; color: #1a1a2e; }
    .metric-sub  { font-size: 12px; color: #888; margin-top: 2px; }
    .section-title {
        font-size: 18px; font-weight: 700; color: #1a1a2e;
        margin: 26px 0 12px 0; padding-bottom: 6px;
        border-bottom: 2px solid #e9ecef;
    }
    .hint { font-size: 12px; color: #999; margin: -4px 0 10px 0; }
</style>
""", unsafe_allow_html=True)


# ── 데이터 로드 / 정규화 ────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(file):
    """업로드한 엑셀(.xls/.xlsx)을 long-format으로 정규화."""
    # 첫 번째 시트가 데이터(시트명이 깨져 있을 수 있어 인덱스로 접근), SQL 시트는 무시
    raw = pd.read_excel(file, sheet_name=0, header=0)
    raw.columns = [str(c).strip() for c in raw.columns]
    # 빈/Unnamed 컬럼 제거
    raw = raw.loc[:, [c for c in raw.columns if c and not c.startswith("Unnamed")]]

    needed = {"STD_DD", "GRADE_CD"}
    if not needed.issubset(raw.columns):
        raise ValueError(f"필수 컬럼 누락: {needed - set(raw.columns)}")

    raw = raw.dropna(subset=["STD_DD", "GRADE_CD"])
    raw["STD_DD"] = pd.to_datetime(raw["STD_DD"].astype(str).str.strip(), format="%Y%m%d", errors="coerce")
    raw = raw.dropna(subset=["STD_DD"])
    raw["GRADE_CD"] = raw["GRADE_CD"].astype(str).str.strip()

    num_cols = [c for c in raw.columns if c not in ("STD_DD", "NEW_GBN", "GRADE_CD")]
    for c in num_cols:
        raw[c] = pd.to_numeric(raw[c], errors="coerce").fillna(0)

    # long format: 한 행 = (날짜, 등급, 채널)
    records = []
    for _, r in raw.iterrows():
        act = r.get("ACT_MEM", 0)          # 전체 유효회원수
        act_push = r.get("ACT_PUSH_MEM", 0)
        for ch in CHANNELS:
            records.append({
                "date": r["STD_DD"],
                "grade": r["GRADE_CD"],
                "channel": ch,
                "act": act,                                  # 유효회원수(전체)
                "act_push": act_push,
                "tot": r.get(f"TOT_{ch}_MEM", 0),            # 수신자수
                "new": r.get(f"NEW_{ch}_MEM", 0),            # 신규수신
                "out": r.get(f"OUT_{ch}_MEM", 0),            # 수신거부(이탈)
            })
    df = pd.DataFrame(records)
    df["net"] = df["new"] - df["out"]                        # 순증감
    df["out_rate"] = np.where(df["tot"] > 0, df["out"] / df["tot"] * 100, 0)   # 수신거부율(%)
    df["reach"] = np.where(df["act"] > 0, df["tot"] / df["act"] * 100, 0)      # 도달률(%)
    return df


def fnum(x):
    return f"{int(round(x)):,}"


def metric_card(label, value, sub=""):
    st.markdown(
        f'<div class="metric-card"><div class="metric-label">{label}</div>'
        f'<div class="metric-value">{value}</div>'
        f'<div class="metric-sub">{sub}</div></div>',
        unsafe_allow_html=True,
    )


def section(title, hint=""):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)
    if hint:
        st.markdown(f'<div class="hint">{hint}</div>', unsafe_allow_html=True)


# ── 헤더 ───────────────────────────────────────────────────
st.title("📭 등급별·채널별 수신거부 대시보드")
st.caption("PUSH / SMS / EMAIL 채널의 수신거부(이탈)·신규수신·순증감을 등급별로 분석합니다.")

# ── 사이드바: 업로드 & 필터 ─────────────────────────────────
with st.sidebar:
    st.header("⚙️ 설정")
    up = st.file_uploader("발송/수신 데이터 업로드 (.xls / .xlsx)", type=["xls", "xlsx"])
    st.caption("STD_DD, GRADE_CD, ACT/TOT/NEW/OUT_채널_MEM 컬럼을 포함한 export 파일")

if up is None:
    st.info("👈 사이드바에서 데이터 파일을 업로드하면 대시보드가 표시됩니다.")
    with st.expander("📋 예상 컬럼 구조"):
        st.markdown("""
| 컬럼 | 의미 |
|---|---|
| `STD_DD` | 기준일자 (YYYYMMDD) |
| `GRADE_CD` | 등급코드 (BK/GD/PP/PT/RD/SP/SV/TOTAL) |
| `ACT_MEM`, `ACT_PUSH_MEM` | 유효회원수 |
| `TOT_{PUSH/SMS/EMAIL}_MEM` | 수신자수 |
| `NEW_{PUSH/SMS/EMAIL}_MEM` | 신규수신 |
| `OUT_{PUSH/SMS/EMAIL}_MEM` | 수신거부(이탈) |
""")
    st.stop()

try:
    df = load_data(up)
except Exception as e:
    st.error(f"데이터를 읽는 중 오류가 발생했습니다: {e}")
    st.stop()

# 합계행(TOTAL)과 등급행 분리
grades_all = sorted(df["grade"].unique().tolist())
grade_codes = [g for g in grades_all if g != "TOTAL"]
df_grade = df[df["grade"] != "TOTAL"].copy()       # 등급별 분석용
has_total_row = "TOTAL" in grades_all

# ── 사이드바 필터 ───────────────────────────────────────────
with st.sidebar:
    st.divider()
    dmin, dmax = df["date"].min().date(), df["date"].max().date()
    date_range = st.date_input("기간", value=(dmin, dmax), min_value=dmin, max_value=dmax)
    if isinstance(date_range, tuple) and len(date_range) == 2:
        d0, d1 = date_range
    else:
        d0, d1 = dmin, dmax

    sel_channels = st.multiselect("채널", CHANNELS, default=CHANNELS)
    # 등급 순서: 유효회원수 큰 순
    order = (df_grade.groupby("grade")["act"].max().sort_values(ascending=False).index.tolist())
    sel_grades = st.multiselect("등급", order, default=order)

mask = (
    (df["date"].dt.date >= d0) & (df["date"].dt.date <= d1) &
    (df["channel"].isin(sel_channels))
)
fdf = df[mask].copy()
fg = fdf[(fdf["grade"] != "TOTAL") & (fdf["grade"].isin(sel_grades))].copy()  # 등급 필터 적용

if fg.empty:
    st.warning("선택한 조건에 해당하는 데이터가 없습니다. 필터를 조정해 주세요.")
    st.stop()

n_days = fg["date"].nunique()
last_day = fg["date"].max()

# ── KPI ────────────────────────────────────────────────────
section("핵심 지표", f"기간: {d0} ~ {d1} ({n_days}일) · 채널: {', '.join(sel_channels)} · 등급별 합산")
tot_out = fg["out"].sum()
tot_new = fg["new"].sum()
tot_net = tot_new - tot_out
recv_last = fg[fg["date"] == last_day]["tot"].sum()
out_last = fg[fg["date"] == last_day]["out"].sum()
out_rate_period = (tot_out / fg["tot"].sum() * 100) if fg["tot"].sum() > 0 else 0
avg_daily_out = tot_out / n_days if n_days else 0

c1, c2, c3, c4, c5 = st.columns(5)
with c1: metric_card("총 수신거부", fnum(tot_out), f"일평균 {fnum(avg_daily_out)}건")
with c2: metric_card("총 신규수신", fnum(tot_new), f"기간 {n_days}일 합계")
with c3:
    sign = "▲" if tot_net >= 0 else "▼"
    metric_card("순증감 (신규−거부)", f"{sign} {fnum(abs(tot_net))}",
                "구독자 순증가" if tot_net >= 0 else "구독자 순감소")
with c4: metric_card("기간 수신거부율", f"{out_rate_period:.3f}%", "수신거부 / 수신자수")
with c5: metric_card(f"최근일({last_day.date()}) 거부", fnum(out_last), f"수신자 {fnum(recv_last)}")

# ── 채널별 요약 ────────────────────────────────────────────
section("채널별 요약", "선택 기간 합계 기준")
ch_sum = (fg.groupby("channel")
            .agg(out=("out", "sum"), new=("new", "sum"), tot=("tot", "sum"))
            .reindex([c for c in CHANNELS if c in sel_channels]))
ch_sum["net"] = ch_sum["new"] - ch_sum["out"]
ch_sum["out_rate"] = np.where(ch_sum["tot"] > 0, ch_sum["out"] / ch_sum["tot"] * 100, 0)

cols = st.columns(len(ch_sum))
for col, (ch, row) in zip(cols, ch_sum.iterrows()):
    with col:
        net = row["net"]; sign = "▲" if net >= 0 else "▼"
        st.markdown(
            f'<div class="metric-card" style="border-left-color:{CH_COLOR[ch]}">'
            f'<div class="metric-value" style="font-size:18px">{ch}</div>'
            f'<div class="metric-sub">수신거부 <b>{fnum(row["out"])}</b> · 신규 <b>{fnum(row["new"])}</b></div>'
            f'<div class="metric-sub">순증감 <b>{sign} {fnum(abs(net))}</b> · 거부율 <b>{row["out_rate"]:.3f}%</b></div>'
            f'<div class="metric-sub">수신자 {fnum(row["tot"])}</div>'
            f'</div>', unsafe_allow_html=True)

# ── 일별 추이 ──────────────────────────────────────────────
section("일별 수신거부 추이", "채널별 일자 수신거부(이탈) 건수")
daily_ch = (fg.groupby(["date", "channel"])["out"].sum().reset_index())
fig = px.line(daily_ch, x="date", y="out", color="channel",
              color_discrete_map=CH_COLOR, markers=True,
              labels={"out": "수신거부", "date": "일자", "channel": "채널"})
fig.update_layout(height=360, legend_title_text="", margin=dict(t=20, b=10), hovermode="x unified")
st.plotly_chart(fig, use_container_width=True)

# 신규 vs 거부 (순증감 막대)
section("신규수신 vs 수신거부 (순증감)", "막대=순증감, 양수면 구독자 증가 · 채널 합산")
daily_net = (fg.groupby("date").agg(new=("new", "sum"), out=("out", "sum")).reset_index())
daily_net["net"] = daily_net["new"] - daily_net["out"]
fig2 = go.Figure()
fig2.add_bar(x=daily_net["date"], y=daily_net["net"],
             marker_color=np.where(daily_net["net"] >= 0, "#55A868", "#C44E52"), name="순증감")
fig2.add_scatter(x=daily_net["date"], y=daily_net["new"], mode="lines+markers",
                 line=dict(color="#4C72B0", dash="dot"), name="신규수신")
fig2.add_scatter(x=daily_net["date"], y=daily_net["out"], mode="lines+markers",
                 line=dict(color="#DD8452", dash="dot"), name="수신거부")
fig2.update_layout(height=360, margin=dict(t=20, b=10), hovermode="x unified", legend_title_text="")
st.plotly_chart(fig2, use_container_width=True)

# ── 등급 × 채널 히트맵 ──────────────────────────────────────
section("등급 × 채널 매트릭스", "선택 기간 합계 — 지표를 바꿔서 비교하세요")
metric_choice = st.radio("지표 선택", ["수신거부 건수", "수신거부율(%)", "순증감"],
                         horizontal=True, label_visibility="collapsed")
pivot_out = fg.pivot_table(index="grade", columns="channel", values="out", aggfunc="sum", fill_value=0)
pivot_tot = fg.pivot_table(index="grade", columns="channel", values="tot", aggfunc="sum", fill_value=0)
pivot_new = fg.pivot_table(index="grade", columns="channel", values="new", aggfunc="sum", fill_value=0)
grade_order = [g for g in order if g in pivot_out.index]
ch_order = [c for c in CHANNELS if c in sel_channels]
pivot_out = pivot_out.reindex(index=grade_order, columns=ch_order)
pivot_tot = pivot_tot.reindex(index=grade_order, columns=ch_order)
pivot_new = pivot_new.reindex(index=grade_order, columns=ch_order)

if metric_choice == "수신거부 건수":
    z = pivot_out; fmt = ":,.0f"; cs = "Reds"
elif metric_choice == "수신거부율(%)":
    z = (pivot_out / pivot_tot.replace(0, np.nan) * 100).fillna(0); fmt = ":.3f"; cs = "OrRd"
else:
    z = (pivot_new - pivot_out); fmt = ":,.0f"; cs = "RdYlGn"

fig3 = px.imshow(z, text_auto=fmt.lstrip(":"), aspect="auto", color_continuous_scale=cs,
                 labels=dict(x="채널", y="등급", color=metric_choice))
fig3.update_layout(height=max(300, 40 * len(grade_order) + 80), margin=dict(t=20, b=10))
st.plotly_chart(fig3, use_container_width=True)

# ── 등급별 분석 ────────────────────────────────────────────
section("등급별 분석", "선택 기간 합계 — 수신거부·신규·순증감")
g_sum = (fg.groupby("grade").agg(out=("out", "sum"), new=("new", "sum"),
                                 tot=("tot", "sum"), act=("act", "max")).reset_index())
g_sum["net"] = g_sum["new"] - g_sum["out"]
g_sum["out_rate"] = np.where(g_sum["tot"] > 0, g_sum["out"] / g_sum["tot"] * 100, 0)
g_sum = g_sum.set_index("grade").reindex(grade_order).reset_index()

colA, colB = st.columns(2)
with colA:
    figg = go.Figure()
    figg.add_bar(x=g_sum["grade"], y=g_sum["new"], name="신규수신", marker_color="#4C72B0")
    figg.add_bar(x=g_sum["grade"], y=g_sum["out"], name="수신거부", marker_color="#C44E52")
    figg.update_layout(barmode="group", height=340, title="등급별 신규 vs 거부",
                       margin=dict(t=40, b=10), legend_title_text="")
    st.plotly_chart(figg, use_container_width=True)
with colB:
    figr = px.bar(g_sum, x="grade", y="out_rate", color="out_rate", color_continuous_scale="OrRd",
                  labels={"out_rate": "수신거부율(%)", "grade": "등급"}, title="등급별 수신거부율")
    figr.update_layout(height=340, margin=dict(t=40, b=10), coloraxis_showscale=False)
    st.plotly_chart(figr, use_container_width=True)

# ── 상세 테이블 ────────────────────────────────────────────
section("상세 데이터", "등급 × 채널 집계 (선택 기간)")
detail = (fg.groupby(["grade", "channel"])
            .agg(유효회원수=("act", "max"), 수신자수=("tot", "max"),
                 신규수신=("new", "sum"), 수신거부=("out", "sum")).reset_index())
detail["순증감"] = detail["신규수신"] - detail["수신거부"]
detail["수신거부율(%)"] = np.where(detail["수신자수"] > 0,
                                  detail["수신거부"] / detail["수신자수"] * 100, 0).round(3)
detail = detail.rename(columns={"grade": "등급", "channel": "채널"})
st.dataframe(detail, use_container_width=True, hide_index=True,
             column_config={
                 "유효회원수": st.column_config.NumberColumn(format="localized"),
                 "수신자수": st.column_config.NumberColumn(format="localized"),
                 "신규수신": st.column_config.NumberColumn(format="localized"),
                 "수신거부": st.column_config.NumberColumn(format="localized"),
                 "순증감": st.column_config.NumberColumn(format="localized"),
             })

csv = detail.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ 집계 CSV 다운로드", csv, "optout_summary.csv", "text/csv")

st.caption("ⓘ 수신거부율 = 수신거부(OUT) / 수신자수(TOT) · 순증감 = 신규수신(NEW) − 수신거부(OUT) · 도달률 = 수신자수 / 유효회원수")
