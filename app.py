from __future__ import annotations

import base64
import html
import io
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from streamlit_option_menu import option_menu

st.set_page_config(
    page_title="YVF Management Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "data" / "YVF_PowerBI_Ready.xlsx"
ASSET_DIR = BASE_DIR / "assets"

NAVY = "#07376E"
DARK_NAVY = "#052A59"
BLUE = "#0864C7"
ORANGE = "#FF7900"
RED = "#F04438"
GREEN = "#16A34A"
LIGHT_BG = "#F5F7FA"
BORDER = "#D9E2EC"
TEXT = "#24364B"
MUTED = "#667085"


def data_uri(path: Path) -> str:
    if not path.exists():
        return ""
    mime = "image/svg+xml" if path.suffix.lower() == ".svg" else "image/png"
    return f"data:{mime};base64,{base64.b64encode(path.read_bytes()).decode('ascii')}"


LOGO_URI = data_uri(ASSET_DIR / "yusen_logo.png")
BOOKING_ICON = data_uri(ASSET_DIR / "icon_booking.svg")
SI_ICON = data_uri(ASSET_DIR / "icon_si.svg")
ISSUE_ICON = data_uri(ASSET_DIR / "icon_issue.svg")

st.markdown(
    f"""
<style>
:root {{ --navy:{NAVY}; --blue:{BLUE}; --orange:{ORANGE}; --red:{RED}; }}
html, body, [class*="css"] {{ font-family:"Segoe UI", Arial, sans-serif; }}
.stApp {{ background:{LIGHT_BG}; color:{TEXT}; }}
.block-container {{ padding:1.0rem 1.35rem 0.7rem 1.35rem; max-width:1700px; }}
header[data-testid="stHeader"] {{ height:0; background:transparent; }}
[data-testid="stToolbar"], #MainMenu, footer {{ display:none !important; }}
[data-testid="stSidebar"] {{ background:linear-gradient(180deg,{DARK_NAVY} 0%, {NAVY} 68%, #05244A 100%); border-right:none; }}
[data-testid="stSidebar"] > div:first-child {{ padding-top:0; }}
[data-testid="stSidebar"] * {{ color:white; }}
[data-testid="stSidebar"] .stFileUploader section {{ background:rgba(255,255,255,.08); border:1px dashed rgba(255,255,255,.45); }}
[data-testid="stSidebar"] .stFileUploader small {{ color:#D5E3F3; }}
[data-testid="stSidebar"] div[data-baseweb="select"] > div,
[data-testid="stSidebar"] div[data-baseweb="input"] {{ background:rgba(255,255,255,.08); border-color:rgba(255,255,255,.34); }}
[data-testid="stSidebar"] hr {{ border-color:rgba(255,255,255,.18); }}
.sidebar-logo {{ background:#fff; margin:0 -1rem .75rem -1rem; padding:12px 22px 9px; text-align:center; border-radius:0 0 8px 0; }}
.sidebar-logo img {{ max-width:180px; width:100%; }}
.side-section {{ color:#BFD1E6; font-size:12px; font-weight:800; letter-spacing:.7px; margin:14px 0 3px; }}
.side-tagline {{ margin-top:25px; font-size:16px; line-height:1.25; font-weight:600; }}
.side-wave {{ height:8px; margin-top:12px; border-radius:9px; background:linear-gradient(165deg,transparent 0 20%,#0A78D8 21% 43%,transparent 44% 50%,#E33B35 51% 63%,transparent 64%); }}
.top-ribbon {{ position:fixed; top:0; right:0; width:235px; height:72px; z-index:0; clip-path:polygon(37% 0,100% 0,100% 100%,0 100%); background:linear-gradient(122deg,transparent 0 22%,#1871BE 22% 38%,#73B7E2 38% 51%,#fff 51% 59%,#FF7A24 59% 75%,#EF3E36 75% 100%); opacity:.98; }}
.dashboard-header {{ position:relative; z-index:1; display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:13px; padding-right:16px; }}
.main-title {{ color:{NAVY}; font-size:29px; font-weight:850; letter-spacing:.2px; line-height:1.05; }}
.sub-title {{ color:{MUTED}; font-size:14px; margin-top:7px; }}
.report-time {{ color:{NAVY}; font-weight:700; font-size:12px; padding:5px 8px; margin-right:105px; white-space:nowrap; }}
.filter-panel {{ background:white; border:1px solid {BORDER}; border-radius:10px; box-shadow:0 2px 8px rgba(31,54,79,.06); padding:5px 15px 10px; margin-bottom:14px; }}
.filter-title {{ color:{NAVY}; font-size:14px; font-weight:800; margin:2px 0 3px; }}
.kpi-card {{ background:white; border:1px solid {BORDER}; border-left:5px solid var(--accent); border-radius:10px; min-height:128px; padding:15px 18px; box-shadow:0 2px 9px rgba(31,54,79,.09); display:grid; grid-template-columns:76px 1fr; align-items:center; gap:10px; }}
.kpi-icon {{ width:66px; height:66px; border-radius:50%; background:var(--soft); display:flex; align-items:center; justify-content:center; }}
.kpi-icon img {{ width:42px; height:42px; }}
.kpi-label {{ color:var(--accent); font-weight:850; font-size:14px; text-align:center; text-transform:uppercase; }}
.kpi-value {{ color:var(--accent); font-size:42px; font-weight:850; text-align:center; line-height:1.06; margin:5px 0; }}
.kpi-note {{ color:#555; font-size:11.5px; text-align:center; }}
.kpi-note b {{ color:{GREEN}; }}
.panel {{ background:white; border:1px solid {BORDER}; border-radius:9px; box-shadow:0 2px 7px rgba(31,54,79,.07); padding:10px 11px 9px; height:100%; }}
.panel-title {{ color:{NAVY}; font-size:14px; font-weight:850; border-left:5px solid {BLUE}; padding-left:7px; margin:1px 0 9px; text-transform:uppercase; }}
.panel-title.orange {{ border-left-color:{ORANGE}; }}
table.yvf {{ width:100%; border-collapse:collapse; font-size:12px; }}
table.yvf th {{ background:{NAVY}; color:white; text-align:center; padding:7px 6px; border:1px solid #7792B2; font-weight:700; }}
table.yvf th.orange {{ background:{ORANGE}; border-color:#FFAD63; }}
table.yvf td {{ padding:7px 7px; border:1px solid #E0E6ED; text-align:center; color:#26364A; background:#fff; }}
table.yvf td.left {{ text-align:left; }}
table.yvf tr.total td {{ background:#EAF3FC; font-weight:850; color:{NAVY}; }}
.empty {{ padding:25px; text-align:center; color:{MUTED}; font-size:12px; }}
.footer-yvf {{ text-align:center; color:#566270; font-size:11px; margin-top:10px; border-top:1px solid #E7ECF1; padding:8px 0 0; }}
.stDownloadButton button {{ width:100%; background:transparent; color:white; border:1px solid rgba(255,255,255,.55); }}
[data-testid="stPlotlyChart"] {{ margin-top:-7px; }}
@media(max-width:1050px) {{ .kpi-card{{grid-template-columns:1fr;}} .kpi-icon{{margin:auto;}} .report-time{{margin-right:0;}} }}
</style>
<div class="top-ribbon"></div>
""",
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    result: dict[str, pd.DataFrame] = {}
    for sheet in ["Data_Booking", "Data_SI", "Data_Issue", "Customer_Feedback"]:
        if sheet in xls.sheet_names:
            result[sheet] = pd.read_excel(xls, sheet_name=sheet).dropna(how="all")
        else:
            result[sheet] = pd.DataFrame()
    return result


def safe_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def esc(value: object) -> str:
    return html.escape("" if pd.isna(value) else str(value))


def html_table(df: pd.DataFrame, total_row: bool = False, left: set[str] | None = None, orange_header: bool = False) -> str:
    left = left or set()
    if df.empty:
        return '<div class="empty">Không có dữ liệu phù hợp với bộ lọc.</div>'
    hclass = ' class="orange"' if orange_header else ""
    heads = "".join(f"<th{hclass}>{esc(c)}</th>" for c in df.columns)
    rows: list[str] = []
    for i, (_, row) in enumerate(df.iterrows()):
        cls = "total" if total_row and i == len(df) - 1 else ""
        cells = "".join(f'<td class="{"left" if col in left else ""}">{esc(row[col])}</td>' for col in df.columns)
        rows.append(f'<tr class="{cls}">{cells}</tr>')
    return f'<table class="yvf"><thead><tr>{heads}</tr></thead><tbody>{"".join(rows)}</tbody></table>'


def apply_filters(df: pd.DataFrame, customers: list[str], modes: list[str], statuses: list[str], date_range) -> pd.DataFrame:
    out = df.copy()
    if "Booking Date" in out.columns:
        out["Booking Date"] = pd.to_datetime(out["Booking Date"], errors="coerce")
        if date_range and len(date_range) == 2:
            out = out[out["Booking Date"].between(pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1]))]
    if customers and "Customer" in out.columns:
        out = out[out["Customer"].astype(str).isin(customers)]
    if modes and "Mode" in out.columns:
        out = out[out["Mode"].astype(str).isin(modes)]
    if statuses and "Status" in out.columns:
        out = out[out["Status"].astype(str).isin(statuses)]
    return out


def kpi(label: str, value: str, icon_uri: str, accent: str, soft: str, note: str) -> str:
    return f'''<div class="kpi-card" style="--accent:{accent};--soft:{soft}">
      <div class="kpi-icon"><img src="{icon_uri}" alt=""></div>
      <div><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-note">{note}</div></div>
    </div>'''


with st.sidebar:
    st.markdown(f'<div class="sidebar-logo"><img src="{LOGO_URI}" alt="Yusen Logistics"></div>', unsafe_allow_html=True)
    page = option_menu(
        menu_title=None,
        options=["Overview", "Booking & SI", "SI Submission", "Issues", "Customer Feedback", "Raw Data"],
        icons=["speedometer2", "clipboard-data", "file-earmark-text", "exclamation-triangle", "star", "table"],
        default_index=0,
        styles={
            "container": {"padding": "0", "background-color": "transparent"},
            "icon": {"color": "white", "font-size": "18px"},
            "nav-link": {"font-size": "14px", "text-align": "left", "margin": "4px 0", "padding": "11px 12px", "border-radius": "7px", "color": "white"},
            "nav-link-selected": {"background-color": "#096BD5", "font-weight": "700"},
        },
    )
    st.markdown('<div class="side-section">DATA SOURCE</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload Excel", type=["xlsx"], label_visibility="collapsed")

file_bytes = uploaded.getvalue() if uploaded is not None else DATA_FILE.read_bytes()
try:
    data = load_excel(file_bytes)
except Exception as exc:
    st.error(f"Không đọc được file Excel: {exc}")
    st.stop()

booking = data["Data_Booking"].copy()
si = data["Data_SI"].copy()
issues = data["Data_Issue"].copy()
feedback = data["Customer_Feedback"].copy()
for frame in (booking, si):
    if "Booking Date" in frame.columns:
        frame["Booking Date"] = pd.to_datetime(frame["Booking Date"], errors="coerce")

all_dates = pd.concat([booking.get("Booking Date", pd.Series(dtype="datetime64[ns]")), si.get("Booking Date", pd.Series(dtype="datetime64[ns]"))]).dropna()
today = datetime.today().date()
min_date = all_dates.min().date() if not all_dates.empty else today
max_date = all_dates.max().date() if not all_dates.empty else today
customers = sorted(set(booking.get("Customer", pd.Series(dtype=str)).dropna().astype(str)) | set(si.get("Customer", pd.Series(dtype=str)).dropna().astype(str)))
modes = sorted(set(booking.get("Mode", pd.Series(dtype=str)).dropna().astype(str)) | set(si.get("Mode", pd.Series(dtype=str)).dropna().astype(str)))
statuses = sorted(set(booking.get("Status", pd.Series(dtype=str)).dropna().astype(str)) | set(si.get("Status", pd.Series(dtype=str)).dropna().astype(str)))

with st.sidebar:
    st.divider()
    st.markdown('<div class="side-section">FILTERS</div>', unsafe_allow_html=True)
    date_range = st.date_input("Date Range (ETD)", value=(min_date, max_date), min_value=min_date, max_value=max_date, format="DD/MM/YYYY")
    selected_customers = st.multiselect("Customer", customers, placeholder="All")
    selected_modes = st.multiselect("Mode", modes, placeholder="All")
    selected_statuses = st.multiselect("Status", statuses, placeholder="All")
    st.markdown('<div class="side-tagline">Together we<br>connect value</div><div class="side-wave"></div>', unsafe_allow_html=True)

booking_f = apply_filters(booking, selected_customers, selected_modes, selected_statuses, date_range)
si_f = apply_filters(si, selected_customers, selected_modes, selected_statuses, date_range)

st.markdown(f'''<div class="dashboard-header"><div><div class="main-title">YVF MANAGEMENT DASHBOARD</div>
<div class="sub-title">Overview of YVF System Performance</div></div>
<div class="report-time">▣&nbsp;&nbsp;{datetime.now():%d/%m/%Y %H:%M}</div></div>''', unsafe_allow_html=True)

booking_total = int(safe_num(booking_f.get("Booking Qty", pd.Series(dtype=float))).sum())
si_total = int(safe_num(si_f.get("SI Qty", pd.Series(dtype=float))).sum())
issue_total = len(issues)

if page == "Overview":
    c1, c2, c3 = st.columns(3)
    with c1: st.markdown(kpi("TOTAL BOOKINGS", f"{booking_total:,}", BOOKING_ICON, BLUE, "#EAF3FC", f"{len(booking_f):,} booking records"), unsafe_allow_html=True)
    with c2: st.markdown(kpi("TOTAL SI SUBMISSIONS", f"{si_total:,}", SI_ICON, ORANGE, "#FFF0E2", f"{len(si_f):,} SI records"), unsafe_allow_html=True)
    with c3: st.markdown(kpi("TOTAL ISSUES", f"{issue_total:,}", ISSUE_ICON, RED, "#FDE9E7", "Items in issue list"), unsafe_allow_html=True)
    st.write("")

    left_col, right_col = st.columns([1.08, .92])
    with left_col:
        cols = ["Customer", "Mode", "Total Booking", "Normal", "Slow", "Closed"]
        summary = pd.DataFrame(columns=cols)
        if not booking_f.empty:
            work = booking_f.copy(); work["Booking Qty"] = safe_num(work.get("Booking Qty", pd.Series(index=work.index, dtype=float)))
            rows = []
            for (customer, mode), g in work.groupby(["Customer", "Mode"], dropna=False):
                status = g.get("Status", pd.Series(index=g.index, dtype=str)).astype(str).str.lower()
                rows.append({"Customer":customer,"Mode":mode,"Total Booking":int(g["Booking Qty"].sum()),"Normal":int(g.loc[status.eq("normal"),"Booking Qty"].sum()),"Slow":int(g.loc[status.eq("slow"),"Booking Qty"].sum()),"Closed":int(g.loc[status.eq("closed"),"Booking Qty"].sum())})
            summary = pd.DataFrame(rows)
            total = {"Customer":"TOTAL","Mode":"","Total Booking":int(summary["Total Booking"].sum()),"Normal":int(summary["Normal"].sum()),"Slow":int(summary["Slow"].sum()),"Closed":int(summary["Closed"].sum())}
            summary = pd.concat([summary, pd.DataFrame([total])], ignore_index=True)
        st.markdown('<div class="panel"><div class="panel-title">BOOKING SUMMARY</div>'+html_table(summary, True)+'</div>', unsafe_allow_html=True)

    with right_col:
        cols = ["Customer", "Total SI", "On-time", "Late", "On-time Rate"]
        si_summary = pd.DataFrame(columns=cols)
        if not si_f.empty:
            work = si_f.copy(); work["SI Qty"] = safe_num(work.get("SI Qty", pd.Series(index=work.index, dtype=float)))
            rows=[]
            for customer, g in work.groupby("Customer", dropna=False):
                total=int(g["SI Qty"].sum()); status=g.get("Status",pd.Series(index=g.index,dtype=str)).astype(str).str.lower(); ontime=int(g.loc[status.isin(["normal","on-time","ontime"]),"SI Qty"].sum()); late=total-ontime
                rows.append({"Customer":customer,"Total SI":total,"On-time":ontime,"Late":late,"On-time Rate":f"{ontime/total*100:.1f}%" if total else "0.0%"})
            si_summary=pd.DataFrame(rows); t=int(si_summary["Total SI"].sum()); o=int(si_summary["On-time"].sum())
            si_summary=pd.concat([si_summary,pd.DataFrame([{"Customer":"TOTAL","Total SI":t,"On-time":o,"Late":int(si_summary["Late"].sum()),"On-time Rate":f"{o/t*100:.1f}%" if t else "0.0%"}])],ignore_index=True)
        st.markdown('<div class="panel"><div class="panel-title">SI SUBMISSION SUMMARY</div>'+html_table(si_summary, True)+'</div>', unsafe_allow_html=True)

    st.write("")
    b1,b2,b3=st.columns([1.07,1,1.02])
    with b1:
        issue_cols=[c for c in ["Module","Suggestion"] if c in issues.columns]
        issue_view=issues[issue_cols].copy() if issue_cols else pd.DataFrame(columns=["Module","Suggestion / Improvement"])
        if len(issue_view.columns)==2: issue_view.columns=["Module","Suggestion / Improvement"]
        st.markdown('<div class="panel"><div class="panel-title orange">ISSUE LIST</div>'+html_table(issue_view,left=set(issue_view.columns),orange_header=True)+'</div>',unsafe_allow_html=True)
    with b2:
        if feedback.empty or "Feedback" not in feedback.columns: fb=pd.DataFrame(columns=["Feedback","Count"])
        else:
            fb=feedback.groupby("Feedback",as_index=False).size().rename(columns={"size":"Count"}); fb["Feedback"]="♡  "+fb["Feedback"].astype(str); fb=pd.concat([fb,pd.DataFrame([{"Feedback":"TOTAL","Count":int(fb["Count"].sum())}])],ignore_index=True)
        st.markdown('<div class="panel"><div class="panel-title">CUSTOMER FEEDBACK</div>'+html_table(fb,True,{"Feedback"})+'</div>',unsafe_allow_html=True)
    with b3:
        order=["Normal","Slow","Closed"]; qty={}
        for status_name in order:
            mask=booking_f.get("Status",pd.Series(index=booking_f.index,dtype=str)).astype(str).str.lower().eq(status_name.lower())
            qty[status_name]=int(safe_num(booking_f.loc[mask,"Booking Qty"] if "Booking Qty" in booking_f else pd.Series(dtype=float)).sum())
        vals=[qty[x] for x in order]; total=sum(vals)
        fig=go.Figure(go.Pie(labels=order,values=vals,hole=.62,marker=dict(colors=[BLUE,ORANGE,RED]),textinfo="none",sort=False))
        fig.update_layout(height=252,margin=dict(l=3,r=3,t=2,b=2),showlegend=True,legend=dict(orientation="v",x=.72,y=.74,font=dict(size=11)),paper_bgcolor="white",annotations=[dict(text=f"<b>{total}</b><br>Total",x=.29,y=.5,font_size=18,showarrow=False)])
        st.markdown('<div class="panel"><div class="panel-title">STATUS OVERVIEW</div>',unsafe_allow_html=True); st.plotly_chart(fig,use_container_width=True,config={"displayModeBar":False}); st.markdown('</div>',unsafe_allow_html=True)

elif page == "Booking & SI":
    st.markdown('<div class="panel"><div class="panel-title">BOOKING DETAIL</div></div>', unsafe_allow_html=True)
    st.dataframe(booking_f, use_container_width=True, hide_index=True, height=310)
    st.markdown('<div class="panel"><div class="panel-title">SI DETAIL</div></div>', unsafe_allow_html=True)
    st.dataframe(si_f, use_container_width=True, hide_index=True, height=310)
elif page == "SI Submission":
    st.markdown('<div class="panel"><div class="panel-title">SI SUBMISSION</div></div>', unsafe_allow_html=True)
    st.dataframe(si_f, use_container_width=True, hide_index=True, height=620)
elif page == "Issues":
    st.markdown('<div class="panel"><div class="panel-title orange">ISSUE & IMPROVEMENT LIST</div></div>', unsafe_allow_html=True)
    st.dataframe(issues, use_container_width=True, hide_index=True, height=620)
elif page == "Customer Feedback":
    st.markdown('<div class="panel"><div class="panel-title">CUSTOMER FEEDBACK</div></div>', unsafe_allow_html=True)
    st.dataframe(feedback, use_container_width=True, hide_index=True, height=620)
else:
    tabs=st.tabs(["Data_Booking","Data_SI","Data_Issue","Customer_Feedback"])
    for tab,(name,frame) in zip(tabs,data.items()):
        with tab: st.dataframe(frame,use_container_width=True,hide_index=True,height=570)

st.markdown('<div class="footer-yvf">YVF Dashboard | Confidential &amp; Internal Use Only</div>', unsafe_allow_html=True)
