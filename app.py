
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
    page_title="YVF Adoption Dashboard – CS HAD",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = (
    BASE_DIR / "data" / "YVF_PowerBI_Ready.xlsx"
    if (BASE_DIR / "data" / "YVF_PowerBI_Ready.xlsx").exists()
    else BASE_DIR / "YVF_PowerBI_Ready.xlsx"
)
ASSET_DIR = BASE_DIR / "assets" if (BASE_DIR / "assets").exists() else BASE_DIR

NAVY = "#07376E"
DARK_NAVY = "#052A59"
BLUE = "#0864C7"
ORANGE = "#F58A24"
RED = "#F04438"
GREEN = "#16A34A"
AMBER = "#F59E0B"
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
html, body, [class*="css"] {{ font-family:"Segoe UI", Arial, sans-serif; }}
.stApp {{ background:{LIGHT_BG}; color:{TEXT}; }}
.block-container {{ padding:1rem 1.35rem .7rem; max-width:1700px; }}
header[data-testid="stHeader"] {{ height:0; background:transparent; }}
[data-testid="stToolbar"], #MainMenu, footer {{ display:none !important; }}
[data-testid="stSidebar"] {{
    background:linear-gradient(180deg,{DARK_NAVY} 0%, {NAVY} 68%, #05244A 100%);
    border-right:none;
}}
[data-testid="stSidebar"] * {{ color:white; }}
.sidebar-logo {{
    background:#fff; margin:0 -1rem .75rem; padding:12px 22px 9px;
    text-align:center; border-radius:0 0 8px 0;
}}
.sidebar-logo img {{ max-width:180px; width:100%; }}
.side-section {{
    color:#BFD1E6; font-size:12px; font-weight:800;
    letter-spacing:.7px; margin:14px 0 3px;
}}
.side-tagline {{ margin-top:25px; font-size:16px; line-height:1.25; font-weight:600; }}
.side-wave {{
    height:8px; margin-top:12px; border-radius:9px;
    background:linear-gradient(165deg,transparent 0 20%,#0A78D8 21% 43%,
    transparent 44% 50%,#E33B35 51% 63%,transparent 64%);
}}
.top-ribbon {{
    position:fixed; top:0; right:0; width:235px; height:72px; z-index:0;
    clip-path:polygon(37% 0,100% 0,100% 100%,0 100%);
    background:linear-gradient(122deg,transparent 0 22%,#1871BE 22% 38%,
    #73B7E2 38% 51%,#fff 51% 59%,#FF7A24 59% 75%,#EF3E36 75% 100%);
}}
.dashboard-header {{
    position:relative; z-index:1; display:flex; justify-content:space-between;
    align-items:flex-start; margin-bottom:13px; padding-right:16px;
}}
.main-title {{ color:{NAVY}; font-size:29px; font-weight:850; line-height:1.05; }}
.sub-title {{ color:{MUTED}; font-size:14px; margin-top:7px; }}
.report-time {{
    color:{NAVY}; font-weight:700; font-size:12px;
    padding:5px 8px; margin-right:105px; white-space:nowrap;
}}
.kpi-card {{
    background:white; border:1px solid {BORDER}; border-left:5px solid var(--accent);
    border-radius:10px; min-height:128px; padding:15px 18px;
    box-shadow:0 2px 9px rgba(31,54,79,.09);
    display:grid; grid-template-columns:76px 1fr; align-items:center; gap:10px;
}}
.kpi-icon {{
    width:66px; height:66px; border-radius:50%; background:var(--soft);
    display:flex; align-items:center; justify-content:center;
}}
.kpi-icon img {{ width:42px; height:42px; }}
.kpi-label {{
    color:var(--accent); font-weight:850; font-size:14px;
    text-align:center; text-transform:uppercase;
}}
.kpi-value {{
    color:var(--accent); font-size:42px; font-weight:850;
    text-align:center; line-height:1.06; margin:5px 0;
}}
.kpi-note {{ color:#555; font-size:11.5px; text-align:center; }}
.panel {{
    background:white; border:1px solid {BORDER}; border-radius:9px;
    box-shadow:0 2px 7px rgba(31,54,79,.07);
    padding:10px 11px 9px; height:100%;
}}
.panel-title {{
    color:{NAVY}; font-size:14px; font-weight:850;
    border-left:5px solid {BLUE}; padding-left:7px;
    margin:1px 0 9px; text-transform:uppercase;
}}
.panel-title.orange {{ border-left-color:{ORANGE}; }}
table.yvf {{ width:100%; border-collapse:collapse; font-size:12px; }}
table.yvf th {{
    background:{NAVY}; color:white; text-align:center;
    padding:7px 6px; border:1px solid #7792B2; font-weight:700;
}}
table.yvf th.orange {{ background:{ORANGE}; border-color:#FFB16D; }}
table.yvf td {{
    padding:7px; border:1px solid #E0E6ED;
    text-align:center; color:#26364A; background:#fff;
}}
table.yvf td.left {{ text-align:left; }}
table.yvf tr.total td {{ background:#EAF3FC; font-weight:850; color:{NAVY}; }}
.empty {{ padding:25px; text-align:center; color:{MUTED}; font-size:12px; }}
.footer-yvf {{
    text-align:center; color:#566270; font-size:11px;
    margin-top:10px; border-top:1px solid #E7ECF1; padding:8px 0 0;
}}
[data-testid="stPlotlyChart"] {{ margin-top:-7px; }}
@media(max-width:1050px) {{
    .kpi-card{{grid-template-columns:1fr;}}
    .kpi-icon{{margin:auto;}}
    .report-time{{margin-right:0;}}
}}
</style>
<div class="top-ribbon"></div>
""",
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    result = {}
    for sheet in ["Data_Booking", "Data_SI", "Data_Issue", "Customer_Feedback"]:
        result[sheet] = (
            pd.read_excel(xls, sheet_name=sheet).dropna(how="all")
            if sheet in xls.sheet_names else pd.DataFrame()
        )
    return result


def safe_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def esc(value: object) -> str:
    return html.escape("" if pd.isna(value) else str(value))


def html_table(
    df: pd.DataFrame,
    total_row: bool = False,
    left: set[str] | None = None,
    orange_header: bool = False,
) -> str:
    left = left or set()
    if df.empty:
        return '<div class="empty">Không có dữ liệu phù hợp với bộ lọc.</div>'
    hclass = ' class="orange"' if orange_header else ""
    heads = "".join(f"<th{hclass}>{esc(c)}</th>" for c in df.columns)
    rows = []
    for i, (_, row) in enumerate(df.iterrows()):
        cls = "total" if total_row and i == len(df) - 1 else ""
        cells = "".join(
            f'<td class="{"left" if col in left else ""}">{esc(row[col])}</td>'
            for col in df.columns
        )
        rows.append(f'<tr class="{cls}">{cells}</tr>')
    return (
        f'<table class="yvf"><thead><tr>{heads}</tr></thead>'
        f'<tbody>{"".join(rows)}</tbody></table>'
    )


def apply_filters(df, customers, modes, statuses, date_range):
    out = df.copy()
    if "Booking Date" in out.columns:
        out["Booking Date"] = pd.to_datetime(out["Booking Date"], errors="coerce")
        if date_range and len(date_range) == 2:
            out = out[
                out["Booking Date"].between(
                    pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
                )
            ]
    if customers and "Customer" in out.columns:
        out = out[out["Customer"].astype(str).isin(customers)]
    if modes and "Mode" in out.columns:
        out = out[out["Mode"].astype(str).isin(modes)]
    if statuses and "Status" in out.columns:
        out = out[out["Status"].astype(str).isin(statuses)]
    return out


def kpi(label, value, icon_uri, accent, soft, note):
    return f"""
    <div class="kpi-card" style="--accent:{accent};--soft:{soft}">
      <div class="kpi-icon"><img src="{icon_uri}" alt=""></div>
      <div>
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-note">{note}</div>
      </div>
    </div>
    """


def status_overview_chart(status_qty: dict[str, int]) -> go.Figure:
    """Create a stable Plotly horizontal status chart for Streamlit Cloud."""
    display_order = [
        ("Normal", BLUE),
        ("Slow", ORANGE),
        ("Very Slow", RED),
        ("Closed", GREEN),
    ]

    labels = []
    values = []
    colors = []
    total = sum(int(status_qty.get(name, 0)) for name, _ in display_order)

    for name, color in display_order:
        qty = int(status_qty.get(name, 0))
        if qty > 0:
            labels.append(name)
            values.append(qty)
            colors.append(color)

    if not values:
        fig = go.Figure()
        fig.add_annotation(
            text="No status data",
            x=0.5,
            y=0.5,
            showarrow=False,
            font=dict(size=14, color=MUTED),
        )
        fig.update_layout(
            height=235,
            margin=dict(l=10, r=10, t=10, b=10),
            paper_bgcolor="white",
            plot_bgcolor="white",
            xaxis=dict(visible=False),
            yaxis=dict(visible=False),
        )
        return fig

    percentages = [(v / total * 100) if total else 0 for v in values]
    custom_text = [f"{v} ({p:.0f}%)" for v, p in zip(values, percentages)]

    fig = go.Figure(
        go.Bar(
            x=percentages,
            y=labels,
            orientation="h",
            marker_color=colors,
            text=custom_text,
            textposition="outside",
            cliponaxis=False,
            hovertemplate="<b>%{y}</b><br>Qty: %{customdata[0]}<br>Rate: %{x:.1f}%<extra></extra>",
            customdata=[[v] for v in values],
        )
    )

    fig.update_layout(
        height=235,
        margin=dict(l=15, r=70, t=8, b=30),
        paper_bgcolor="white",
        plot_bgcolor="white",
        showlegend=False,
        bargap=0.48,
        xaxis=dict(
            range=[0, 110],
            showgrid=False,
            zeroline=False,
            showticklabels=False,
            title=None,
        ),
        yaxis=dict(
            autorange="reversed",
            showgrid=False,
            tickfont=dict(size=12, color=TEXT),
            title=None,
        ),
        font=dict(family="Segoe UI, Arial, sans-serif", color=TEXT),
    )

    slow_qty = int(status_qty.get("Slow", 0)) + int(status_qty.get("Very Slow", 0))
    slow_rate = (slow_qty / total * 100) if total else 0

    fig.add_annotation(
        x=0,
        y=-0.22,
        xref="paper",
        yref="paper",
        text=f"<b>Total Bookings:</b> {total}",
        showarrow=False,
        xanchor="left",
        font=dict(size=12, color=NAVY),
    )
    fig.add_annotation(
        x=1,
        y=-0.22,
        xref="paper",
        yref="paper",
        text=f"<b>Slow Rate:</b> {slow_rate:.0f}%",
        showarrow=False,
        xanchor="right",
        font=dict(size=12, color=NAVY),
    )
    return fig


with st.sidebar:
    st.markdown(
        f'<div class="sidebar-logo"><img src="{LOGO_URI}" alt="Yusen Logistics"></div>',
        unsafe_allow_html=True,
    )
    page = option_menu(
        menu_title=None,
        options=[
            "Overview",
            "Booking",
            "SI Submission",
            "YVF User Issues",
            "YVF Enhancement Requests",
            "Customer Feedback",
            "Raw Data",
        ],
        icons=[
            "speedometer2",
            "clipboard-data",
            "file-earmark-text",
            "exclamation-triangle",
            "lightbulb",
            "star",
            "table",
        ],
        default_index=0,
        styles={
            "container": {"padding": "0", "background-color": "transparent"},
            "icon": {"color": "white", "font-size": "18px"},
            "nav-link": {
                "font-size": "14px",
                "text-align": "left",
                "margin": "4px 0",
                "padding": "11px 12px",
                "border-radius": "7px",
                "color": "white",
            },
            "nav-link-selected": {
                "background-color": "#096BD5",
                "font-weight": "700",
            },
        },
    )
    st.markdown('<div class="side-section">DATA SOURCE</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Upload Excel", type=["xlsx"], label_visibility="collapsed"
    )

if uploaded is not None:
    file_bytes = uploaded.getvalue()
elif DATA_FILE.exists():
    file_bytes = DATA_FILE.read_bytes()
else:
    st.error("Không tìm thấy file YVF_PowerBI_Ready.xlsx.")
    st.stop()

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

all_dates = pd.concat(
    [
        booking.get("Booking Date", pd.Series(dtype="datetime64[ns]")),
        si.get("Booking Date", pd.Series(dtype="datetime64[ns]")),
    ]
).dropna()

today = datetime.today().date()
min_date = all_dates.min().date() if not all_dates.empty else today
max_date = all_dates.max().date() if not all_dates.empty else today

customers = sorted(
    set(booking.get("Customer", pd.Series(dtype=str)).dropna().astype(str))
    | set(si.get("Customer", pd.Series(dtype=str)).dropna().astype(str))
)
modes = sorted(
    set(booking.get("Mode", pd.Series(dtype=str)).dropna().astype(str))
    | set(si.get("Mode", pd.Series(dtype=str)).dropna().astype(str))
)
statuses = sorted(
    set(booking.get("Status", pd.Series(dtype=str)).dropna().astype(str))
    | set(si.get("Status", pd.Series(dtype=str)).dropna().astype(str))
)

with st.sidebar:
    st.divider()
    st.markdown('<div class="side-section">FILTERS</div>', unsafe_allow_html=True)
    date_range = st.date_input(
        "Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
        format="DD/MM/YYYY",
    )
    selected_customers = st.multiselect("Customer", customers, placeholder="All")
    selected_modes = st.multiselect("Mode", modes, placeholder="All")
    selected_statuses = st.multiselect("Status", statuses, placeholder="All")
    st.markdown(
        '<div class="side-tagline">Together we<br>connect value</div>'
        '<div class="side-wave"></div>',
        unsafe_allow_html=True,
    )

booking_f = apply_filters(
    booking, selected_customers, selected_modes, selected_statuses, date_range
)
si_f = apply_filters(
    si, selected_customers, selected_modes, selected_statuses, date_range
)

st.markdown(
    f"""
    <div class="dashboard-header">
      <div>
        <div class="main-title">YVF ADOPTION DASHBOARD – CS HAD</div>
        <div class="sub-title">Customer Adoption &amp; Usage Monitoring</div>
      </div>
      <div class="report-time">▣&nbsp;&nbsp;Last refresh: {datetime.now():%d-%b-%Y %H:%M}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

booking_total = int(
    safe_num(booking_f.get("Booking Qty", pd.Series(dtype=float))).sum()
)
si_total = int(safe_num(si_f.get("SI Qty", pd.Series(dtype=float))).sum())

all_known_customers = sorted(
    set(booking.get("Customer", pd.Series(dtype=str)).dropna().astype(str))
    | set(si.get("Customer", pd.Series(dtype=str)).dropna().astype(str))
)
active_customers = sorted(
    set(booking_f.get("Customer", pd.Series(dtype=str)).dropna().astype(str))
    | set(si_f.get("Customer", pd.Series(dtype=str)).dropna().astype(str))
)
adoption_rate = (
    len(active_customers) / len(all_known_customers) * 100
    if all_known_customers else 0
)

issue_rows = []
for source_name, frame in [("Booking", booking_f), ("SI Submission", si_f)]:
    for _, row in frame.iterrows():
        system_issue = str(row.get("System Issue", "")).strip().lower()
        remark = "" if pd.isna(row.get("Remark")) else str(row.get("Remark")).strip()
        if system_issue in {"yes", "y", "true", "1"} or remark:
            issue_rows.append(
                {
                    "Date": row.get("Booking Date", ""),
                    "Customer": row.get("Customer", ""),
                    "Module": source_name,
                    "Issue": remark or "System issue reported",
                    "PIC": row.get("PIC", ""),
                    "Status": row.get("Status", ""),
                }
            )

user_issues = pd.DataFrame(issue_rows)
if not user_issues.empty:
    user_issues["Date"] = pd.to_datetime(
        user_issues["Date"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")

enhancement_requests = issues.rename(
    columns={"Suggestion": "Enhancement Request"}
).copy()

if page == "Overview":
    c1, c2, c3, c4, c5 = st.columns(5)
    cards = [
        (
            c1, "CUSTOMER ADOPTION RATE", f"{adoption_rate:.0f}%",
            BOOKING_ICON, BLUE, "#EAF3FC",
            f"{len(active_customers)}/{len(all_known_customers)} customers"
        ),
        (
            c2, "ACTIVE CUSTOMERS", str(len(active_customers)),
            BOOKING_ICON, GREEN, "#EAF8EF", "Customers with activity"
        ),
        (
            c3, "BOOKING VIA YVF %", "100%" if booking_total else "0%",
            BOOKING_ICON, ORANGE, "#FFF0E2",
            f"{booking_total} bookings in scope"
        ),
        (
            c4, "YVF ENHANCEMENT REQUESTS", str(len(enhancement_requests)),
            SI_ICON, NAVY, "#EAF0F7", "Improvement requests"
        ),
        (
            c5, "YVF USER ISSUES", str(len(user_issues)),
            ISSUE_ICON, RED, "#FDE9E7", "Reported user issues"
        ),
    ]
    for col, *args in cards:
        with col:
            st.markdown(kpi(*args), unsafe_allow_html=True)

    st.write("")
    left_col, right_col = st.columns([1.08, .92])

    with left_col:
        summary = pd.DataFrame(
            columns=["Customer", "Mode", "Total Booking", "Normal", "Slow", "Closed"]
        )
        if not booking_f.empty:
            work = booking_f.copy()
            work["Booking Qty"] = safe_num(work["Booking Qty"])
            rows = []
            for (customer, mode), g in work.groupby(
                ["Customer", "Mode"], dropna=False
            ):
                status = g["Status"].astype(str).str.lower()
                rows.append(
                    {
                        "Customer": customer,
                        "Mode": mode,
                        "Total Booking": int(g["Booking Qty"].sum()),
                        "Normal": int(g.loc[status.eq("normal"), "Booking Qty"].sum()),
                        "Slow": int(
                            g.loc[status.isin(["slow", "very slow"]), "Booking Qty"].sum()
                        ),
                        "Closed": int(g.loc[status.eq("closed"), "Booking Qty"].sum()),
                    }
                )
            summary = pd.DataFrame(rows)
            total_row = {
                "Customer": "TOTAL",
                "Mode": "",
                "Total Booking": int(summary["Total Booking"].sum()),
                "Normal": int(summary["Normal"].sum()),
                "Slow": int(summary["Slow"].sum()),
                "Closed": int(summary["Closed"].sum()),
            }
            summary = pd.concat(
                [summary, pd.DataFrame([total_row])], ignore_index=True
            )

        st.markdown(
            '<div class="panel"><div class="panel-title">BOOKING SUMMARY</div>'
            + html_table(summary, total_row=True)
            + "</div>",
            unsafe_allow_html=True,
        )

    with right_col:
        si_summary = pd.DataFrame(
            columns=["Customer", "Total SI", "On-time", "Late", "On-time Rate"]
        )
        if not si_f.empty:
            work = si_f.copy()
            work["SI Qty"] = safe_num(work["SI Qty"])
            rows = []
            for customer, g in work.groupby("Customer", dropna=False):
                total = int(g["SI Qty"].sum())
                status = g["Status"].astype(str).str.lower()
                ontime = int(
                    g.loc[status.isin(["normal", "on-time", "ontime"]), "SI Qty"].sum()
                )
                rows.append(
                    {
                        "Customer": customer,
                        "Total SI": total,
                        "On-time": ontime,
                        "Late": total - ontime,
                        "On-time Rate": f"{ontime / total * 100:.1f}%" if total else "0.0%",
                    }
                )
            si_summary = pd.DataFrame(rows)
            t = int(si_summary["Total SI"].sum())
            o = int(si_summary["On-time"].sum())
            si_summary = pd.concat(
                [
                    si_summary,
                    pd.DataFrame(
                        [{
                            "Customer": "TOTAL",
                            "Total SI": t,
                            "On-time": o,
                            "Late": int(si_summary["Late"].sum()),
                            "On-time Rate": f"{o / t * 100:.1f}%" if t else "0.0%",
                        }]
                    ),
                ],
                ignore_index=True,
            )

        st.markdown(
            '<div class="panel"><div class="panel-title">SI SUBMISSION SUMMARY</div>'
            + html_table(si_summary, total_row=True)
            + "</div>",
            unsafe_allow_html=True,
        )

    st.write("")
    b1, b2, b3 = st.columns([1.07, 1, 1.02])

    with b1:
        enhancement_cols = [
            c for c in ["Module", "Enhancement Request"]
            if c in enhancement_requests.columns
        ]
        enhancement_view = (
            enhancement_requests[enhancement_cols]
            if enhancement_cols
            else pd.DataFrame(columns=["Module", "Enhancement Request"])
        )
        st.markdown(
            '<div class="panel"><div class="panel-title orange">'
            "YVF ENHANCEMENT REQUESTS</div>"
            + html_table(
                enhancement_view,
                left=set(enhancement_view.columns),
                orange_header=True,
            )
            + "</div>",
            unsafe_allow_html=True,
        )

    with b2:
        if feedback.empty or "Feedback" not in feedback.columns:
            fb = pd.DataFrame(columns=["Feedback", "Count"])
        else:
            fb = (
                feedback.groupby("Feedback", as_index=False)
                .size()
                .rename(columns={"size": "Count"})
            )
            fb["Feedback"] = fb["Feedback"].astype(str).apply(
                lambda x: ("⚠  " if "chậm" in x.lower() or "đơ" in x.lower() else "✓  ") + x
            )
            fb = pd.concat(
                [fb, pd.DataFrame([{"Feedback": "TOTAL", "Count": int(fb["Count"].sum())}])],
                ignore_index=True,
            )
        st.markdown(
            '<div class="panel"><div class="panel-title">CUSTOMER FEEDBACK</div>'
            + html_table(fb, total_row=True, left={"Feedback"})
            + "</div>",
            unsafe_allow_html=True,
        )

    with b3:
        status_qty = {}
        for status_name in ["Normal", "Slow", "Very Slow", "Closed"]:
            mask = booking_f.get(
                "Status", pd.Series(index=booking_f.index, dtype=str)
            ).astype(str).str.lower().eq(status_name.lower())
            status_qty[status_name] = int(
                safe_num(
                    booking_f.loc[mask, "Booking Qty"]
                    if "Booking Qty" in booking_f.columns
                    else pd.Series(dtype=float)
                ).sum()
            )
        st.markdown(
            '<div class="panel"><div class="panel-title">STATUS OVERVIEW</div>',
            unsafe_allow_html=True,
        )
        st.plotly_chart(
            status_overview_chart(status_qty),
            use_container_width=True,
            config={"displayModeBar": False},
        )
        st.markdown("</div>", unsafe_allow_html=True)

elif page == "Booking":
    st.markdown(
        '<div class="panel"><div class="panel-title">BOOKING DETAIL</div></div>',
        unsafe_allow_html=True,
    )
    st.dataframe(booking_f, use_container_width=True, hide_index=True, height=620)

elif page == "SI Submission":
    st.markdown(
        '<div class="panel"><div class="panel-title">SI SUBMISSION</div></div>',
        unsafe_allow_html=True,
    )
    st.dataframe(si_f, use_container_width=True, hide_index=True, height=620)

elif page == "YVF User Issues":
    st.markdown(
        '<div class="panel"><div class="panel-title orange">YVF USER ISSUES</div></div>',
        unsafe_allow_html=True,
    )
    st.dataframe(user_issues, use_container_width=True, hide_index=True, height=620)

elif page == "YVF Enhancement Requests":
    st.markdown(
        '<div class="panel"><div class="panel-title">YVF ENHANCEMENT REQUESTS</div></div>',
        unsafe_allow_html=True,
    )
    st.dataframe(
        enhancement_requests, use_container_width=True, hide_index=True, height=620
    )

elif page == "Customer Feedback":
    st.markdown(
        '<div class="panel"><div class="panel-title">CUSTOMER FEEDBACK</div></div>',
        unsafe_allow_html=True,
    )
    st.dataframe(feedback, use_container_width=True, hide_index=True, height=620)

else:
    raw_frames = {
        "Data_Booking": booking.drop(columns=["Remark"], errors="ignore"),
        "Data_SI": si.drop(columns=["Remark"], errors="ignore"),
        "Data_Issue": enhancement_requests,
        "Customer_Feedback": feedback,
    }
    tabs = st.tabs(list(raw_frames.keys()))
    for tab, (_, frame) in zip(tabs, raw_frames.items()):
        with tab:
            st.dataframe(frame, use_container_width=True, hide_index=True, height=570)

st.markdown(
    '<div class="footer-yvf">YVF Adoption Dashboard – CS HAD | '
    "Confidential &amp; Internal Use Only</div>",
    unsafe_allow_html=True,
)
