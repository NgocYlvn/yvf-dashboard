
from __future__ import annotations

import io
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="YVF ADOPTION DASHBOARD – CS HAD",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.markdown("""
<style>

/* Ẩn thanh header của Streamlit */
header[data-testid="stHeader"]{
    display:none;
}

/* Ẩn menu */
#MainMenu{
    visibility:hidden;
}

/* Ẩn footer */
footer{
    visibility:hidden;
}

/* Đưa Dashboard sát lên trên */
.block-container{
    padding-top:0.4rem;
}

/* Loại bỏ khoảng trắng phía trên */
[data-testid="stAppViewContainer"]{
    margin-top:0rem;
}

</style>
""", unsafe_allow_html=True)

DEFAULT_FILE = Path("YVF_Adoption_Dashboard_CS_HAD.xlsm")

SHEET_NAMES = {
    "booking": "Data_Booking",
    "si": "Data_SI",
    "customer": "Data_CustomerActive",
    "proposal": "Data_Issue",
    "feedback": "Customer_Feedback",
}

YES_VALUES = {"yes", "y", "true", "1", "có", "co", "eligible", "active"}
CLOSED_VALUES = {"closed", "resolved", "completed", "done", "implemented"}
IN_PROGRESS_VALUES = {"in progress", "processing", "under review", "ongoing"}

st.markdown(
    """
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        .stApp {background: #F4F7FB;}
        .block-container {
            max-width: 1540px;
            padding-top: 1rem;
            padding-bottom: 2rem;
        }
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #17365D 0%, #102A49 100%);
        }
        [data-testid="stSidebar"] * {color: white;}
        .dashboard-title {
            color: #17365D;
            font-size: 30px;
            font-weight: 800;
            margin-bottom: 0;
        }
        .dashboard-subtitle {
            color: #6D7988;
            font-size: 13px;
            margin-top: 4px;
            margin-bottom: 16px;
        }
        .section-title {
            color: #17365D;
            font-size: 18px;
            font-weight: 800;
            margin-top: 20px;
            margin-bottom: 8px;
        }
        .kpi-card {
            background: white;
            border: 1px solid #E1E7EF;
            border-radius: 14px;
            box-shadow: 0 4px 14px rgba(23, 54, 93, 0.055);
            min-height: 126px;
            padding: 16px 18px;
        }
        .kpi-label {
            color: #6B7788;
            font-size: 11px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.45px;
        }
        .kpi-value {
            color: #17365D;
            font-size: 31px;
            font-weight: 850;
            line-height: 1.15;
            margin-top: 6px;
        }
        .kpi-note {
            color: #7C8795;
            font-size: 12px;
            line-height: 1.35;
            margin-top: 8px;
        }
        .action-box {
            background: white;
            border: 1px solid #E2E7EE;
            border-left: 5px solid #ED7D31;
            border-radius: 10px;
            margin: 8px 0;
            padding: 12px 15px;
        }
        .action-title {
            color: #17365D;
            font-size: 14px;
            font-weight: 800;
        }
        .action-text {
            color: #596675;
            font-size: 13px;
            margin-top: 3px;
        }
        .definition-box {
            background: #EEF4FB;
            border: 1px solid #D9E5F2;
            border-radius: 10px;
            padding: 11px 14px;
            color: #445466;
            font-size: 12px;
            margin-bottom: 10px;
        }
        div[data-testid="stDataFrame"] {
            background: white;
            border-radius: 10px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


def normalize_text(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def normalize_lower(series: pd.Series) -> pd.Series:
    return normalize_text(series).str.lower()


def safe_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def first_existing_column(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    lookup = {str(c).strip().lower(): str(c) for c in df.columns}
    for candidate in candidates:
        if candidate.strip().lower() in lookup:
            return lookup[candidate.strip().lower()]
    return None


def get_text(df: pd.DataFrame, candidates: Iterable[str]) -> pd.Series:
    col = first_existing_column(df, candidates)
    if col is None:
        return pd.Series("", index=df.index, dtype="object")
    return normalize_text(df[col])


def get_num(df: pd.DataFrame, candidates: Iterable[str]) -> pd.Series:
    col = first_existing_column(df, candidates)
    if col is None:
        return pd.Series(0.0, index=df.index, dtype="float64")
    return safe_numeric(df[col])


def yes_mask(series: pd.Series) -> pd.Series:
    return normalize_lower(series).isin(YES_VALUES)


def status_group(series: pd.Series) -> pd.Series:
    values = normalize_lower(series)

    def classify(v: str) -> str:
        if v in CLOSED_VALUES:
            return "Closed"
        if v in IN_PROGRESS_VALUES:
            return "In Progress"
        if v in {"open", "pending", "new", "active", ""}:
            return "Open"
        return v.title()

    return values.map(classify)


def format_number(value: float | int) -> str:
    return f"{value:,.0f}"


def format_percent(value: float) -> str:
    return f"{value:.1%}"


def show_title(title: str, subtitle: str) -> None:
    st.markdown(f'<div class="dashboard-title">{title}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="dashboard-subtitle">{subtitle}</div>', unsafe_allow_html=True)


def show_section(title: str) -> None:
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)


def kpi_card(label: str, value: str, note: str = "") -> None:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def action_box(title: str, text: str) -> None:
    st.markdown(
        f"""
        <div class="action-box">
            <div class="action-title">{title}</div>
            <div class="action-text">{text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def clean_chart(fig, height: int = 360):
    fig.update_layout(
        height=height,
        margin=dict(l=18, r=18, t=55, b=28),
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(color="#445466"),
        title_font=dict(color="#17365D", size=16),
        legend_title_text="",
    )
    fig.update_xaxes(showgrid=False, linecolor="#DDE3EA")
    fig.update_yaxes(gridcolor="#EDF1F5", zeroline=False)
    return fig


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    excel = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    result = {}
    for key, sheet in SHEET_NAMES.items():
        result[key] = (
            pd.read_excel(excel, sheet_name=sheet)
            if sheet in excel.sheet_names
            else pd.DataFrame()
        )
    return result


def prepare_data(raw: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    prepared = {}
    for key, df in raw.items():
        df = df.copy().dropna(how="all")
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = normalize_text(df[col])
        prepared[key] = df

    booking = prepared["booking"]
    customer = prepared["customer"]

    for col_name in ["Booking Qty", "Processing Time (Min)"]:
        col = first_existing_column(booking, [col_name])
        if col:
            booking[col] = safe_numeric(booking[col])

    booking_date = first_existing_column(booking, ["Booking Date", "Date"])
    if booking_date:
        booking[booking_date] = pd.to_datetime(booking[booking_date], errors="coerce")

    for col_name in [
        "Target Booking/Year",
        "Actual Booking YTD",
        "Achievement %",
        "Actual This Month",
    ]:
        col = first_existing_column(customer, [col_name])
        if col:
            customer[col] = safe_numeric(customer[col])

    start_date = first_existing_column(customer, ["Start Using YVF", "Activation Date"])
    if start_date:
        customer[start_date] = pd.to_datetime(customer[start_date], errors="coerce")

    return prepared


if not DEFAULT_FILE.exists():
    st.error(
        "Không tìm thấy file YVF_Adoption_Dashboard_CS_HAD.xlsm. "
        "Hãy đặt file Excel cùng thư mục với app.py."
    )
    st.stop()

try:
    data = prepare_data(load_excel(DEFAULT_FILE.read_bytes()))
except Exception as exc:
    st.error(f"Không thể đọc file Excel: {exc}")
    st.stop()

booking = data["booking"]
customer = data["customer"]
proposal = data["proposal"]
feedback = data["feedback"]

customer_name = get_text(customer, ["Customer", "Customer Name"])
customer_code = get_text(customer, ["Customer Code", "Code"])
actual_ytd_customer = get_num(customer, ["Actual Booking YTD", "Actual Bookings YTD"])
annual_target_customer = get_num(customer, ["Target Booking/Year", "Annual Booking Target"])
actual_month_customer = get_num(customer, ["Actual This Month", "Booking This Month"])

eligible_col = first_existing_column(
    customer, ["Eligible", "Eligible for YVF", "YVF Eligible", "Customer Target"]
)
active_col = first_existing_column(
    customer, ["Active", "YVF Active", "Adoption Status"]
)

eligible_mask = (
    yes_mask(customer[eligible_col])
    if eligible_col
    else customer_name.ne("")
)

active_mask = (
    yes_mask(customer[active_col])
    if active_col
    else actual_ytd_customer.gt(0)
)
active_mask = active_mask & eligible_mask

eligible_customers = int(eligible_mask.sum())
active_customers = int(active_mask.sum())
pending_adoption = max(eligible_customers - active_customers, 0)
adoption_rate = active_customers / eligible_customers if eligible_customers else 0

actual_ytd = float(actual_ytd_customer.sum())
annual_target = float(annual_target_customer.sum())
actual_this_month = float(actual_month_customer.sum())
target_achievement = actual_ytd / annual_target if annual_target else 0

booking_qty = get_num(booking, ["Booking Qty", "Booking Quantity"])
processing_time = get_num(booking, ["Processing Time (Min)", "Processing Time"])
booking_customer = get_text(booking, ["Customer", "Customer Name"])
booking_mode = get_text(booking, ["Mode"])
system_issue = get_text(booking, ["System Issue", "Issue"])

total_booking = float(booking_qty.sum())
avg_processing_time = float(processing_time.mean()) if len(processing_time) else 0
issue_mask = yes_mask(system_issue)
issue_rows = booking[issue_mask].copy() if not booking.empty else pd.DataFrame()
issue_count = len(issue_rows)
affected_customers = (
    int(get_text(issue_rows, ["Customer", "Customer Name"]).nunique())
    if not issue_rows.empty
    else 0
)

proposal_count = len(proposal)
proposal_module = get_text(proposal, ["Module", "Category"])
proposal_status_raw = get_text(proposal, ["Status", "Proposal Status"])
proposal_status = status_group(proposal_status_raw)
open_proposals = (
    int(proposal_status.eq("Open").sum())
    if not proposal.empty and proposal_status_raw.ne("").any()
    else proposal_count
)
implemented_proposals = (
    int(proposal_status.eq("Closed").sum())
    if not proposal.empty and proposal_status_raw.ne("").any()
    else 0
)

feedback_type = get_text(feedback, ["Type", "Feedback Type"])
negative_feedback = int(normalize_lower(feedback_type).eq("negative").sum())

st.sidebar.markdown("## YVF Dashboard - CS HAD")
page = st.sidebar.radio(
    "Navigation",
    [
        "1. Overview",
        "2. Customer Adoption",
        "3. Booking Performance",
        "4. User Issues",
        "5. Improvement Proposals",
    ],
    label_visibility="collapsed",
)
st.sidebar.markdown("---")


if page == "1. Overview":
    show_title(
        "YVF Adoption Dashboard – CS HAD",
        "Executive overview of customer adoption, booking performance, user issues, and improvement proposals.",
    )

    cols = st.columns(4)
    with cols[0]:
        kpi_card("Eligible Customers", format_number(eligible_customers),
                 "Customers eligible for YVF adoption")
    with cols[1]:
        kpi_card("Active Customers", format_number(active_customers),
                 f"{pending_adoption} customer(s) pending adoption")
    with cols[2]:
        kpi_card("Customer Adoption Rate", format_percent(adoption_rate),
                 "Active Customers ÷ Eligible Customers")
    with cols[3]:
        kpi_card("Actual Booking YTD", format_number(actual_ytd),
                 f"Annual target: {format_number(annual_target)}")

    st.write("")
    cols = st.columns(4)
    with cols[0]:
        kpi_card("Booking This Month", format_number(actual_this_month),
                 "Current-month YVF booking volume")
    with cols[1]:
        kpi_card("Avg. Booking Processing Time", f"{avg_processing_time:.1f} min",
                 "Average processing time per booking")
    with cols[2]:
        kpi_card("Open User Issues", format_number(issue_count),
                 f"{affected_customers} affected customer(s)")
    with cols[3]:
        kpi_card("Improvement Proposals", format_number(proposal_count),
                 f"{open_proposals} proposal(s) under follow-up")

    show_section("Adoption and Booking Snapshot")
    c1, c2 = st.columns([1, 1.4])

    with c1:
        gauge = go.Figure(go.Indicator(
            mode="gauge+number",
            value=adoption_rate * 100,
            number={"suffix": "%", "font": {"color": "#17365D", "size": 38}},
            title={"text": "Customer Adoption Rate"},
            gauge={
                "axis": {"range": [0, 100], "ticksuffix": "%"},
                "bar": {"color": "#ED7D31"},
                "steps": [
                    {"range": [0, 50], "color": "#F6E5D8"},
                    {"range": [50, 80], "color": "#E5EDF6"},
                    {"range": [80, 100], "color": "#DDEDE3"},
                ],
            },
        ))
        gauge.update_layout(height=350, margin=dict(l=20, r=20, t=55, b=15))
        st.plotly_chart(gauge, use_container_width=True)

    with c2:
        adoption_df = pd.DataFrame({
            "Status": ["Active Customers", "Pending Adoption"],
            "Customers": [active_customers, pending_adoption],
        })
        fig = px.bar(adoption_df, x="Status", y="Customers",
                     text="Customers", title="Active vs. Pending Adoption")
        fig.update_traces(textposition="outside")
        st.plotly_chart(clean_chart(fig, 350), use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        if not customer.empty:
            df = pd.DataFrame({
                "Customer": customer_code.where(customer_code.ne(""), customer_name),
                "Annual Target": annual_target_customer,
                "Actual Booking YTD": actual_ytd_customer,
            })
            df = df[df["Customer"].ne("")]
            fig = px.bar(
                df, x="Customer", y=["Annual Target", "Actual Booking YTD"],
                barmode="group", title="Actual Booking YTD vs. Annual Target",
                labels={"value": "Booking Volume", "variable": ""}
            )
            fig.update_layout(legend_orientation="h", legend_y=-0.22)
            st.plotly_chart(clean_chart(fig, 380), use_container_width=True)

    with c4:
        if not booking.empty:
            df = pd.DataFrame({
                "Customer": booking_customer,
                "Booking Volume": booking_qty,
            }).groupby("Customer", as_index=False)["Booking Volume"].sum()
            df = df[df["Customer"].ne("")].sort_values("Booking Volume")
            fig = px.bar(
                df, x="Booking Volume", y="Customer", orientation="h",
                title="Booking Volume by Customer", text="Booking Volume"
            )
            st.plotly_chart(clean_chart(fig, 380), use_container_width=True)

    show_section("Management Actions")
    action_box(
        "1. Customers Pending YVF Adoption",
        f"{pending_adoption} of {eligible_customers} eligible customer(s) have not yet generated a YVF booking."
    )
    action_box(
        "2. Open User Issues Requiring Follow-up",
        f"{issue_count} booking record(s) contain a system issue. Review the PIC, status, and required escalation."
    )
    action_box(
        "3. Improvement Proposals",
        f"{proposal_count} proposal(s) are recorded. Add Proposal Status, Owner, Target Date, and Implementation Date for follow-up."
    )

elif page == "2. Customer Adoption":
    show_title(
        "Customer Adoption",
        "Customer-level rollout status, adoption progress, and booking achievement.",
    )
    st.markdown(
        """
        <div class="definition-box">
        <b>Standard definition:</b> Eligible Customers are customers identified as suitable for YVF adoption.
        Active Customers are eligible customers that have generated at least one YVF booking.
        </div>
        """,
        unsafe_allow_html=True,
    )

    cols = st.columns(4)
    with cols[0]:
        kpi_card("Eligible Customers", format_number(eligible_customers))
    with cols[1]:
        kpi_card("Active Customers", format_number(active_customers))
    with cols[2]:
        kpi_card("Pending Adoption", format_number(pending_adoption))
    with cols[3]:
        kpi_card("Customer Adoption Rate", format_percent(adoption_rate))

    if customer.empty:
        st.info("No data is available in Data_CustomerActive.")
    else:
        detail = customer.copy()
        detail["Eligible for YVF"] = eligible_mask.map({True: "Yes", False: "No"})
        detail["Adoption Status"] = active_mask.map({True: "Active", False: "Pending"})
        detail["Calculated Achievement %"] = (
            actual_ytd_customer / annual_target_customer.replace(0, pd.NA)
        ).fillna(0)

        show_section("Customer Adoption Detail")
        preferred = [
            "Customer", "Customer Code", "Eligible for YVF", "Adoption Status",
            "Start Using YVF", "Target Booking/Year", "Actual Booking YTD",
            "Calculated Achievement %", "Actual This Month",
        ]
        display_cols = [c for c in preferred if c in detail.columns]
        st.dataframe(
            detail[display_cols],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Start Using YVF": st.column_config.DateColumn(format="DD-MMM-YYYY"),
                "Calculated Achievement %": st.column_config.ProgressColumn(
                    format="%.1f%%", min_value=0, max_value=1
                ),
            },
        )

        show_section("Adoption Progress by Customer")
        df = pd.DataFrame({
            "Customer": customer_code.where(customer_code.ne(""), customer_name),
            "Status": active_mask.map({True: "Active", False: "Pending Adoption"}),
            "Actual Booking YTD": actual_ytd_customer,
        })
        df = df[df["Customer"].ne("")]
        fig = px.bar(
            df.sort_values("Actual Booking YTD"),
            x="Actual Booking YTD", y="Customer", color="Status",
            orientation="h", title="Actual Booking YTD and Adoption Status",
            text="Actual Booking YTD"
        )
        st.plotly_chart(clean_chart(fig, 430), use_container_width=True)

elif page == "3. Booking Performance":
    show_title(
        "Booking Performance",
        "Booking volume, customer ranking, booking mode, and processing-time performance.",
    )

    sea_booking = float(booking_qty[normalize_lower(booking_mode).eq("sea")].sum())
    air_booking = float(booking_qty[normalize_lower(booking_mode).eq("air")].sum())

    cols = st.columns(4)
    with cols[0]:
        kpi_card("Total Booking Volume", format_number(total_booking))
    with cols[1]:
        kpi_card("Sea Booking Volume", format_number(sea_booking))
    with cols[2]:
        kpi_card("Air Booking Volume", format_number(air_booking))
    with cols[3]:
        kpi_card("Avg. Booking Processing Time", f"{avg_processing_time:.1f} min")

    if booking.empty:
        st.info("No data is available in Data_Booking.")
    else:
        customer_options = sorted(booking_customer[booking_customer.ne("")].unique())
        mode_options = sorted(booking_mode[booking_mode.ne("")].unique())

        f1, f2 = st.columns(2)
        with f1:
            selected_customers = st.multiselect(
                "Customer", customer_options, default=customer_options
            )
        with f2:
            selected_modes = st.multiselect(
                "Mode", mode_options, default=mode_options
            )

        mask = booking_customer.isin(selected_customers) & booking_mode.isin(selected_modes)
        filtered = booking[mask].copy()
        filtered_qty = booking_qty[mask]
        filtered_customer = booking_customer[mask]

        show_section("Booking Volume by Customer")
        summary = pd.DataFrame({
            "Customer": filtered_customer,
            "Booking Volume": filtered_qty,
        }).groupby("Customer", as_index=False)["Booking Volume"].sum()
        summary = summary.sort_values("Booking Volume", ascending=False)

        c1, c2 = st.columns([2, 1])
        with c1:
            fig = px.bar(
                summary.sort_values("Booking Volume"),
                x="Booking Volume", y="Customer", orientation="h",
                title="Booking Volume by Customer", text="Booking Volume"
            )
            st.plotly_chart(clean_chart(fig, 420), use_container_width=True)
        with c2:
            st.markdown("#### Top Customers")
            st.dataframe(summary.head(5), use_container_width=True, hide_index=True)
            st.markdown("#### Bottom Customers")
            st.dataframe(
                summary.tail(5).sort_values("Booking Volume"),
                use_container_width=True,
                hide_index=True,
            )

        show_section("Booking Data")
        st.dataframe(filtered, use_container_width=True, hide_index=True)

elif page == "4. User Issues":
    show_title(
        "User Issues",
        "Issues reported by customers or CS during the YVF booking process.",
    )

    issue_avg_time = (
        float(get_num(issue_rows, ["Processing Time (Min)", "Processing Time"]).mean())
        if not issue_rows.empty else 0
    )

    cols = st.columns(4)
    with cols[0]:
        kpi_card("Issue Reports", format_number(issue_count))
    with cols[1]:
        kpi_card("Affected Customers", format_number(affected_customers))
    with cols[2]:
        kpi_card("Avg. Processing Time", f"{issue_avg_time:.1f} min")
    with cols[3]:
        kpi_card("Negative Feedback", format_number(negative_feedback))

    show_section("Issue Log")
    if issue_rows.empty:
        st.info("No issue is currently marked in Data_Booking.")
    else:
        st.dataframe(issue_rows, use_container_width=True, hide_index=True)

    show_section("Customer Feedback")
    if feedback.empty:
        st.info("No data is available in Customer_Feedback.")
    else:
        st.dataframe(feedback, use_container_width=True, hide_index=True)

elif page == "5. Improvement Proposals":
    show_title(
        "Improvement Proposals",
        "Improvement proposals submitted by customers or CS.",
    )

    booking_module = int(proposal_module.str.contains("Booking", case=False, na=False).sum())
    si_module = int(proposal_module.str.contains("SI", case=False, na=False).sum())
    vgm_module = int(proposal_module.str.contains("VGM", case=False, na=False).sum())

    cols = st.columns(4)
    with cols[0]:
        kpi_card("Total Improvement Proposals", format_number(proposal_count))
    with cols[1]:
        kpi_card("Open Proposals", format_number(open_proposals))
    with cols[2]:
        kpi_card("Implemented Proposals", format_number(implemented_proposals))
    with cols[3]:
        kpi_card("Main Modules", format_number(booking_module + si_module + vgm_module),
                 "Booking / SI / VGM")

    show_section("Improvement Proposal Register")
    if proposal.empty:
        st.info("No data is available in Data_Issue.")
    else:
        display = proposal.rename(columns={
            "Issue ID": "Proposal ID",
            "Issue": "Current Limitation",
            "Impact": "Business Impact",
            "Suggestion": "Improvement Proposal",
        })
        st.dataframe(display, use_container_width=True, hide_index=True)

        if proposal_module.ne("").any():
            show_section("Proposals by Module")
            module_df = (
                proposal_module[proposal_module.ne("")]
                .value_counts()
                .rename_axis("Module")
                .reset_index(name="Proposals")
            )
            fig = px.bar(
                module_df.sort_values("Proposals"),
                x="Proposals", y="Module", orientation="h",
                title="Improvement Proposals by Module", text="Proposals"
            )
            st.plotly_chart(clean_chart(fig, 380), use_container_width=True)
