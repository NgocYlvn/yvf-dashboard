
from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="YVF Adoption Dashboard – CS HAD",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>

/* Chỉ ẩn menu và footer */
#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

/* KHÔNG ẩn header */

/* Ẩn khoảng trắng phía trên */
.block-container{
    padding-top:1rem;
}

/* Nếu muốn sát hẳn mép trên */
[data-testid="stAppViewContainer"]{
    margin-top:0;
}

/* Ẩn menu góc phải nếu cần */
#MainMenu{
    visibility:hidden;
}

footer{
    visibility:hidden;
}

</style>
""", unsafe_allow_html=True)
# =========================================================
# UI STYLE
# =========================================================
st.markdown(
    """
    <style>
        .stApp {
            background-color: #F5F7FA;
        }

        .block-container {
            max-width: 1500px;
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }

        [data-testid="stSidebar"] {
            background-color: #17365D;
        }

        [data-testid="stSidebar"] * {
            color: white;
        }

        .dashboard-title {
            color: #17365D;
            font-size: 30px;
            font-weight: 800;
            margin-bottom: 0;
        }

        .dashboard-subtitle {
            color: #6C7887;
            font-size: 14px;
            margin-top: 2px;
            margin-bottom: 18px;
        }

        .section-title {
            color: #17365D;
            font-size: 18px;
            font-weight: 750;
            margin-top: 18px;
            margin-bottom: 8px;
        }

        .kpi-card {
            background-color: white;
            border: 1px solid #E2E7EE;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(23, 54, 93, 0.06);
            min-height: 118px;
            padding: 16px 18px;
        }

        .kpi-label {
            color: #6B7788;
            font-size: 12px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.4px;
        }

        .kpi-value {
            color: #17365D;
            font-size: 30px;
            font-weight: 800;
            line-height: 1.15;
            margin-top: 5px;
        }

        .kpi-note {
            color: #7B8794;
            font-size: 12px;
            margin-top: 7px;
        }

        .attention-box {
            background-color: white;
            border: 1px solid #E3E8EF;
            border-left: 5px solid #ED7D31;
            border-radius: 9px;
            margin: 7px 0;
            padding: 12px 15px;
        }

        .attention-title {
            color: #17365D;
            font-size: 14px;
            font-weight: 750;
        }

        .attention-text {
            color: #5C6876;
            font-size: 13px;
            margin-top: 3px;
        }

        .small-note {
            color: #7B8794;
            font-size: 12px;
        }

        div[data-testid="stDataFrame"] {
            background-color: white;
            border-radius: 10px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# CONSTANTS
# =========================================================
DEFAULT_FILE = Path("YVF_Adoption_Dashboard_CS_HAD.xlsx")

SHEET_NAMES = {
    "booking": "Data_Booking",
    "si": "Data_SI",
    "customer": "Data_CustomerActive",
    "enhancement": "Data_Issue",
    "feedback": "Customer_Feedback",
}


# =========================================================
# HELPERS
# =========================================================
def normalize_text(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def safe_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def format_number(value: float) -> str:
    return f"{value:,.0f}"


def format_percent(value: float) -> str:
    return f"{value:.1%}"


def show_title(title: str, subtitle: str) -> None:
    st.markdown(f'<div class="dashboard-title">{title}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="dashboard-subtitle">{subtitle}</div>',
        unsafe_allow_html=True,
    )


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


def attention_box(title: str, text: str) -> None:
    st.markdown(
        f"""
        <div class="attention-box">
            <div class="attention-title">{title}</div>
            <div class="attention-text">{text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    excel = pd.ExcelFile(io.BytesIO(file_bytes))
    result: dict[str, pd.DataFrame] = {}

    for key, sheet_name in SHEET_NAMES.items():
        if sheet_name in excel.sheet_names:
            result[key] = pd.read_excel(excel, sheet_name=sheet_name)
        else:
            result[key] = pd.DataFrame()

    return result


def prepare_data(raw: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    booking = raw["booking"].copy()
    customer = raw["customer"].copy()
    enhancement = raw["enhancement"].copy()
    feedback = raw["feedback"].copy()
    si = raw["si"].copy()

    # Booking data
    if not booking.empty:
        if "Booking Date" in booking.columns:
            booking["Booking Date"] = pd.to_datetime(
                booking["Booking Date"], errors="coerce"
            )

        for col in ["Booking Qty", "Processing Time (Min)"]:
            if col in booking.columns:
                booking[col] = safe_numeric(booking[col])

        for col in [
            "Customer",
            "Mode",
            "Status",
            "System Issue",
            "PIC",
            "Remark",
        ]:
            if col in booking.columns:
                booking[col] = normalize_text(booking[col])

        if "Customer" in booking.columns:
            booking = booking[booking["Customer"].ne("")].copy()

    # Customer adoption data
    if not customer.empty:
        if "Start Using YVF" in customer.columns:
            customer["Start Using YVF"] = pd.to_datetime(
                customer["Start Using YVF"], errors="coerce"
            )

        for col in [
            "Target Booking/Year",
            "Actual Booking YTD",
            "Achievement %",
            "Actual This Month",
        ]:
            if col in customer.columns:
                customer[col] = safe_numeric(customer[col])

        for col in ["Customer", "Customer Code", "Status"]:
            if col in customer.columns:
                customer[col] = normalize_text(customer[col])

        if "Customer" in customer.columns:
            customer = customer[customer["Customer"].ne("")].copy()

    # Enhancement data
    if not enhancement.empty:
        enhancement = enhancement.dropna(how="all").copy()
        for col in enhancement.columns:
            if enhancement[col].dtype == "object":
                enhancement[col] = normalize_text(enhancement[col])

    # Feedback data
    if not feedback.empty:
        feedback = feedback.dropna(how="all").copy()
        for col in feedback.columns:
            if feedback[col].dtype == "object":
                feedback[col] = normalize_text(feedback[col])

    return {
        "booking": booking,
        "customer": customer,
        "enhancement": enhancement,
        "feedback": feedback,
        "si": si,
    }


def get_file_bytes() -> bytes | None:
    uploaded_file = st.sidebar.file_uploader(
        "Upload dữ liệu Excel",
        type=["xlsx"],
        help="File cần có các sheet Data_Booking, Data_CustomerActive, Data_Issue và Customer_Feedback.",
    )

    if uploaded_file is not None:
        return uploaded_file.getvalue()

    if DEFAULT_FILE.exists():
        return DEFAULT_FILE.read_bytes()

    return None


# =========================================================
# LOAD DATA
# =========================================================
file_bytes = get_file_bytes()

if file_bytes is None:
    st.error(
        "Không tìm thấy file dữ liệu. Hãy upload file Excel ở thanh bên trái "
        "hoặc đặt file YVF_Adoption_Dashboard_CS_HAD.xlsx cùng thư mục với app.py."
    )
    st.stop()

try:
    data = prepare_data(load_excel(file_bytes))
except Exception as exc:
    st.error(f"Không thể đọc file Excel: {exc}")
    st.stop()

booking = data["booking"]
customer = data["customer"]
enhancement = data["enhancement"]
feedback = data["feedback"]


# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.markdown("## YVF Dashboard")
page = st.sidebar.radio(
    "Chọn nội dung",
    [
        "1. Overview",
        "2. Customer Adoption",
        "3. Booking Details",
        "4. User Issues",
        "5. Enhancement Requests",
    ],
)

st.sidebar.markdown("---")
st.sidebar.caption("Dashboard nội bộ CS HAD – Không sử dụng logo công ty.")


# =========================================================
# COMMON METRICS
# =========================================================
total_customers = len(customer)

active_customers = (
    int(customer["Status"].str.lower().eq("active").sum())
    if not customer.empty and "Status" in customer.columns
    else 0
)

adoption_rate = active_customers / total_customers if total_customers else 0

year_target = (
    float(customer["Target Booking/Year"].sum())
    if "Target Booking/Year" in customer.columns
    else 0
)

actual_ytd = (
    float(customer["Actual Booking YTD"].sum())
    if "Actual Booking YTD" in customer.columns
    else (
        float(booking["Booking Qty"].sum())
        if "Booking Qty" in booking.columns
        else 0
    )
)

target_achievement = actual_ytd / year_target if year_target else 0

actual_this_month = (
    float(customer["Actual This Month"].sum())
    if "Actual This Month" in customer.columns
    else 0
)

avg_processing_time = (
    float(booking["Processing Time (Min)"].mean())
    if not booking.empty and "Processing Time (Min)" in booking.columns
    else 0
)

issue_rows = (
    booking[
        booking["System Issue"].str.lower().isin(
            ["yes", "y", "true", "1", "có", "co"]
        )
    ].copy()
    if not booking.empty and "System Issue" in booking.columns
    else pd.DataFrame()
)

issue_count = len(issue_rows)
enhancement_count = len(enhancement)


# =========================================================
# PAGE 1: OVERVIEW
# =========================================================
if page == "1. Overview":
    show_title(
        "YVF Adoption Dashboard – CS HAD",
        "Executive Overview.",
    )

    # Most important KPIs first
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        kpi_card(
            "Active Customers",
            format_number(active_customers),
            f"Trên tổng số {total_customers} khách hàng",
        )
    with col2:
        kpi_card(
            "Adoption Rate",
            format_percent(adoption_rate),
            "Tỷ lệ khách hàng đang sử dụng YVF",
        )
    with col3:
        kpi_card(
            "Booking Actual YTD",
            format_number(actual_ytd),
            f"Target năm: {format_number(year_target)}",
        )
    with col4:
        kpi_card(
            "Target Achievement",
            format_percent(target_achievement),
            "Actual YTD / Target năm",
        )

    st.write("")

    col5, col6, col7, col8 = st.columns(4)
    with col5:
        kpi_card(
            "Booking This Month",
            format_number(actual_this_month),
            "Sản lượng tháng hiện tại",
        )
    with col6:
        kpi_card(
            "Avg. Processing Time",
            f"{avg_processing_time:.1f} min",
            "Thời gian xử lý booking trung bình",
        )
    with col7:
        kpi_card(
            "User Issue Reports",
            format_number(issue_count),
            "Issue ghi nhận từ dữ liệu booking",
        )
    with col8:
        kpi_card(
            "Enhancement Requests",
            format_number(enhancement_count),
            "Đề xuất cải tiến đang được theo dõi",
        )

    show_section("Actual vs. Target và Booking Volume")

    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        if not customer.empty and "Customer Code" in customer.columns:
            chart_data = customer.copy()
            chart_data["Target"] = chart_data.get(
                "Target Booking/Year", pd.Series(0, index=chart_data.index)
            )
            chart_data["Actual YTD"] = chart_data.get(
                "Actual Booking YTD", pd.Series(0, index=chart_data.index)
            )

            fig = px.bar(
                chart_data,
                x="Customer Code",
                y=["Target", "Actual YTD"],
                barmode="group",
                title="Actual YTD vs. Annual Target",
                labels={"value": "Booking Volume", "variable": ""},
            )
            fig.update_layout(
                height=360,
                margin=dict(l=10, r=10, t=55, b=10),
                plot_bgcolor="white",
                paper_bgcolor="white",
                legend_orientation="h",
                legend_y=-0.2,
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Chưa có dữ liệu Customer Adoption để vẽ biểu đồ.")

    with chart_col2:
        if not booking.empty and {"Customer", "Booking Qty"}.issubset(booking.columns):
            booking_by_customer = (
                booking.groupby("Customer", as_index=False)["Booking Qty"]
                .sum()
                .sort_values("Booking Qty", ascending=True)
            )

            fig = px.bar(
                booking_by_customer,
                x="Booking Qty",
                y="Customer",
                orientation="h",
                title="Booking Volume by Customer",
                labels={"Booking Qty": "Booking Volume", "Customer": ""},
            )
            fig.update_layout(
                height=360,
                margin=dict(l=10, r=10, t=55, b=10),
                plot_bgcolor="white",
                paper_bgcolor="white",
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Chưa có dữ liệu Booking để vẽ biểu đồ.")

    show_section("Management Attention")

    no_volume_customers = 0
    if not customer.empty and {
        "Status",
        "Actual Booking YTD",
    }.issubset(customer.columns):
        no_volume_customers = len(
            customer[
                customer["Status"].str.lower().eq("active")
                & customer["Actual Booking YTD"].eq(0)
            ]
        )

    if year_target > 0 and target_achievement < 0.8:
        attention_box(
            "1. Target achievement cần theo dõi",
            f"Actual YTD hiện đạt {format_percent(target_achievement)} target năm. "
            "Cần kiểm tra lại tiến độ rollout và mức target của từng khách hàng.",
        )
    else:
        attention_box(
            "1. Target achievement",
            f"Actual YTD hiện đạt {format_percent(target_achievement)} target năm.",
        )

    if issue_count > 0:
        attention_box(
            "2. System performance / User issues",
            f"Đang có {issue_count} lượt ghi nhận issue. "
            f"Thời gian xử lý trung bình hiện là {avg_processing_time:.1f} phút.",
        )
    else:
        attention_box(
            "2. System performance / User issues",
            "Chưa ghi nhận issue từ dữ liệu booking.",
        )

    attention_box(
        "3. Active customers chưa có booking",
        f"Có {no_volume_customers} khách hàng đang ở trạng thái Active "
        "nhưng chưa ghi nhận booking YVF.",
    )


# =========================================================
# PAGE 2: CUSTOMER ADOPTION
# =========================================================
elif page == "2. Customer Adoption":
    show_title(
        "Customer Adoption",
        "Theo dõi khách hàng active, adoption rate và tiến độ Actual so với Target.",
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        kpi_card("Total Customers", format_number(total_customers))
    with col2:
        kpi_card("Active Customers", format_number(active_customers))
    with col3:
        kpi_card("Adoption Rate", format_percent(adoption_rate))
    with col4:
        kpi_card("Target Achievement", format_percent(target_achievement))

    show_section("Customer Adoption Detail")

    if customer.empty:
        st.info("Chưa có dữ liệu trong sheet Data_CustomerActive.")
    else:
        display_customer = customer.copy()

        preferred_cols = [
            "Customer",
            "Customer Code",
            "Start Using YVF",
            "Status",
            "Target Booking/Year",
            "Actual Booking YTD",
            "Achievement %",
            "Actual This Month",
        ]
        available_cols = [c for c in preferred_cols if c in display_customer.columns]

        if {
            "Target Booking/Year",
            "Actual Booking YTD",
        }.issubset(display_customer.columns):
            display_customer["Achievement %"] = (
                display_customer["Actual Booking YTD"]
                / display_customer["Target Booking/Year"].replace(0, pd.NA)
            ).fillna(0)

        st.dataframe(
            display_customer[available_cols],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Start Using YVF": st.column_config.DateColumn(
                    "Start Using YVF", format="DD-MMM-YYYY"
                ),
                "Target Booking/Year": st.column_config.NumberColumn(
                    "Target Booking/Year", format="%,.0f"
                ),
                "Actual Booking YTD": st.column_config.NumberColumn(
                    "Actual Booking YTD", format="%,.0f"
                ),
                "Actual This Month": st.column_config.NumberColumn(
                    "Actual This Month", format="%,.0f"
                ),
                "Achievement %": st.column_config.ProgressColumn(
                    "Achievement %",
                    format="%.1f%%",
                    min_value=0,
                    max_value=1,
                ),
            },
        )

        if "Customer Code" in customer.columns:
            plot_df = customer.copy()
            plot_df["Target"] = plot_df.get(
                "Target Booking/Year", pd.Series(0, index=plot_df.index)
            )
            plot_df["Actual YTD"] = plot_df.get(
                "Actual Booking YTD", pd.Series(0, index=plot_df.index)
            )

            fig = px.bar(
                plot_df,
                x="Customer Code",
                y=["Target", "Actual YTD"],
                barmode="group",
                title="Target vs. Actual YTD by Customer",
                labels={"value": "Booking Volume", "variable": ""},
            )
            fig.update_layout(
                height=420,
                plot_bgcolor="white",
                paper_bgcolor="white",
                legend_orientation="h",
                legend_y=-0.2,
            )
            st.plotly_chart(fig, use_container_width=True)


# =========================================================
# PAGE 3: BOOKING DETAILS
# =========================================================
elif page == "3. Booking Details":
    show_title(
        "Booking Details",
        "Theo dõi booking volume, top/bottom customers, mode và thời gian xử lý.",
    )

    total_booking = (
        float(booking["Booking Qty"].sum())
        if not booking.empty and "Booking Qty" in booking.columns
        else 0
    )
    sea_booking = (
        float(
            booking.loc[
                booking["Mode"].str.lower().eq("sea"), "Booking Qty"
            ].sum()
        )
        if not booking.empty and {"Mode", "Booking Qty"}.issubset(booking.columns)
        else 0
    )
    air_booking = (
        float(
            booking.loc[
                booking["Mode"].str.lower().eq("air"), "Booking Qty"
            ].sum()
        )
        if not booking.empty and {"Mode", "Booking Qty"}.issubset(booking.columns)
        else 0
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        kpi_card("Total Booking Volume", format_number(total_booking))
    with col2:
        kpi_card("Sea Booking Volume", format_number(sea_booking))
    with col3:
        kpi_card("Air Booking Volume", format_number(air_booking))
    with col4:
        kpi_card("Avg. Processing Time", f"{avg_processing_time:.1f} min")

    if booking.empty:
        st.info("Chưa có dữ liệu trong sheet Data_Booking.")
    else:
        filter_col1, filter_col2 = st.columns(2)

        customer_options = sorted(
            booking["Customer"].dropna().astype(str).unique().tolist()
        )
        mode_options = sorted(
            booking["Mode"].dropna().astype(str).unique().tolist()
        )

        with filter_col1:
            selected_customers = st.multiselect(
                "Customer",
                customer_options,
                default=customer_options,
            )
        with filter_col2:
            selected_modes = st.multiselect(
                "Mode",
                mode_options,
                default=mode_options,
            )

        filtered_booking = booking[
            booking["Customer"].isin(selected_customers)
            & booking["Mode"].isin(selected_modes)
        ].copy()

        show_section("Booking Volume by Customer")

        customer_summary = (
            filtered_booking.groupby("Customer", as_index=False)["Booking Qty"]
            .sum()
            .sort_values("Booking Qty", ascending=False)
        )

        col_chart, col_rank = st.columns([2, 1])

        with col_chart:
            fig = px.bar(
                customer_summary.sort_values("Booking Qty"),
                x="Booking Qty",
                y="Customer",
                orientation="h",
                title="Booking Volume by Customer",
            )
            fig.update_layout(
                height=420,
                plot_bgcolor="white",
                paper_bgcolor="white",
            )
            st.plotly_chart(fig, use_container_width=True)

        with col_rank:
            st.markdown("#### Top Customers")
            st.dataframe(
                customer_summary.head(5),
                use_container_width=True,
                hide_index=True,
            )

            st.markdown("#### Bottom Customers")
            st.dataframe(
                customer_summary.tail(5).sort_values("Booking Qty"),
                use_container_width=True,
                hide_index=True,
            )

        show_section("Booking Data")
        st.dataframe(
            filtered_booking,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Booking Date": st.column_config.DateColumn(
                    "Booking Date", format="DD-MMM-YYYY"
                ),
                "Booking Qty": st.column_config.NumberColumn(
                    "Booking Qty", format="%,.0f"
                ),
                "Processing Time (Min)": st.column_config.NumberColumn(
                    "Processing Time (Min)", format="%.1f"
                ),
            },
        )


# =========================================================
# PAGE 4: USER ISSUES
# =========================================================
elif page == "4. User Issues":
    show_title(
        "User Issues",
        "Các vấn đề được phản ánh bởi Customers hoặc CS trong quá trình sử dụng YVF.",
    )

    negative_feedback = (
        int(feedback["Type"].str.lower().eq("negative").sum())
        if not feedback.empty and "Type" in feedback.columns
        else 0
    )

    affected_customers = (
        issue_rows["Customer"].nunique()
        if not issue_rows.empty and "Customer" in issue_rows.columns
        else 0
    )

    issue_avg_time = (
        float(issue_rows["Processing Time (Min)"].mean())
        if not issue_rows.empty and "Processing Time (Min)" in issue_rows.columns
        else 0
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        kpi_card("Issue Reports", format_number(issue_count))
    with col2:
        kpi_card("Affected Customers", format_number(affected_customers))
    with col3:
        kpi_card("Avg. Time on Issue Rows", f"{issue_avg_time:.1f} min")
    with col4:
        kpi_card("Negative Feedback", format_number(negative_feedback))

    show_section("Issue Log")

    if issue_rows.empty:
        st.info("Chưa có issue được đánh dấu trong sheet Data_Booking.")
    else:
        issue_columns = [
            col
            for col in [
                "Booking Date",
                "Customer",
                "PIC",
                "Mode",
                "Status",
                "System Issue",
                "Processing Time (Min)",
                "Remark",
            ]
            if col in issue_rows.columns
        ]

        st.dataframe(
            issue_rows[issue_columns],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Booking Date": st.column_config.DateColumn(
                    "Booking Date", format="DD-MMM-YYYY"
                ),
                "Processing Time (Min)": st.column_config.NumberColumn(
                    "Processing Time (Min)", format="%.1f"
                ),
            },
        )

    if not feedback.empty:
        show_section("Customer Feedback")
        st.dataframe(
            feedback,
            use_container_width=True,
            hide_index=True,
        )


# =========================================================
# PAGE 5: ENHANCEMENT REQUESTS
# =========================================================
else:
    show_title(
        "Enhancement Requests",
        "Danh sách đề xuất cải tiến từ Customers hoặc CS.",
    )

    booking_module = (
        int(enhancement["Module"].str.contains("Booking", case=False, na=False).sum())
        if not enhancement.empty and "Module" in enhancement.columns
        else 0
    )
    si_module = (
        int(enhancement["Module"].str.contains("SI", case=False, na=False).sum())
        if not enhancement.empty and "Module" in enhancement.columns
        else 0
    )
    vgm_module = (
        int(enhancement["Module"].str.contains("VGM", case=False, na=False).sum())
        if not enhancement.empty and "Module" in enhancement.columns
        else 0
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        kpi_card("Total Requests", format_number(enhancement_count))
    with col2:
        kpi_card("Booking Module", format_number(booking_module))
    with col3:
        kpi_card("SI Module", format_number(si_module))
    with col4:
        kpi_card("VGM Module", format_number(vgm_module))

    show_section("Enhancement Request Register")

    if enhancement.empty:
        st.info("Chưa có dữ liệu trong sheet Data_Issue.")
    else:
        display_enhancement = enhancement.copy()

        rename_map = {
            "Issue ID": "Request ID",
            "Issue": "Current Limitation",
            "Impact": "Business Impact",
            "Suggestion": "Requested Enhancement",
        }
        display_enhancement = display_enhancement.rename(columns=rename_map)

        st.dataframe(
            display_enhancement,
            use_container_width=True,
            hide_index=True,
        )
