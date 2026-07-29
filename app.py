import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# ---------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------
st.set_page_config(
    page_title="YVF Adoption Dashboard - CS HAD",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS cho giao diện
st.markdown("""
<style>
    .main-header { font-size: 26px; font-weight: bold; color: #1E3A8A; margin-bottom: 20px; }
    .metric-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# DATA LOADING & PREPROCESSING
# ---------------------------------------------------------
@st.cache_data
def load_data():
    # Tự động tìm file Excel ở thư mục hiện tại hoặc trong thư mục 'data'
    file_name = "YVF Booking & On board status.xlsx"
    if os.path.exists(file_name):
        file_path = file_name
    elif os.path.exists(os.path.join("data", file_name)):
        file_path = os.path.join("data", file_name)
    else:
        st.error(f"❌ Không tìm thấy file dữ liệu '{file_name}'. Vui lòng kiểm tra lại thư mục lưu trữ!")
        st.stop()

    # 1. Sheet: YVF Status
    df_status = pd.read_excel(file_path, sheet_name="YVF Status", header=None)
    
    eligible = int(df_status.iloc[3, 0]) if pd.notnull(df_status.iloc[3, 0]) else 0
    total_hbl = int(df_status.iloc[3, 1]) if pd.notnull(df_status.iloc[3, 1]) else 0
    onboarded = int(df_status.iloc[3, 2]) if pd.notnull(df_status.iloc[3, 2]) else 0
    pending_onboard = int(df_status.iloc[3, 4]) if pd.notnull(df_status.iloc[3, 4]) else 0
    active_customers = int(df_status.iloc[3, 5]) if pd.notnull(df_status.iloc[3, 5]) else 0
    total_bookings = int(df_status.iloc[3, 6]) if pd.notnull(df_status.iloc[3, 6]) else 0
    avg_booking_time = float(df_status.iloc[3, 7]) if pd.notnull(df_status.iloc[3, 7]) else 0.0
    booking_target = int(df_status.iloc[3, 9]) if pd.notnull(df_status.iloc[3, 9]) else 0
    
    kpis = {
        "eligible": eligible,
        "total_hbl": total_hbl,
        "onboarded": onboarded,
        "pending_onboard": pending_onboard,
        "active_customers": active_customers,
        "adoption_rate": (active_customers / onboarded) if onboarded > 0 else 0.0,
        "onboarding_rate": (onboarded / eligible) if eligible > 0 else 0.0,
        "total_bookings": total_bookings,
        "avg_booking_time": avg_booking_time,
        "booking_target": booking_target,
        "booking_achievement": (total_bookings / booking_target) if booking_target > 0 else 0.0
    }
    
    # 2. Sheet: Target YVF customer
    df_target = pd.read_excel(file_path, sheet_name="Target YVF customer", skiprows=2)
    df_target = df_target.dropna(subset=['Customer']).copy()
    
    # Tính tổng volume hàng tháng từ các shipment modes (AE, FCL, LCL)
    ae_apr = df_target.iloc[:, 2].fillna(0)
    ae_may = df_target.iloc[:, 3].fillna(0)
    ae_jun = df_target.iloc[:, 4].fillna(0)
    
    fcl_apr = df_target.iloc[:, 5].fillna(0)
    fcl_may = df_target.iloc[:, 6].fillna(0)
    fcl_jun = df_target.iloc[:, 7].fillna(0)
    
    lcl_apr = df_target.iloc[:, 8].fillna(0)
    lcl_may = df_target.iloc[:, 9].fillna(0)
    lcl_jun = df_target.iloc[:, 10].fillna(0)
    
    df_target['Apr_Vol'] = ae_apr + fcl_apr + lcl_apr
    df_target['May_Vol'] = ae_may + fcl_may + lcl_may
    df_target['Jun_Vol'] = ae_jun + fcl_jun + lcl_jun
    df_target['Total_Volume'] = df_target.iloc[:, 11].fillna(df_target['Apr_Vol'] + df_target['May_Vol'] + df_target['Jun_Vol'])
    df_target['Status'] = df_target.iloc[:, 12].fillna('Not yet')
    
    # 3. Sheet: Onboard customer list
    df_onboard = pd.read_excel(file_path, sheet_name="Onboard customer list", skiprows=1)
    df_onboard = df_onboard.dropna(subset=['Shipper Company Name']).copy()
    
    def classify_status(row):
        booking_status = str(row['Booking status\n(number of booking on YVF per month or per year)'])
        if '100%' in booking_status:
            return 'Active'
        elif 'trial' in booking_status.lower():
            return 'Trial'
        elif 'Not booking' in booking_status:
            return 'Pending'
        return 'Inactive'
        
    df_onboard['Category_Status'] = df_onboard.apply(classify_status, axis=1)

    # 4. Standardized User Issues Data
    issues_data = [
        {"Issue": "Slow response", "Description": "VYF sometimes responds slowly, impacting booking process", "Category": "Performance", "Status": "Open"},
        {"Issue": "Login failure", "Description": "Occasional timeout during login", "Category": "Authentication", "Status": "Closed"},
        {"Issue": "Additional Email", "Description": "Notifications not delivered to secondary email", "Category": "Notification", "Status": "Pending"}
    ]
    df_issues = pd.DataFrame(issues_data)

    # 5. Standardized Improvement Proposals Data
    proposals_data = [
        {"Proposal": "Additional Email Notification", "Category": "Notification", "Status": "Pending"},
        {"Proposal": "Hide Verification Code", "Category": "Security", "Status": "Pending"},
        {"Proposal": "Tracking by Carrier Booking No.", "Category": "Tracking", "Status": "Proposed"},
        {"Proposal": "Separate CBM & GRW", "Category": "Feature", "Status": "Proposed"},
        {"Proposal": "Submit VGM via YVF", "Category": "Feature", "Status": "Proposed"}
    ]
    df_proposals = pd.DataFrame(proposals_data)

    return kpis, df_target, df_onboard, df_issues, df_proposals

# Load Data
kpis, df_target, df_onboard, df_issues, df_proposals = load_data()

# ---------------------------------------------------------
# SIDEBAR NAVIGATION
# ---------------------------------------------------------
st.sidebar.title("🚢 YVF Navigation")
menu = st.sidebar.radio(
    "Danh mục Dashboard",
    [
        "1. Overview",
        "2. Customer Adoption",
        "3. Booking Performance",
        "4. User Issues",
        "5. Improvement Proposals",
        "6. Customer Details"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("YVF Adoption Dashboard | CS HAD Division")

# ---------------------------------------------------------
# 1. OVERVIEW PAGE
# ---------------------------------------------------------
if menu == "1. Overview":
    st.markdown('<div class="main-header">📌 Executive Summary & Overview</div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Eligible Customers", kpis["eligible"])
    col2.metric("Onboarded Customers", kpis["onboarded"])
    col3.metric("Active Customers", kpis["active_customers"])
    col4.metric("Adoption Rate", f"{kpis['adoption_rate']:.1%}")
    col5.metric("Booking Achievement", f"{kpis['booking_achievement']:.1%}")
        
    st.markdown("---")
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Pending Onboard", kpis["pending_onboard"])
    c2.metric("Onboarding Rate", f"{kpis['onboarding_rate']:.1%}")
    c3.metric("Total Bookings via YVF", kpis["total_bookings"])
    c4.metric("Avg Booking Time", f"{kpis['avg_booking_time']} min/bk")

    st.markdown("---")
    
    col_left, col_right = st.columns(2)
    with col_left:
        st.subheader("🎯 Customer Adoption Breakdown")
        fig_donut = px.pie(
            names=["Active", "Onboarded (Inactive)", "Not Onboarded"],
            values=[kpis["active_customers"], kpis["onboarded"] - kpis["active_customers"], kpis["eligible"] - kpis["onboarded"]],
            hole=0.5,
            color_discrete_sequence=["#10B981", "#F59E0B", "#EF4444"]
        )
        st.plotly_chart(fig_donut, use_container_width=True)
        
    with col_right:
        st.subheader("📈 Monthly Booking Target Achievement")
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=kpis["total_bookings"],
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': f"Actual ({kpis['total_bookings']}) vs Target ({kpis['booking_target']})"},
            gauge={
                'axis': {'range': [0, kpis['booking_target']]},
                'bar': {'color': "#2563EB"},
                'steps': [
                    {'range': [0, kpis['booking_target']*0.3], 'color': "#FEE2E2"},
                    {'range': [kpis['booking_target']*0.3, kpis['booking_target']*0.7], 'color': "#FEF3C7"},
                    {'range': [kpis['booking_target']*0.7, kpis['booking_target']], 'color': "#D1FAE5"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': kpis['booking_target']
                }
            }
        ))
        st.plotly_chart(fig_gauge, use_container_width=True)

# ---------------------------------------------------------
# 2. CUSTOMER ADOPTION PAGE
# ---------------------------------------------------------
elif menu == "2. Customer Adoption":
    st.markdown('<div class="main-header">👥 Customer Adoption Breakdown</div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Eligible Customers", kpis["eligible"])
    col2.metric("Onboarded Customers", kpis["onboarded"])
    col3.metric("Active Customers", kpis["active_customers"])
    col4.metric("Pending Onboard", kpis["pending_onboard"])
    col5.metric("Adoption Rate", f"{kpis['adoption_rate']:.1%}")

    st.markdown("---")
    
    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
        st.subheader("🍩 Onboarding Ratio")
        fig_donut = px.pie(
            names=["Onboarded", "Not Onboarded"],
            values=[kpis["onboarded"], kpis["eligible"] - kpis["onboarded"]],
            hole=0.5,
            color_discrete_sequence=["#3B82F6", "#9CA3AF"]
        )
        st.plotly_chart(fig_donut, use_container_width=True)
        
    with col_chart2:
        st.subheader("📊 Customer Status Distribution")
        status_counts = df_onboard['Category_Status'].value_counts().reset_index()
        status_counts.columns = ['Status', 'Count']
        fig_bar = px.bar(status_counts, x='Status', y='Count', color='Status', color_discrete_sequence=px.colors.qualitative.Set2)
        st.plotly_chart(fig_bar, use_container_width=True)

    st.subheader("📋 Onboarded Customer Tracking List")
    st.dataframe(
        df_onboard[['Shipper Company Name', 'Shipment Mode (FCL/LCL/AIR)', 'Booking status\n(number of booking on YVF per month or per year)', 'Category_Status', 'Remark']],
        use_container_width=True
    )

# ---------------------------------------------------------
# 3. BOOKING PERFORMANCE PAGE
# ---------------------------------------------------------
elif menu == "3. Booking Performance":
    st.markdown('<div class="main-header">🚀 Booking Performance & Trends</div>', unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Booking via YVF", kpis["total_bookings"])
    col2.metric("Target / Month", kpis["booking_target"])
    col3.metric("Achievement %", f"{kpis['booking_achievement']:.1%}")
    col4.metric("Avg Booking Time", f"{kpis['avg_booking_time']} Min")

    st.markdown("---")
    
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📅 Total Volume Trend by Month (Apr - Jun)")
        monthly_vol = pd.DataFrame({
            "Month": ["Apr", "May", "Jun"],
            "Volume": [df_target['Apr_Vol'].sum(), df_target['May_Vol'].sum(), df_target['Jun_Vol'].sum()]
        })
        fig_line = px.line(monthly_vol, x="Month", y="Volume", markers=True, title="Export Shipment Volume Trend")
        st.plotly_chart(fig_line, use_container_width=True)
        
    with c2:
        st.subheader("🏆 Top 10 Customers by Volume")
        top10 = df_target.sort_values(by="Total_Volume", ascending=False).head(10)
        fig_top = px.bar(top10, x="Total_Volume", y="Customer", orientation='h', color="Total_Volume", color_continuous_scale="Blues")
        fig_top.update_layout(yaxis={'categoryorder': 'total ascending'})
        st.plotly_chart(fig_top, use_container_width=True)

# ---------------------------------------------------------
# 4. YVF USER ISSUES PAGE
# ---------------------------------------------------------
elif menu == "4. User Issues":
    st.markdown('<div class="main-header">🛠️ User Issues & Support Log</div>', unsafe_allow_html=True)
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Issues", len(df_issues))
    c2.metric("Open Issues", len(df_issues[df_issues['Status'] == 'Open']))
    c3.metric("Pending Issues", len(df_issues[df_issues['Status'] == 'Pending']))
    c4.metric("Closed Issues", len(df_issues[df_issues['Status'] == 'Closed']))

    st.markdown("---")
    
    col_pie, col_bar = st.columns(2)
    with col_pie:
        st.subheader("🍩 Issues by Category")
        fig_cat = px.pie(df_issues, names="Category", hole=0.4)
        st.plotly_chart(fig_cat, use_container_width=True)
        
    with col_bar:
        st.subheader("📊 Issues Status Count")
        fig_stat = px.bar(df_issues['Status'].value_counts().reset_index(), x="Status", y="count", color="Status")
        st.plotly_chart(fig_stat, use_container_width=True)

    st.subheader("📄 Issue Details Table")
    st.dataframe(df_issues, use_container_width=True)

# ---------------------------------------------------------
# 5. IMPROVEMENT PROPOSALS PAGE
# ---------------------------------------------------------
elif menu == "5. Improvement Proposals":
    st.markdown('<div class="main-header">💡 System Improvement Proposals</div>', unsafe_allow_html=True)
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Proposals", len(df_proposals))
    c2.metric("Pending", len(df_proposals[df_proposals['Status'] == 'Pending']))
    c3.metric("Proposed", len(df_proposals[df_proposals['Status'] == 'Proposed']))
    c4.metric("Implemented", len(df_proposals[df_proposals['Status'] == 'Implemented']))

    st.markdown("---")
    
    st.subheader("📌 Top Requested Improvements")
    fig_prop = px.bar(df_proposals, x="Proposal", y="Category", color="Status", barmode="group")
    st.plotly_chart(fig_prop, use_container_width=True)

    st.subheader("📋 Detailed Proposals List")
    st.dataframe(df_proposals, use_container_width=True)

# ---------------------------------------------------------
# 6. CUSTOMER DETAILS PAGE
# ---------------------------------------------------------
elif menu == "6. Customer Details":
    st.markdown('<div class="main-header">🔍 Target Customer Master List</div>', unsafe_allow_html=True)
    
    search = st.text_input("🔍 Tìm kiếm tên khách hàng:", "")
    filtered_df = df_target if not search else df_target[df_target['Customer'].str.contains(search, case=False, na=False)]

    st.dataframe(
        filtered_df[['No', 'Customer', 'Apr_Vol', 'May_Vol', 'Jun_Vol', 'Total_Volume', 'Status']],
        use_container_width=True,
        height=500
    )
