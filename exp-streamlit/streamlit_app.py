import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import altair as alt

# Set page config
st.set_page_config(
    page_title="Project Management Dashboard",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
        .main {
            padding: 0rem 1rem;
        }
        .stMetric {
            background-color: #f0f2f6;
            padding: 1rem;
            border-radius: 0.5rem;
        }
    </style>
""", unsafe_allow_html=True)

# Load Excel file
@st.cache_data
def load_excel_data():
    excel_file = "Project 1_Project_MIS_070120252.xlsx"
    
    # Load relevant sheets
    bplan_tracker = pd.read_excel(excel_file, sheet_name="BPlan Tracker")
    monthly_sales = pd.read_excel(excel_file, sheet_name="Monthly Sale Tracker ")
    bank_balances = pd.read_excel(excel_file, sheet_name="Bank Balances")
    cashflow = pd.read_excel(excel_file, sheet_name="Cashflow Process")
    admin_expenses = pd.read_excel(excel_file, sheet_name="Admin Expense Tracker ")
    
    return {
        "bplan": bplan_tracker,
        "sales": monthly_sales,
        "bank": bank_balances,
        "cashflow": cashflow,
        "expenses": admin_expenses
    }

try:
    data = load_excel_data()
    
    # Header
    st.title("📊 Project Management Dashboard")
    st.markdown("---")

    # Sidebar
    st.sidebar.header("Dashboard Controls")
    selected_view = st.sidebar.selectbox(
        "Select View",
        ["Overview", "Financial Analysis", "Sales Tracker", "Expense Analysis"]
    )

    # Overview Section
    if selected_view == "Overview":
        # Key Metrics Row
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                label="Total Bank Balance",
                value=f"₹{data['bank']['Balance'].sum():,.2f}",
                delta="4.5%"
            )
            
        with col2:
            st.metric(
                label="Monthly Sales",
                value=f"₹{data['sales']['Total Sales'].sum():,.2f}",
                delta="2.8%"
            )
            
        with col3:
            st.metric(
                label="Expense Utilization",
                value="87%",
                delta="-2.1%",
                delta_color="inverse"
            )
            
        with col4:
            st.metric(
                label="Project Completion",
                value="73%",
                delta="5.2%"
            )

        # Cashflow Analysis
        st.subheader("Cashflow Analysis")
        cashflow_data = data['cashflow'].melt(
            id_vars=['Date'],
            value_vars=['Inflow', 'Outflow'],
            var_name='Flow Type',
            value_name='Amount'
        )
        
        cashflow_chart = px.line(
            cashflow_data,
            x='Date',
            y='Amount',
            color='Flow Type',
            title='Cash Inflow vs Outflow Trend'
        )
        st.plotly_chart(cashflow_chart, use_container_width=True)

    # Financial Analysis Section
    elif selected_view == "Financial Analysis":
        st.header("Financial Analysis")
        
        # Bank Balance Distribution
        st.subheader("Bank Balance Distribution")
        fig_bank = px.pie(
            data['bank'],
            values='Balance',
            names='Account Type',
            title='Distribution of Funds Across Accounts'
        )
        st.plotly_chart(fig_bank, use_container_width=True)
        
        # Budget vs Actual Comparison
        st.subheader("Budget vs Actual Expenditure")
        budget_comparison = go.Figure()
        budget_comparison.add_trace(go.Bar(
            name='Budgeted',
            x=data['bplan']['Category'],
            y=data['bplan']['Budgeted Amount'],
        ))
        budget_comparison.add_trace(go.Bar(
            name='Actual',
            x=data['bplan']['Category'],
            y=data['bplan']['Actual Amount'],
        ))
        budget_comparison.update_layout(barmode='group')
        st.plotly_chart(budget_comparison, use_container_width=True)

    # Sales Tracker Section
    elif selected_view == "Sales Tracker":
        st.header("Sales Analysis")
        
        # Monthly Sales Trend
        sales_trend = px.line(
            data['sales'],
            x='Month',
            y='Total Sales',
            title='Monthly Sales Trend'
        )
        st.plotly_chart(sales_trend, use_container_width=True)
        
        # Sales by Category
        col1, col2 = st.columns(2)
        
        with col1:
            sales_category = px.pie(
                data['sales'],
                values='Amount',
                names='Category',
                title='Sales Distribution by Category'
            )
            st.plotly_chart(sales_category)
            
        with col2:
            sales_location = px.bar(
                data['sales'],
                x='Location',
                y='Amount',
                title='Sales by Location'
            )
            st.plotly_chart(sales_location)

    # Expense Analysis Section
    elif selected_view == "Expense Analysis":
        st.header("Expense Analysis")
        
        # Monthly Expense Trend
        expense_trend = px.line(
            data['expenses'],
            x='Month',
            y='Total Expense',
            title='Monthly Expense Trend'
        )
        st.plotly_chart(expense_trend, use_container_width=True)
        
        # Expense Categories
        expense_categories = px.bar(
            data['expenses'],
            x='Category',
            y='Amount',
            title='Expenses by Category',
            color='Category'
        )
        st.plotly_chart(expense_categories, use_container_width=True)

except Exception as e:
    st.error(f"An error occurred: {str(e)}")
    st.info("Please ensure the Excel file is in the correct format and all required sheets are present.")

# Footer
st.markdown("---")
st.markdown("Dashboard last updated: January 21, 2025")
