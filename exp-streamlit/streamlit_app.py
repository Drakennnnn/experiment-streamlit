import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import altair as alt
import io

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
        .upload-section {
            padding: 2rem;
            border-radius: 0.5rem;
            border: 2px dashed #4a90e2;
            text-align: center;
            margin-bottom: 2rem;
        }
    </style>
""", unsafe_allow_html=True)

def load_excel_data(uploaded_file):
    """Load and process Excel data from uploaded file"""
    try:
        # Create a dictionary to store DataFrames
        data_dict = {}
        
        # Read required sheets
        required_sheets = [
            "BPlan Tracker",
            "Monthly Sale Tracker ",
            "Bank Balances",
            "Cashflow Process",
            "Admin Expense Tracker "
        ]
        
        # Read all sheets at once
        with pd.ExcelFile(uploaded_file) as xls:
            available_sheets = xls.sheet_names
            
            # Check if all required sheets are present
            missing_sheets = [sheet for sheet in required_sheets if sheet not in available_sheets]
            if missing_sheets:
                st.error(f"Missing required sheets: {', '.join(missing_sheets)}")
                return None
            
            # Load each required sheet
            for sheet_name in required_sheets:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                # Basic data cleaning
                df = df.replace([pd.NA, float('inf'), float('-inf')], pd.NA)
                data_dict[sheet_name.strip().lower().replace(' ', '_')] = df
        
        return {
            "bplan": data_dict["bplan_tracker"],
            "sales": data_dict["monthly_sale_tracker"],
            "bank": data_dict["bank_balances"],
            "cashflow": data_dict["cashflow_process"],
            "expenses": data_dict["admin_expense_tracker"]
        }
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return None

def render_overview(data):
    """Render the overview section"""
    # Key Metrics Row
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        try:
            total_balance = data['bank']['Balance'].sum() if 'Balance' in data['bank'].columns else 0
            st.metric(
                label="Total Bank Balance",
                value=f"₹{total_balance:,.2f}",
                delta="4.5%"
            )
        except:
            st.metric(label="Total Bank Balance", value="N/A")
    
    with col2:
        try:
            total_sales = data['sales']['Total Sales'].sum() if 'Total Sales' in data['sales'].columns else 0
            st.metric(
                label="Monthly Sales",
                value=f"₹{total_sales:,.2f}",
                delta="2.8%"
            )
        except:
            st.metric(label="Monthly Sales", value="N/A")
    
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
    try:
        if 'Inflow' in data['cashflow'].columns and 'Outflow' in data['cashflow'].columns:
            cashflow_data = data['cashflow'].melt(
                id_vars=['Date'] if 'Date' in data['cashflow'].columns else [],
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
        else:
            st.warning("Required columns not found in cashflow data")
    except Exception as e:
        st.error(f"Error rendering cashflow analysis: {str(e)}")

def render_financial_analysis(data):
    """Render the financial analysis section"""
    st.header("Financial Analysis")
    
    try:
        # Bank Balance Distribution
        if 'Balance' in data['bank'].columns and 'Account Type' in data['bank'].columns:
            st.subheader("Bank Balance Distribution")
            fig_bank = px.pie(
                data['bank'],
                values='Balance',
                names='Account Type',
                title='Distribution of Funds Across Accounts'
            )
            st.plotly_chart(fig_bank, use_container_width=True)
        else:
            st.warning("Required columns not found in bank balance data")
        
        # Budget vs Actual Comparison
        if all(col in data['bplan'].columns for col in ['Category', 'Budgeted Amount', 'Actual Amount']):
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
        else:
            st.warning("Required columns not found in budget data")
    except Exception as e:
        st.error(f"Error rendering financial analysis: {str(e)}")

def render_sales_tracker(data):
    """Render the sales tracker section"""
    st.header("Sales Analysis")
    
    try:
        # Monthly Sales Trend
        if 'Month' in data['sales'].columns and 'Total Sales' in data['sales'].columns:
            sales_trend = px.line(
                data['sales'],
                x='Month',
                y='Total Sales',
                title='Monthly Sales Trend'
            )
            st.plotly_chart(sales_trend, use_container_width=True)
        else:
            st.warning("Required columns not found for sales trend")
        
        # Sales by Category
        col1, col2 = st.columns(2)
        
        with col1:
            if 'Amount' in data['sales'].columns and 'Category' in data['sales'].columns:
                sales_category = px.pie(
                    data['sales'],
                    values='Amount',
                    names='Category',
                    title='Sales Distribution by Category'
                )
                st.plotly_chart(sales_category)
            else:
                st.warning("Required columns not found for sales categories")
        
        with col2:
            if 'Amount' in data['sales'].columns and 'Location' in data['sales'].columns:
                sales_location = px.bar(
                    data['sales'],
                    x='Location',
                    y='Amount',
                    title='Sales by Location'
                )
                st.plotly_chart(sales_location)
            else:
                st.warning("Required columns not found for sales location")
    except Exception as e:
        st.error(f"Error rendering sales tracker: {str(e)}")

def render_expense_analysis(data):
    """Render the expense analysis section"""
    st.header("Expense Analysis")
    
    try:
        # Monthly Expense Trend
        if 'Month' in data['expenses'].columns and 'Total Expense' in data['expenses'].columns:
            expense_trend = px.line(
                data['expenses'],
                x='Month',
                y='Total Expense',
                title='Monthly Expense Trend'
            )
            st.plotly_chart(expense_trend, use_container_width=True)
        else:
            st.warning("Required columns not found for expense trend")
        
        # Expense Categories
        if 'Category' in data['expenses'].columns and 'Amount' in data['expenses'].columns:
            expense_categories = px.bar(
                data['expenses'],
                x='Category',
                y='Amount',
                title='Expenses by Category',
                color='Category'
            )
            st.plotly_chart(expense_categories, use_container_width=True)
        else:
            st.warning("Required columns not found for expense categories")
    except Exception as e:
        st.error(f"Error rendering expense analysis: {str(e)}")

def main():
    st.title("📊 Project Management Dashboard")
    st.markdown("---")

    # File Upload Section
    st.markdown("""
        <div class="upload-section">
            <h3>Upload Project Excel File</h3>
            <p>Please upload your project management Excel file to view the dashboard.</p>
        </div>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Show loading spinner while processing data
        with st.spinner('Loading and processing data...'):
            data = load_excel_data(uploaded_file)
        
        if data is not None:
            # Sidebar
            st.sidebar.header("Dashboard Controls")
            selected_view = st.sidebar.selectbox(
                "Select View",
                ["Overview", "Financial Analysis", "Sales Tracker", "Expense Analysis"]
            )
            
            # Render selected view
            if selected_view == "Overview":
                render_overview(data)
            elif selected_view == "Financial Analysis":
                render_financial_analysis(data)
            elif selected_view == "Sales Tracker":
                render_sales_tracker(data)
            elif selected_view == "Expense Analysis":
                render_expense_analysis(data)
            
            # Footer
            st.markdown("---")
            st.markdown(f"Dashboard last updated: {datetime.now().strftime('%B %d, %Y %H:%M:%S')}")
    else:
        st.info("👆 Upload an Excel file to get started")

if __name__ == "__main__":
    main()
