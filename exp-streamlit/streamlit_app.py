import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np

# Page Configuration
st.set_page_config(
    page_title="Project Management Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for dark theme and better visibility
st.markdown("""
    <style>
        .main {
            background-color: #0E1117;
            color: #FAFAFA;
        }
        .stPlotlyChart {
            background-color: #262730;
            border-radius: 0.5rem;
            padding: 1rem;
        }
        .css-1rs6os {
            background-color: #0E1117;
            color: #FAFAFA;
        }
        .stDataFrame {
            background-color: #262730;
        }
        .stMetric {
            background-color: #262730;
            padding: 1rem;
            border-radius: 0.5rem;
        }
        div[data-testid="stMetricValue"] {
            color: #FAFAFA;
            font-size: 1.8rem !important;
        }
        .uploadSection {
            background-color: #262730;
            padding: 2rem;
            border-radius: 0.5rem;
            border: 2px dashed #4A90E2;
            text-align: center;
            margin-bottom: 2rem;
        }
        .custom-info-box {
            background-color: #262730;
            padding: 1rem;
            border-radius: 0.5rem;
            border-left: 5px solid #4A90E2;
            margin-bottom: 1rem;
        }
        div[data-testid="stSidebarNav"] {
            background-color: #262730;
        }
        .stSelectbox {
            background-color: #262730;
        }
        .stProgress .st-bo {
            background-color: #4A90E2;
        }
    </style>
""", unsafe_allow_html=True)

# Data Loading Functions
def load_monthly_sales(excel_file):
    """Load and process monthly sales data"""
    df = pd.read_excel(excel_file, sheet_name="Monthly Sale Tracker ", header=None)
    
    # Find the actual data start
    start_row = df[df[0] == 'Sr No'].index[0]
    headers = df.iloc[start_row]
    df = df.iloc[start_row+1:].reset_index(drop=True)
    df.columns = headers
    
    # Convert numeric columns
    numeric_cols = ['Area', 'Sale Consideration']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
    # Convert date columns if present
    if 'Feed in 30-Month-Year format' in df.columns:
        df['Feed in 30-Month-Year format'] = pd.to_datetime(df['Feed in 30-Month-Year format'], errors='coerce')
    
    return df

def load_bank_balances(excel_file):
    """Load and process bank balance data"""
    df = pd.read_excel(excel_file, sheet_name="Bank Balances", header=None)
    
    # Find the actual data start
    start_row = df[df[0] == 'A/c #'].index[0]
    headers = df.iloc[start_row]
    df = df.iloc[start_row+1:].reset_index(drop=True)
    df.columns = headers
    
    # Convert numeric columns
    numeric_cols = ['Credit', 'Debit', 'Balance']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    return df

def load_cashflow_process(excel_file):
    """Load and process cashflow data"""
    df = pd.read_excel(excel_file, sheet_name="Cashflow Process", header=None)
    
    # Process the cashflow data
    inflows_row = df[df[0] == 'Inflows'].index[0]
    outflows_row = df[df[0] == 'Outflows'].index[0]
    
    # Separate inflows and outflows
    inflows = df.iloc[inflows_row+1:outflows_row].dropna(how='all')
    outflows = df.iloc[outflows_row+1:].dropna(how='all')
    
    return inflows, outflows

def load_project_outflow(excel_file):
    """Load and process project outflow statement"""
    df = pd.read_excel(excel_file, sheet_name="Project Outflow Statement")
    df['Date of Payment'] = pd.to_datetime(df['Date of Payment'], errors='coerce')
    df['Gross Amount'] = pd.to_numeric(df['Gross Amount'], errors='coerce')
    return df.dropna(how='all')

def load_admin_expenses(excel_file):
    """Load and process admin expenses"""
    df = pd.read_excel(excel_file, sheet_name="Admin Expense Tracker ")
    # Convert columns to numeric, skipping the first few categorical columns
    numeric_cols = df.columns[4:]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

# Visualization Functions
def create_sales_analysis(sales_df):
    """Create comprehensive sales analysis"""
    st.header("📈 Sales Analysis")
    
    # Key Metrics Row
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        total_sales = sales_df['Sale Consideration'].sum()
        st.metric("Total Sales Value", f"₹{total_sales:,.2f}Cr")
    with col2:
        avg_price = sales_df['Sale Consideration'].mean()
        st.metric("Average Unit Price", f"₹{avg_price:,.2f}Cr")
    with col3:
        total_units = len(sales_df)
        st.metric("Total Units Sold", f"{total_units:,}")
    with col4:
        total_area = sales_df['Area'].sum()
        st.metric("Total Area Sold", f"{total_area:,.2f} sq.ft")
    
    # Sales Trends and Distribution
    col1, col2 = st.columns(2)
    
    with col1:
        # Sales by BHK Distribution
        sales_by_bhk = sales_df.groupby('BHK')['Sale Consideration'].sum()
        fig_bhk = px.pie(
            values=sales_by_bhk.values,
            names=sales_by_bhk.index,
            title="Sales Distribution by BHK Type",
            color_discrete_sequence=px.colors.sequential.Viridis
        )
        fig_bhk.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white'
        )
        st.plotly_chart(fig_bhk, use_container_width=True)
        
    with col2:
        # Area vs Price Scatter Plot
        fig_scatter = px.scatter(
            sales_df,
            x='Area',
            y='Sale Consideration',
            color='BHK',
            title='Area vs Price Analysis',
            color_discrete_sequence=px.colors.sequential.Viridis
        )
        fig_scatter.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white'
        )
        st.plotly_chart(fig_scatter, use_container_width=True)
    
    # Monthly Sales Trend
    sales_df['Month'] = pd.to_datetime(sales_df['Feed in 30-Month-Year format']).dt.to_period('M')
    monthly_sales = sales_df.groupby('Month')['Sale Consideration'].sum().reset_index()
    monthly_sales['Month'] = monthly_sales['Month'].astype(str)
    
    fig_trend = px.line(
        monthly_sales,
        x='Month',
        y='Sale Consideration',
        title='Monthly Sales Trend',
        color_discrete_sequence=['#4A90E2']
    )
    fig_trend.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='white'
    )
    st.plotly_chart(fig_trend, use_container_width=True)

def create_financial_analysis(bank_df, outflow_df):
    """Create comprehensive financial analysis"""
    st.header("💰 Financial Analysis")
    
    # Key Metrics
    col1, col2, col3 = st.columns(3)
    with col1:
        total_balance = bank_df['Balance'].sum()
        st.metric("Total Bank Balance", f"₹{total_balance:,.2f}Cr")
    with col2:
        total_credits = bank_df['Credit'].sum()
        st.metric("Total Credits", f"₹{total_credits:,.2f}Cr")
    with col3:
        total_debits = bank_df['Debit'].sum()
        st.metric("Total Debits", f"₹{total_debits:,.2f}Cr")
    
    # Bank Balance Distribution
    fig_bank = px.bar(
        bank_df,
        x='Account Description',
        y=['Credit', 'Debit', 'Balance'],
        title='Account-wise Financial Overview',
        barmode='group',
        color_discrete_sequence=px.colors.sequential.Viridis
    )
    fig_bank.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='white'
    )
    st.plotly_chart(fig_bank, use_container_width=True)
    
    # Outflow Analysis
    if 'Date of Payment' in outflow_df.columns and 'Gross Amount' in outflow_df.columns:
        monthly_outflow = outflow_df.groupby(outflow_df['Date of Payment'].dt.to_period('M'))['Gross Amount'].sum().reset_index()
        monthly_outflow['Date of Payment'] = monthly_outflow['Date of Payment'].astype(str)
        
        fig_outflow = px.line(
            monthly_outflow,
            x='Date of Payment',
            y='Gross Amount',
            title='Monthly Outflow Trend',
            color_discrete_sequence=['#FF6B6B']
        )
        fig_outflow.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white'
        )
        st.plotly_chart(fig_outflow, use_container_width=True)

def create_expense_analysis(expense_df):
    """Create expense analysis dashboard"""
    st.header("📊 Expense Analysis")
    
    # Calculate monthly totals for each category
    monthly_totals = expense_df.iloc[:, 4:].sum()
    
    # Create trend chart
    fig_trend = px.line(
        x=range(len(monthly_totals)),
        y=monthly_totals,
        title='Monthly Expense Trend',
        color_discrete_sequence=['#4CAF50']
    )
    fig_trend.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='white'
    )
    st.plotly_chart(fig_trend, use_container_width=True)
    
    # Category-wise analysis
    if 'Particulars' in expense_df.columns:
        category_totals = expense_df.groupby('Particulars')[expense_df.columns[4:]].sum().sum(axis=1)
        fig_category = px.pie(
            values=category_totals.values,
            names=category_totals.index,
            title='Expense Distribution by Category',
            color_discrete_sequence=px.colors.sequential.Viridis
        )
        fig_category.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white'
        )
        st.plotly_chart(fig_category, use_container_width=True)

def main():
    st.title("📊 Project Management Dashboard")
    
    uploaded_file = st.file_uploader("Upload Project Excel File", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            with st.spinner("Processing data..."):
                # Load all data
                sales_data = load_monthly_sales(uploaded_file)
                bank_data = load_bank_balances(uploaded_file)
                inflows, outflows = load_cashflow_process(uploaded_file)
                outflow_data = load_project_outflow(uploaded_file)
                expense_data = load_admin_expenses(uploaded_file)
                
                # Navigation
                st.sidebar.title("Navigation")
                page = st.sidebar.radio(
                    "Select Section",
                    ["Sales Analysis", "Financial Overview", "Expense Analysis"]
                )
                
                if page == "Sales Analysis":
                    create_sales_analysis(sales_data)
                elif page == "Financial Overview":
                    create_financial_analysis(bank_data, outflow_data)
                else:
                    create_expense_analysis(expense_data)
                
                # Footer
                st.markdown("---")
                st.markdown(
                    f"<div style='text-align: center; color: #666;'>"
                    f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                    f"</div>",
                    unsafe_allow_html=True
                )
        
        except Exception as e:
            st.error(f"Error processing data: {str(e)}")
            st.info("Please check the Excel file structure and try again.")
    else:
        st.markdown(
            """
            <div class="uploadSection">
                <h3>Welcome to Project Management Dashboard</h3>
                <p>Please upload your project Excel file to view the analysis.</p>
                <p style='color: #666; font-size: 0.8rem;'>Supported formats: .xlsx, .xls</p>
            </div>
            """,
            unsafe_allow_html=True
        )

if __name__ == "__main__":
    main()
