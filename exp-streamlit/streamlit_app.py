import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np

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
        .small-text {
            font-size: 0.8rem;
            color: #666;
        }
    </style>
""", unsafe_allow_html=True)

def process_cashflow_data(df):
    """Process cashflow data from the Cashflow Process sheet"""
    # Skip header rows and find the actual data
    df = df.dropna(how='all')
    inflow_start = df[df['Unnamed: 0'] == 'Inflows'].index[0]
    df = df.iloc[inflow_start:]
    
    # Separate inflows and outflows
    inflows = df[df['Unnamed: 0'] == 'Inflows'].index[0]
    outflows = df[df['Unnamed: 0'] == 'Outflows'].index[0]
    
    # Calculate total inflows and outflows
    inflow_data = df.iloc[inflows+1:outflows].sum(numeric_only=True)
    outflow_data = df.iloc[outflows+1:].sum(numeric_only=True)
    
    return inflow_data.sum(), outflow_data.sum()

def process_bank_balances(df):
    """Process bank balance data"""
    # Find the actual data start
    start_idx = df[df['Unnamed: 0'] == 'A/c #'].index[0]
    df = df.iloc[start_idx:]
    df.columns = df.iloc[0]
    df = df.iloc[1:]
    
    # Clean up and convert balance column
    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
    return df

def process_sales_data(df):
    """Process monthly sales tracker data"""
    # Find the actual data start
    sales_data = df[df['Unnamed: 0'] == 'Sr No'].index[0]
    df = df.iloc[sales_data:]
    df.columns = df.iloc[0]
    df = df.iloc[1:]
    
    # Convert relevant columns to numeric
    numeric_cols = ['Area', 'Sale Consideration']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    return df

def process_bplan_data(df):
    """Process BPlan tracker data"""
    # First row contains dates
    dates = df.iloc[0, 2:].dropna()
    # Second row contains month numbers
    months = df.iloc[1, 2:].dropna()
    # Get category and values
    categories = df.iloc[2:, 1]
    values = df.iloc[2:, 2:len(dates)+2]
    
    # Create processed dataframe
    processed_df = pd.DataFrame(values.values, 
                              index=categories,
                              columns=dates)
    return processed_df

def load_excel_data(uploaded_file):
    """Load and process Excel data from uploaded file"""
    try:
        # Read all sheets
        xls = pd.ExcelFile(uploaded_file)
        
        # Process each sheet
        data = {}
        
        # Cashflow Process
        if 'Cashflow Process' in xls.sheet_names:
            cashflow_df = pd.read_excel(xls, 'Cashflow Process')
            data['inflow'], data['outflow'] = process_cashflow_data(cashflow_df)
        
        # Bank Balances
        if 'Bank Balances' in xls.sheet_names:
            bank_df = pd.read_excel(xls, 'Bank Balances')
            data['bank'] = process_bank_balances(bank_df)
        
        # Monthly Sale Tracker
        if 'Monthly Sale Tracker ' in xls.sheet_names:
            sales_df = pd.read_excel(xls, 'Monthly Sale Tracker ')
            data['sales'] = process_sales_data(sales_df)
        
        # BPlan Tracker
        if 'BPlan Tracker' in xls.sheet_names:
            bplan_df = pd.read_excel(xls, 'BPlan Tracker')
            data['bplan'] = process_bplan_data(bplan_df)
        
        # Admin Expense Tracker
        if 'Admin Expense Tracker ' in xls.sheet_names:
            expense_df = pd.read_excel(xls, 'Admin Expense Tracker ')
            data['expenses'] = expense_df
            
        return data
    except Exception as e:
        st.error(f"Error processing Excel file: {str(e)}")
        return None

def render_overview(data):
    """Render the overview section"""
    # Key Metrics Row
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if 'bank' in data:
            total_balance = data['bank']['Balance'].sum()
            st.metric(
                label="Total Bank Balance",
                value=f"₹{total_balance:,.2f}",
                delta="4.5%"
            )
    
    with col2:
        if 'inflow' in data and 'outflow' in data:
            net_flow = data['inflow'] - data['outflow']
            st.metric(
                label="Net Cashflow",
                value=f"₹{net_flow:,.2f}",
                delta=f"{(net_flow/data['inflow']*100):.1f}%" if data['inflow'] != 0 else "0%"
            )
    
    with col3:
        if 'sales' in data:
            total_sales = data['sales']['Sale Consideration'].sum()
            st.metric(
                label="Total Sales Value",
                value=f"₹{total_sales:,.2f}",
                delta="2.8%"
            )
    
    with col4:
        if 'bplan' in data:
            completion = (data['bplan'].iloc[:, :-3].sum().sum() / 
                        data['bplan']['Budgeted'].sum() * 100)
            st.metric(
                label="Plan Completion",
                value=f"{completion:.1f}%",
                delta=f"{(completion-95):.1f}%"
            )

    # Bank Balance Distribution
    if 'bank' in data:
        st.subheader("Bank Balance Distribution")
        fig_bank = px.pie(
            data['bank'],
            values='Balance',
            names='Account Description',
            title='Distribution of Funds Across Accounts'
        )
        st.plotly_chart(fig_bank, use_container_width=True)

    # Monthly Sales Trend
    if 'sales' in data:
        st.subheader("Sales Analysis")
        sales_by_month = data['sales'].groupby('Feed in 30-Month-Year format')['Sale Consideration'].sum().reset_index()
        sales_trend = px.line(
            sales_by_month,
            x='Feed in 30-Month-Year format',
            y='Sale Consideration',
            title='Monthly Sales Trend'
        )
        st.plotly_chart(sales_trend, use_container_width=True)

def render_financial_analysis(data):
    """Render the financial analysis section"""
    st.header("Financial Analysis")
    
    if 'bplan' in data:
        # Budget vs Actual Analysis
        st.subheader("Budget vs Actual Expenditure")
        budget_comp = pd.DataFrame({
            'Category': data['bplan'].index,
            'Actual': data['bplan'].iloc[:, :-3].sum(axis=1),
            'Budgeted': data['bplan']['Budgeted']
        })
        
        fig = go.Figure(data=[
            go.Bar(name='Actual', x=budget_comp['Category'], y=budget_comp['Actual']),
            go.Bar(name='Budgeted', x=budget_comp['Category'], y=budget_comp['Budgeted'])
        ])
        fig.update_layout(barmode='group')
        st.plotly_chart(fig, use_container_width=True)
        
        # Variance Analysis
        st.subheader("Budget Variance Analysis")
        budget_comp['Variance'] = budget_comp['Actual'] - budget_comp['Budgeted']
        budget_comp['Variance %'] = (budget_comp['Variance'] / budget_comp['Budgeted'] * 100).round(2)
        
        fig_variance = px.bar(
            budget_comp,
            x='Category',
            y='Variance',
            color='Variance',
            title='Budget Variance by Category'
        )
        st.plotly_chart(fig_variance, use_container_width=True)

def render_sales_analysis(data):
    """Render the sales analysis section"""
    st.header("Sales Analysis")
    
    if 'sales' in data:
        col1, col2 = st.columns(2)
        
        with col1:
            # Sales by Unit Type
            sales_by_bhk = data['sales'].groupby('BHK')['Sale Consideration'].sum().reset_index()
            fig_bhk = px.pie(
                sales_by_bhk,
                values='Sale Consideration',
                names='BHK',
                title='Sales Distribution by Unit Type'
            )
            st.plotly_chart(fig_bhk)
        
        with col2:
            # Sales by Tower
            sales_by_tower = data['sales'].groupby('Tower')['Sale Consideration'].sum().reset_index()
            fig_tower = px.bar(
                sales_by_tower,
                x='Tower',
                y='Sale Consideration',
                title='Sales by Tower'
            )
            st.plotly_chart(fig_tower)
        
        # Area vs Price Analysis
        st.subheader("Area vs Price Analysis")
        fig_scatter = px.scatter(
            data['sales'],
            x='Area',
            y='Sale Consideration',
            color='BHK',
            title='Unit Area vs Sale Price'
        )
        st.plotly_chart(fig_scatter, use_container_width=True)

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
        with st.spinner('Loading and processing data...'):
            data = load_excel_data(uploaded_file)
        
        if data:
            # Sidebar
            st.sidebar.header("Dashboard Controls")
            selected_view = st.sidebar.selectbox(
                "Select View",
                ["Overview", "Financial Analysis", "Sales Analysis"]
            )
            
            # Render selected view
            if selected_view == "Overview":
                render_overview(data)
            elif selected_view == "Financial Analysis":
                render_financial_analysis(data)
            elif selected_view == "Sales Analysis":
                render_sales_analysis(data)
            
            # Footer
            st.markdown("---")
            st.markdown(f"Dashboard last updated: {datetime.now().strftime('%B %d, %Y %H:%M:%S')}")
        else:
            st.error("Could not process the Excel file. Please check the file format and try again.")
    else:
        st.info("👆 Upload an Excel file to get started")

if __name__ == "__main__":
    main()
