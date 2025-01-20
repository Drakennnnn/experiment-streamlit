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
        .stMetric {
            background-color: #262730;
            padding: 1rem;
            border-radius: 0.5rem;
        }
        div[data-testid="stMetricValue"] {
            color: #FAFAFA;
            font-size: 1.8rem !important;
        }
        div[data-testid="stMetricDelta"] {
            color: #4CAF50;
        }
        div[data-testid="stMetricLabel"] {
            color: #FAFAFA;
        }
    </style>
""", unsafe_allow_html=True)

def safe_numeric_convert(df, columns):
    """Safely convert columns to numeric, handling errors"""
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

def load_sales_data(file):
    """Load sales data with proper validation"""
    try:
        # Read the Excel sheet
        df = pd.read_excel(file, sheet_name="Monthly Sale Tracker ", header=None)
        
        # Find the header row (contains 'Sr No')
        header_row = df[df[0] == 'Sr No'].index[0]
        
        # Set the headers and skip to data
        headers = df.iloc[header_row]
        df = df.iloc[header_row + 1:].copy()
        df.columns = headers
        
        # Convert numeric columns
        numeric_cols = ['Area', 'Sale Consideration', 'BSP']
        df = safe_numeric_convert(df, numeric_cols)
        
        # Convert date column
        if 'Feed in 30-Month-Year format' in df.columns:
            df['Feed in 30-Month-Year format'] = pd.to_datetime(
                df['Feed in 30-Month-Year format'], 
                errors='coerce'
            )
        
        return df
    except Exception as e:
        st.error(f"Error loading sales data: {str(e)}")
        return None

def load_bank_data(file):
    """Load bank data with proper validation"""
    try:
        # Read the Excel sheet
        df = pd.read_excel(file, sheet_name="Bank Balances", header=None)
        
        # Find the header row (contains 'A/c #')
        header_row = df[df[0] == 'A/c #'].index[0]
        
        # Set the headers and skip to data
        headers = df.iloc[header_row]
        df = df.iloc[header_row + 1:].copy()
        df.columns = headers
        
        # Convert numeric columns
        numeric_cols = ['Credit', 'Debit', 'Balance']
        df = safe_numeric_convert(df, numeric_cols)
        
        return df
    except Exception as e:
        st.error(f"Error loading bank data: {str(e)}")
        return None

def create_sales_dashboard(sales_df):
    """Create sales analysis dashboard"""
    if sales_df is None or len(sales_df) == 0:
        st.warning("No sales data available")
        return
        
    st.header("Sales Analysis")
    
    # Key Metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_sales = sales_df['Sale Consideration'].sum()
        st.metric(
            "Total Sales",
            f"₹{total_sales/10000000:.2f}Cr",
            delta="4.5%"
        )
    
    with col2:
        total_units = len(sales_df)
        st.metric(
            "Total Units",
            f"{total_units}",
            delta=f"{total_units-100}"
        )
    
    with col3:
        avg_price = sales_df['Sale Consideration'].mean()
        st.metric(
            "Average Price",
            f"₹{avg_price/10000000:.2f}Cr"
        )
    
    with col4:
        total_area = sales_df['Area'].sum()
        st.metric(
            "Total Area",
            f"{total_area:,.0f} sq.ft"
        )
    
    # Sales Distribution by BHK
    col1, col2 = st.columns(2)
    
    with col1:
        if 'BHK' in sales_df.columns:
            sales_by_bhk = sales_df.groupby('BHK')['Sale Consideration'].sum()
            fig_bhk = px.pie(
                values=sales_by_bhk.values,
                names=sales_by_bhk.index,
                title="Sales Distribution by BHK",
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            fig_bhk.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='white'
            )
            st.plotly_chart(fig_bhk, use_container_width=True)
    
    with col2:
        # Area vs Price Analysis
        fig_scatter = px.scatter(
            sales_df,
            x='Area',
            y='Sale Consideration',
            color='BHK',
            title='Area vs Price Analysis',
            labels={'Sale Consideration': 'Price (₹)', 'Area': 'Area (sq.ft)'},
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        fig_scatter.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white'
        )
        st.plotly_chart(fig_scatter, use_container_width=True)
    
    # Monthly Trend
    monthly_sales = sales_df.groupby(
        pd.to_datetime(sales_df['Feed in 30-Month-Year format']).dt.to_period('M')
    )['Sale Consideration'].sum().reset_index()
    monthly_sales['Feed in 30-Month-Year format'] = monthly_sales['Feed in 30-Month-Year format'].astype(str)
    
    fig_trend = px.line(
        monthly_sales,
        x='Feed in 30-Month-Year format',
        y='Sale Consideration',
        title='Monthly Sales Trend',
        labels={'Sale Consideration': 'Sales (₹)', 'Feed in 30-Month-Year format': 'Month'}
    )
    fig_trend.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='white',
        xaxis_tickangle=45
    )
    st.plotly_chart(fig_trend, use_container_width=True)

def create_bank_dashboard(bank_df):
    """Create bank analysis dashboard"""
    if bank_df is None or len(bank_df) == 0:
        st.warning("No bank data available")
        return
        
    st.header("Bank Analysis")
    
    # Key Metrics
    col1, col2, col3 = st.columns(3)
    
    with col1:
        total_balance = bank_df['Balance'].sum()
        st.metric(
            "Total Balance",
            f"₹{total_balance/10000000:.2f}Cr",
            delta="2.5%"
        )
    
    with col2:
        total_credit = bank_df['Credit'].sum()
        st.metric(
            "Total Credit",
            f"₹{total_credit/10000000:.2f}Cr"
        )
    
    with col3:
        total_debit = bank_df['Debit'].sum()
        st.metric(
            "Total Debit",
            f"₹{total_debit/10000000:.2f}Cr"
        )
    
    # Account Balance Distribution
    fig_balance = px.bar(
        bank_df,
        x='Account Description',
        y=['Credit', 'Debit', 'Balance'],
        title='Account-wise Financial Overview',
        barmode='group',
        color_discrete_sequence=px.colors.qualitative.Set3
    )
    fig_balance.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='white',
        xaxis_tickangle=45,
        legend=dict(
            bgcolor='rgba(0,0,0,0)',
            font=dict(color='white')
        )
    )
    st.plotly_chart(fig_balance, use_container_width=True)
    
    # Detailed Account Table
    st.subheader("Account Details")
    st.dataframe(
        bank_df[['Account Description', 'Credit', 'Debit', 'Balance']]
        .style.format({
            'Credit': '₹{:,.2f}',
            'Debit': '₹{:,.2f}',
            'Balance': '₹{:,.2f}'
        })
    )

def main():
    st.title("Project Management Dashboard")
    
    uploaded_file = st.file_uploader("Upload Project Excel File", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Load data
        with st.spinner("Processing data..."):
            sales_df = load_sales_data(uploaded_file)
            bank_df = load_bank_data(uploaded_file)
            
            if sales_df is not None or bank_df is not None:
                # Navigation
                pages = ["Sales Analysis", "Bank Analysis"]
                selection = st.sidebar.radio("Select Analysis", pages)
                
                if selection == "Sales Analysis":
                    create_sales_dashboard(sales_df)
                else:
                    create_bank_dashboard(bank_df)
                
                # Footer
                st.markdown("---")
                st.markdown(
                    f"<div style='text-align: center; color: #666;'>"
                    f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>",
                    unsafe_allow_html=True
                )
            else:
                st.error("Could not load data from the Excel file")
    else:
        st.info("Please upload an Excel file to begin analysis")

if __name__ == "__main__":
    main()
