import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import numpy as np

# Set page config
st.set_page_config(
    page_title="Project Management Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
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

@st.cache_data
def process_bank_data(df):
    """Process bank balance data"""
    try:
        # Skip header rows
        header_row = df[df['Unnamed: 0'] == 'A/c #'].index[0]
        df = df.iloc[header_row:].copy()
        
        # Set headers
        df.columns = df.iloc[0]
        df = df.iloc[1:].reset_index(drop=True)
        
        # Convert numeric columns
        numeric_cols = ['Credit', 'Debit', 'Balance']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        return df
    except Exception as e:
        st.error(f"Error processing bank data: {str(e)}")
        return None

@st.cache_data
def process_sales_data(df):
    """Process sales data"""
    try:
        # Find the data start
        start_row = df[df.iloc[:, 0] == 'Sr No'].index[0]
        df = df.iloc[start_row:].copy()
        
        # Set headers and clean data
        df.columns = df.iloc[0]
        df = df.iloc[1:].reset_index(drop=True)
        
        # Convert numeric columns
        if 'Sale Consideration' in df.columns:
            df['Sale Consideration'] = pd.to_numeric(df['Sale Consideration'], errors='coerce')
        if 'Area' in df.columns:
            df['Area'] = pd.to_numeric(df['Area'], errors='coerce')
            
        return df
    except Exception as e:
        st.error(f"Error processing sales data: {str(e)}")
        return None

@st.cache_data
def process_bplan_data(df):
    """Process BPlan tracker data"""
    try:
        # First two rows contain headers and dates
        df = df.iloc[2:].copy()  # Skip the header rows
        
        # Set the category column
        df['Category'] = df.iloc[:, 1]
        
        # Convert all numeric columns
        numeric_cols = df.columns[2:-3]  # Exclude first two columns and last three
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
        return df
    except Exception as e:
        st.error(f"Error processing BPlan data: {str(e)}")
        return None

def load_data(uploaded_file):
    """Load and process all data from Excel file"""
    try:
        # Read all sheets
        xl = pd.ExcelFile(uploaded_file)
        data = {}
        
        # Process Bank Balances
        if 'Bank Balances' in xl.sheet_names:
            df_bank = pd.read_excel(xl, 'Bank Balances', header=None)
            data['bank'] = process_bank_data(df_bank)
            
        # Process Monthly Sale Tracker
        if 'Monthly Sale Tracker ' in xl.sheet_names:
            df_sales = pd.read_excel(xl, 'Monthly Sale Tracker ', header=None)
            data['sales'] = process_sales_data(df_sales)
            
        # Process BPlan Tracker
        if 'BPlan Tracker' in xl.sheet_names:
            df_bplan = pd.read_excel(xl, 'BPlan Tracker', header=None)
            data['bplan'] = process_bplan_data(df_bplan)
            
        return data
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return None

def render_bank_analysis(data):
    """Render bank balance analysis"""
    if 'bank' in data and data['bank'] is not None:
        st.subheader("Bank Balance Analysis")
        
        # Summary metrics
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Balance", 
                     f"₹{data['bank']['Balance'].sum():,.2f}")
        with col2:
            st.metric("Total Credits", 
                     f"₹{data['bank']['Credit'].sum():,.2f}")
        with col3:
            st.metric("Total Debits", 
                     f"₹{data['bank']['Debit'].sum():,.2f}")
        
        # Bank balance distribution
        fig = px.pie(data['bank'], 
                    values='Balance',
                    names='Account Description',
                    title='Distribution of Bank Balances')
        st.plotly_chart(fig, use_container_width=True)
        
        # Detailed balance table
        st.subheader("Detailed Bank Balances")
        st.dataframe(data['bank'][['Account Description', 'Credit', 'Debit', 'Balance']]
                    .style.format({
                        'Credit': '₹{:,.2f}',
                        'Debit': '₹{:,.2f}',
                        'Balance': '₹{:,.2f}'
                    }))

def render_sales_analysis(data):
    """Render sales analysis"""
    if 'sales' in data and data['sales'] is not None:
        st.subheader("Sales Analysis")
        
        # Sales summary
        if 'Sale Consideration' in data['sales'].columns:
            total_sales = data['sales']['Sale Consideration'].sum()
            avg_sale = data['sales']['Sale Consideration'].mean()
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total Sales Value", 
                         f"₹{total_sales:,.2f}")
            with col2:
                st.metric("Average Sale Value", 
                         f"₹{avg_sale:,.2f}")
        
        # Sales by BHK type
        if 'BHK' in data['sales'].columns and 'Sale Consideration' in data['sales'].columns:
            sales_by_bhk = data['sales'].groupby('BHK')['Sale Consideration'].sum().reset_index()
            fig = px.pie(sales_by_bhk,
                        values='Sale Consideration',
                        names='BHK',
                        title='Sales Distribution by BHK Type')
            st.plotly_chart(fig, use_container_width=True)
        
        # Area vs Price Analysis
        if all(col in data['sales'].columns for col in ['Area', 'Sale Consideration']):
            fig = px.scatter(data['sales'],
                           x='Area',
                           y='Sale Consideration',
                           color='BHK' if 'BHK' in data['sales'].columns else None,
                           title='Area vs Sale Price Analysis')
            st.plotly_chart(fig, use_container_width=True)

def render_bplan_analysis(data):
    """Render BPlan analysis"""
    if 'bplan' in data and data['bplan'] is not None:
        st.subheader("Business Plan Analysis")
        
        # Calculate progress
        actual_sum = data['bplan'].iloc[:, 2:-3].sum().sum()
        budget_sum = data['bplan'].iloc[:, -2].sum()
        progress = (actual_sum / budget_sum * 100) if budget_sum != 0 else 0
        
        st.metric("Overall Plan Progress",
                 f"{progress:.1f}%",
                 f"{actual_sum/budget_sum:.2f}x of budget")
        
        # Monthly trend
        monthly_totals = data['bplan'].iloc[:, 2:-3].sum()
        fig = px.line(x=monthly_totals.index,
                     y=monthly_totals.values,
                     title='Monthly Expenditure Trend')
        st.plotly_chart(fig, use_container_width=True)
        
        # Category-wise analysis
        categories = data['bplan']['Category']
        actuals = data['bplan'].iloc[:, 2:-3].sum(axis=1)
        budget = data['bplan'].iloc[:, -2]
        
        comparison_df = pd.DataFrame({
            'Category': categories,
            'Actual': actuals,
            'Budget': budget
        })
        
        fig = go.Figure(data=[
            go.Bar(name='Actual', x=comparison_df['Category'], y=comparison_df['Actual']),
            go.Bar(name='Budget', x=comparison_df['Category'], y=comparison_df['Budget'])
        ])
        fig.update_layout(title='Category-wise Budget vs Actual',
                         barmode='group')
        st.plotly_chart(fig, use_container_width=True)

def main():
    st.title("📊 Project Management Dashboard")
    
    # File uploader
    uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Load data
        with st.spinner("Processing data..."):
            data = load_data(uploaded_file)
            
        if data:
            # Navigation
            st.sidebar.title("Navigation")
            page = st.sidebar.radio("Select Analysis",
                                  ["Bank Analysis",
                                   "Sales Analysis",
                                   "Business Plan Analysis"])
            
            # Render selected page
            if page == "Bank Analysis":
                render_bank_analysis(data)
            elif page == "Sales Analysis":
                render_sales_analysis(data)
            else:
                render_bplan_analysis(data)
                
            # Footer
            st.markdown("---")
            st.markdown(f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    else:
        st.info("Please upload an Excel file to view the dashboard")

if __name__ == "__main__":
    main()
