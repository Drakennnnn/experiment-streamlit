import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np
from streamlit_extras.metric_cards import style_metric_cards
from streamlit_extras.colored_header import colored_header
import plotly.figure_factory as ff

# Page Configuration
st.set_page_config(
    page_title="Project Management Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Enhanced Custom CSS
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
            margin-bottom: 1rem;
        }
        .stMetric {
            background-color: #262730;
            padding: 1.5rem;
            border-radius: 0.5rem;
            margin: 0.5rem 0;
        }
        div[data-testid="stMetricValue"] {
            color: #FAFAFA;
            font-size: 2rem !important;
        }
        div[data-testid="stMetricDelta"] {
            color: #4CAF50;
            font-size: 1rem !important;
        }
        div[data-testid="stMetricLabel"] {
            color: #FAFAFA;
            font-size: 1.2rem !important;
            font-weight: 500;
        }
        .metric-row {
            display: flex;
            justify-content: space-between;
            gap: 1rem;
            margin: 1rem 0;
        }
        .custom-card {
            background-color: #262730;
            padding: 1.5rem;
            border-radius: 0.5rem;
            margin: 1rem 0;
            border-left: 4px solid #4A90E2;
        }
        .custom-header {
            font-size: 1.5rem;
            font-weight: 600;
            margin-bottom: 1rem;
            color: #4A90E2;
        }
        .data-table {
            background-color: #262730;
            border-radius: 0.5rem;
            padding: 1rem;
        }
        .stDataFrame {
            background-color: #262730;
        }
        div[data-testid="stSidebarNav"] {
            background-color: #262730;
        }
    </style>
""", unsafe_allow_html=True)

# Data Loading Functions
def load_sales_data(file):
    """Load and process sales data"""
    df = pd.read_excel(file, sheet_name="Monthly Sale Tracker ", header=None)
    # Find header row
    header_row = df[df[0] == 'Sr No'].index[0]
    headers = df.iloc[header_row]
    df = df.iloc[header_row + 1:].copy()
    df.columns = headers
    
    # Convert numeric columns
    numeric_cols = ['Area', 'Sale Consideration', 'BSP']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # Convert date
    if 'Feed in 30-Month-Year format' in df.columns:
        df['Feed in 30-Month-Year format'] = pd.to_datetime(
            df['Feed in 30-Month-Year format'], 
            errors='coerce'
        )
    
    return df.dropna(subset=['Sale Consideration'])

def load_bank_data(file):
    """Load and process bank balance data"""
    df = pd.read_excel(file, sheet_name="Bank Balances", header=None)
    header_row = df[df[0] == 'A/c #'].index[0]
    headers = df.iloc[header_row]
    df = df.iloc[header_row + 1:].copy()
    df.columns = headers
    
    numeric_cols = ['Credit', 'Debit', 'Balance']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    return df.dropna(subset=['Balance'])

def load_outflow_data(file):
    """Load and process project outflow data"""
    df = pd.read_excel(file, sheet_name="Project Outflow Statement")
    df['Date of Payment'] = pd.to_datetime(df['Date of Payment'], errors='coerce')
    df['Gross Amount'] = pd.to_numeric(df['Gross Amount'], errors='coerce')
    return df.dropna(subset=['Gross Amount'])

def load_collection_data(file):
    """Load and process main collection data"""
    df = pd.read_excel(file, sheet_name="Main Collection AC P1_P2_P3")
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
    return df.dropna(subset=['Amount'])

def create_enhanced_sales_analysis(sales_df):
    """Create enhanced sales analysis dashboard"""
    colored_header("📈 Sales Analysis", description="Comprehensive sales performance metrics")
    
    # Key Metrics Row
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_sales = sales_df['Sale Consideration'].sum()
        st.metric(
            "Total Sales",
            f"₹{total_sales/10000000:.2f}Cr",
            delta=f"+{4.5}%",
            delta_color="normal"
        )
    
    with col2:
        total_units = len(sales_df)
        st.metric(
            "Total Units",
            f"{total_units}",
            delta=f"{total_units-100}",
            delta_color="normal"
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
    
    style_metric_cards()
    
    # Enhanced Area vs Price Analysis
    st.markdown("### Area vs Price Analysis")
    
    # Create violin plot for price distribution by BHK
    fig_violin = go.Figure()
    for bhk in sales_df['BHK'].unique():
        subset = sales_df[sales_df['BHK'] == bhk]
        fig_violin.add_trace(go.Violin(
            x=subset['BHK'],
            y=subset['Sale Consideration']/10000000,
            name=bhk,
            box_visible=True,
            meanline_visible=True
        ))
    
    fig_violin.update_layout(
        title="Price Distribution by BHK Type",
        yaxis_title="Price (Cr)",
        showlegend=True,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='white'
    )
    st.plotly_chart(fig_violin, use_container_width=True)
    
    # Enhanced Scatter Plot
    fig_scatter = px.scatter(
        sales_df,
        x='Area',
        y='Sale Consideration'/10000000,
        color='BHK',
        size='Sale Consideration'/1000000,  # Size based on price
        hover_data={
            'Area': ':.0f',
            'Sale Consideration': ':,.2f',
            'BHK': True,
            'Tower': True
        },
        title='Area vs Price Relationship',
        labels={
            'Sale Consideration': 'Price (Cr)',
            'Area': 'Area (sq.ft)'
        }
    )
    
    fig_scatter.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='white',
        showlegend=True,
        legend=dict(
            bgcolor='rgba(0,0,0,0)',
            font=dict(color='white')
        )
    )
    
    st.plotly_chart(fig_scatter, use_container_width=True)
    
    # Sales Trend Analysis
    st.markdown("### Sales Trend Analysis")
    
    monthly_sales = sales_df.groupby(
        pd.to_datetime(sales_df['Feed in 30-Month-Year format']).dt.to_period('M')
    ).agg({
        'Sale Consideration': 'sum',
        'Sr No': 'count'
    }).reset_index()
    
    monthly_sales['Feed in 30-Month-Year format'] = monthly_sales['Feed in 30-Month-Year format'].astype(str)
    
    # Create double-axis chart
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    fig.add_trace(
        go.Scatter(
            x=monthly_sales['Feed in 30-Month-Year format'],
            y=monthly_sales['Sale Consideration']/10000000,
            name="Sales Value (Cr)",
            line=dict(color="#4CAF50", width=2)
        ),
        secondary_y=False
    )
    
    fig.add_trace(
        go.Bar(
            x=monthly_sales['Feed in 30-Month-Year format'],
            y=monthly_sales['Sr No'],
            name="Units Sold",
            marker_color="rgba(74, 144, 226, 0.7)"
        ),
        secondary_y=True
    )
    
    fig.update_layout(
        title="Monthly Sales Trend",
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='white',
        xaxis_tickangle=45,
        legend=dict(
            bgcolor='rgba(0,0,0,0)',
            font=dict(color='white')
        )
    )
    
    fig.update_yaxes(
        title_text="Sales Value (Cr)",
        secondary_y=False,
        gridcolor='rgba(255,255,255,0.1)'
    )
    fig.update_yaxes(
        title_text="Units Sold",
        secondary_y=True,
        gridcolor='rgba(255,255,255,0.1)'
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Tower-wise Analysis
    st.markdown("### Tower-wise Analysis")
    
    tower_data = sales_df.groupby('Tower').agg({
        'Sale Consideration': ['sum', 'mean', 'count'],
        'Area': 'sum'
    }).reset_index()
    
    tower_data.columns = ['Tower', 'Total Sales', 'Avg Price', 'Units Sold', 'Total Area']
    
    # Format the values
    tower_data['Total Sales'] = tower_data['Total Sales']/10000000
    tower_data['Avg Price'] = tower_data['Avg Price']/10000000
    
    fig_tower = go.Figure()
    
    fig_tower.add_trace(go.Bar(
        x=tower_data['Tower'],
        y=tower_data['Total Sales'],
        name='Total Sales (Cr)',
        marker_color='#4CAF50'
    ))
    
    fig_tower.add_trace(go.Scatter(
        x=tower_data['Tower'],
        y=tower_data['Units Sold'],
        name='Units Sold',
        yaxis='y2',
        line=dict(color='#4A90E2', width=2)
    ))
    
    fig_tower.update_layout(
        title='Tower-wise Performance',
        yaxis=dict(title='Total Sales (Cr)'),
        yaxis2=dict(title='Units Sold', overlaying='y', side='right'),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='white',
        showlegend=True,
        legend=dict(
            bgcolor='rgba(0,0,0,0)',
            font=dict(color='white')
        )
    )
    
    st.plotly_chart(fig_tower, use_container_width=True)

def create_financial_analysis(bank_df, outflow_df=None):
    """Create enhanced financial analysis dashboard"""
    colored_header("💰 Financial Analysis", description="Bank balances and financial metrics")
    
    # Key Metrics
    col1, col2, col3 = st.columns(3)
    
    with col1:
        total_balance = bank_df['Balance'].sum()
        st.metric(
            "Total Balance",
            f"₹{total_balance/10000000:.2f}Cr",
            delta=f"+{2.5}%"
        )
    
    with col2:
        total_credit = bank_df['Credit'].sum()
        st.metric(
            "Total Credits",
            f"₹{total_credit/10000000:.2f}Cr"
        )
    
    with col3:
        total_debit = bank_df['Debit'].sum()
        st.metric(
            "Total Debits",
            f"₹{total_debit/10000000:.2f}Cr"
        )
    
    style_metric_cards()
    
    # Account Balance Distribution
    st.markdown("### Account Balance Distribution")
    
    fig_balance = go.Figure()
    
    # Add traces for Credit, Debit, and Balance
    fig_balance.add_trace(go.Bar(
        name='Credit',
        x=bank_df['Account Description'],
        y=bank_df['Credit']/10000000,
        marker_color='#4CAF50'
    ))
    
    fig_balance.add_trace(go.Bar(
        name='Debit',
        x=bank_df['Account Description'],
        y=bank_df['Debit']/10000000,
        marker_color='#FF6B6B'
    ))
    
    fig_balance.add_trace(go.Bar(
        name='Balance',
        x=bank_df['Account Description'],
        y=bank_df['Balance']/10000000,
        marker_color='#4A90E2'
    ))
    
    fig_balance.update_layout(
        barmode='group',
        title='Account-wise Financial Overview (Cr)',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='white',
        xaxis_tickangle=45,
        legend=dict(
            bgcolor='rgba(0,0,0,0)',
            font=dict(color='white')
        ),
        height=600
    )
    
    st.plotly_chart(fig_balance, use_container_width=True)
    
    # Detailed Account Table with formatting
    st.markdown("### Account Details")

    formatted_df = bank_df.copy()
    formatted_df[['Credit', 'Debit', 'Balance']] = formatted_df[['Credit', 'Debit', 'Balance']]/10000000
    
    st.dataframe(
        formatted_df[['Account Description', 'Credit', 'Debit', 'Balance']]
        .style.format({
            'Credit': '₹{:,.2f} Cr',
            'Debit': '₹{:,.2f} Cr',
            'Balance': '₹{:,.2f} Cr'
        }).background_gradient(cmap='RdYlGn', subset=['Balance'])
    )

    if outflow_df is not None:
        st.markdown("### Project Outflow Analysis")
        
        # Monthly outflow trend
        monthly_outflow = outflow_df.groupby(
            pd.to_datetime(outflow_df['Date of Payment']).dt.to_period('M')
        )['Gross Amount'].sum().reset_index()
        
        monthly_outflow['Date of Payment'] = monthly_outflow['Date of Payment'].astype(str)
        
        fig_outflow = px.line(
            monthly_outflow,
            x='Date of Payment',
            y='Gross Amount'/10000000,
            title='Monthly Project Outflow Trend',
            labels={'Gross Amount': 'Amount (Cr)'},
            line_shape='linear'
        )
        
        fig_outflow.update_traces(line=dict(color="#FF6B6B", width=2))
        fig_outflow.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            xaxis_tickangle=45
        )
        
        st.plotly_chart(fig_outflow, use_container_width=True)

def create_project_progress(outflow_df, collection_df=None):
    """Create project progress analysis"""
    colored_header("🏗️ Project Progress", description="Construction and collection progress")
    
    if outflow_df is not None:
        # Vendor Analysis
        st.markdown("### Vendor Payment Analysis")
        
        vendor_payments = outflow_df.groupby('Vendor')['Gross Amount'].agg(['sum', 'count']).reset_index()
        vendor_payments.columns = ['Vendor', 'Total Amount', 'Number of Payments']
        vendor_payments = vendor_payments.sort_values('Total Amount', ascending=False).head(10)
        
        fig_vendor = go.Figure()
        
        fig_vendor.add_trace(go.Bar(
            x=vendor_payments['Vendor'],
            y=vendor_payments['Total Amount']/10000000,
            name='Total Amount (Cr)',
            marker_color='#4CAF50'
        ))
        
        fig_vendor.add_trace(go.Scatter(
            x=vendor_payments['Vendor'],
            y=vendor_payments['Number of Payments'],
            name='Number of Payments',
            yaxis='y2',
            line=dict(color='#4A90E2', width=2)
        ))
        
        fig_vendor.update_layout(
            title='Top 10 Vendors by Payment Amount',
            yaxis=dict(title='Total Amount (Cr)'),
            yaxis2=dict(title='Number of Payments', overlaying='y', side='right'),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            showlegend=True,
            xaxis_tickangle=45,
            height=600
        )
        
        st.plotly_chart(fig_vendor, use_container_width=True)
        
        # Payment Category Analysis
        st.markdown("### Payment Category Analysis")
        
        category_payments = outflow_df.groupby('Code Tagging')['Gross Amount'].sum().reset_index()
        category_payments = category_payments.sort_values('Gross Amount', ascending=True)
        
        fig_category = go.Figure(go.Bar(
            x=category_payments['Gross Amount']/10000000,
            y=category_payments['Code Tagging'],
            orientation='h',
            marker_color='#4A90E2'
        ))
        
        fig_category.update_layout(
            title='Payment Distribution by Category',
            xaxis_title='Amount (Cr)',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            height=600
        )
        
        st.plotly_chart(fig_category, use_container_width=True)

def main():
    st.title("Project Management Dashboard")
    
    st.markdown("""
        <div style='background-color: #262730; padding: 1rem; border-radius: 0.5rem; margin-bottom: 1rem;'>
            <h4 style='color: #4A90E2; margin: 0;'>Upload Project Excel File</h4>
            <p style='color: #FAFAFA; margin: 0.5rem 0 0 0;'>Please upload your project management Excel file to view the analysis.</p>
        </div>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            with st.spinner("Processing data..."):
                # Load all required data
                sales_df = load_sales_data(uploaded_file)
                bank_df = load_bank_data(uploaded_file)
                outflow_df = load_outflow_data(uploaded_file)
                collection_df = load_collection_data(uploaded_file)
                
                # Navigation
                st.sidebar.title("Navigation")
                page = st.sidebar.radio(
                    "Select Analysis",
                    ["Sales Analysis", "Financial Overview", "Project Progress"]
                )
                
                # Display selected analysis
                if page == "Sales Analysis":
                    create_enhanced_sales_analysis(sales_df)
                elif page == "Financial Overview":
                    create_financial_analysis(bank_df, outflow_df)
                else:
                    create_project_progress(outflow_df, collection_df)
                
                # Footer
                st.markdown("---")
                st.markdown(
                    f"<div style='text-align: center; color: #666;'>"
                    f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>",
                    unsafe_allow_html=True
                )
        
        except Exception as e:
            st.error(f"Error processing data: {str(e)}")
            st.info("Please check the Excel file structure and try again.")
    
if __name__ == "__main__":
    main()
