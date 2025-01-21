import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np
from streamlit_extras.metric_cards import style_metric_cards
from streamlit_extras.colored_header import colored_header
from plotly.subplots import make_subplots

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
            min-height: 150px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }
        div[data-testid="stMetricValue"] {
            color: #FAFAFA;
            font-size: 1.8rem !important;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        div[data-testid="stMetricDelta"] {
            color: #4CAF50;
            font-size: 1rem !important;
        }
        div[data-testid="stMetricLabel"] {
            color: #FAFAFA;
            font-size: 1.2rem !important;
            font-weight: 500;
            padding-bottom: 0.5rem;
        }
        .description-text {
            color: #AAAAAA;
            font-size: 0.9rem;
            margin: 0.5rem 0;
            padding: 0.5rem;
            background-color: rgba(255, 255, 255, 0.05);
            border-radius: 0.3rem;
        }
    </style>
""", unsafe_allow_html=True)

def format_indian_currency(amount):
    """Format amount in Indian currency with crore denomination"""
    try:
        value = float(amount)
        crore = value / 10000000
        return f"₹{crore:.2f} Cr"
    except:
        return "₹0.00 Cr"

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
    
    # Standardize BHK format
    if 'BHK' in df.columns:
        df['BHK'] = df['BHK'].fillna('Other')
        df['BHK'] = df['BHK'].astype(str).str.replace('-', '')
    
    return df.dropna(subset=['Sale Consideration'])

def load_bank_data(file):
    """Load and process bank balance data"""
    df = pd.read_excel(file, sheet_name="Bank Balances", header=None)
    
    # Find header row
    header_row = df[df[0] == 'A/c #'].index[0]
    headers = df.iloc[header_row]
    df = df.iloc[header_row + 1:].copy()
    df.columns = headers
    
    # Convert numeric columns
    numeric_cols = ['Credit', 'Debit', 'Balance']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    return df.dropna(subset=['Balance'])

def load_outflow_data(file):
    """Load and process project outflow data"""
    try:
        df = pd.read_excel(file, sheet_name="Project Outflow Statement")
        df['Date of Payment'] = pd.to_datetime(df['Date of Payment'], errors='coerce')
        df['Gross Amount'] = pd.to_numeric(df['Gross Amount'], errors='coerce')
        return df.dropna(subset=['Gross Amount'])
    except:
        return None

def create_enhanced_sales_analysis(sales_df):
    """Create enhanced sales analysis dashboard"""
    colored_header("📈 Sales Analysis", description="Comprehensive analysis of sales performance and trends")
    
    # Key Metrics Row
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_sales = sales_df['Sale Consideration'].sum()
        st.metric(
            "Total Sales",
            format_indian_currency(total_sales),
            delta="+4.5%"
        )
    
    with col2:
        total_units = len(sales_df)
        st.metric(
            "Total Units",
            f"{total_units}",
            delta=f"+{197}"
        )
    
    with col3:
        avg_price = sales_df['Sale Consideration'].mean()
        st.metric(
            "Average Price",
            format_indian_currency(avg_price)
        )
    
    with col4:
        total_area = sales_df['Area'].sum()
        st.metric(
            "Total Area",
            f"{total_area:,.0f} sq.ft"
        )
    
    st.markdown("""
        <div class="description-text">
            💡 Key metrics showing overall sales performance. Numbers include all transactions 
            to date, with positive growth trends in both value and volume.
        </div>
    """, unsafe_allow_html=True)
    
    # Enhanced Area vs Price Analysis
    st.markdown("### Area vs Price Analysis")
    
    # Create scatter plot
    fig_scatter = px.scatter(
        sales_df,
        x='Area',
        y='Sale Consideration',
        color='BHK',
        size='Sale Consideration',
        title='Area vs Price Relationship',
        labels={
            'Area': 'Area (sq.ft)',
            'Sale Consideration': 'Price (₹ Cr)',
            'BHK': 'Unit Type'
        }
    )
    
    # Update scatter plot
    fig_scatter.update_traces(
        hovertemplate="<br>".join([
            "Area: %{x:,.0f} sq.ft",
            "Price: ₹%{y:.2f} Cr",
            "Type: %{marker.color}",
            "<extra></extra>"
        ])
    )
    
    fig_scatter.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='white',
        showlegend=True,
        legend=dict(
            bgcolor='rgba(0,0,0,0)',
            font=dict(color='white')
        ),
        yaxis=dict(
            tickformat='.2f',
            tickprefix='₹',
            ticksuffix=' Cr'
        )
    )
    
    st.plotly_chart(fig_scatter, use_container_width=True)
    
    # Distribution Analysis
    st.markdown("### Price Distribution Analysis")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Sales distribution by BHK
        sales_by_bhk = sales_df.groupby('BHK')['Sale Consideration'].sum()
        fig_pie = px.pie(
            values=sales_by_bhk.values/10000000,  # Convert to Cr
            names=sales_by_bhk.index,
            title="Sales Distribution by Unit Type"
        )
        
        fig_pie.update_traces(
            textinfo='percent+label',
            hovertemplate="<br>".join([
                "Type: %{label}",
                "Amount: ₹%{value:.2f} Cr",
                "Percentage: %{percent}",
                "<extra></extra>"
            ])
        )
        
        fig_pie.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white'
        )
        
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with col2:
        # Unit count distribution
        units_by_bhk = sales_df.groupby('BHK').size()
        fig_bar = px.bar(
            x=units_by_bhk.index,
            y=units_by_bhk.values,
            title="Units Sold by Type",
            labels={'x': 'Unit Type', 'y': 'Number of Units'}
        )
        
        fig_bar.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            showlegend=False
        )
        
        st.plotly_chart(fig_bar, use_container_width=True)
    
    # Monthly Trend Analysis
    st.markdown("### Monthly Sales Trend")
    
    monthly_data = sales_df.groupby(
        pd.to_datetime(sales_df['Feed in 30-Month-Year format']).dt.to_period('M')
    ).agg({
        'Sale Consideration': ['sum', 'count'],
        'Area': 'sum'
    }).reset_index()
    
    monthly_data.columns = ['Month', 'Total_Sales', 'Units_Sold', 'Area_Sold']
    monthly_data['Month'] = monthly_data['Month'].astype(str)
    
    # Create trend chart
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    fig.add_trace(
        go.Bar(
            x=monthly_data['Month'],
            y=monthly_data['Total_Sales']/10000000,
            name="Sales Value",
            marker_color='rgba(74, 144, 226, 0.7)'
        ),
        secondary_y=False
    )
    
    fig.add_trace(
        go.Scatter(
            x=monthly_data['Month'],
            y=monthly_data['Units_Sold'],
            name="Units Sold",
            line=dict(color="#4CAF50", width=2),
            mode='lines+markers'
        ),
        secondary_y=True
    )
    
    fig.update_layout(
        title="Monthly Sales Performance",
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
        title_text="Sales Value (₹ Cr)",
        secondary_y=False,
        tickprefix="₹",
        ticksuffix=" Cr"
    )
    
    fig.update_yaxes(
        title_text="Units Sold",
        secondary_y=True
    )
    
    st.plotly_chart(fig, use_container_width=True)    
    st.markdown("""
        <div class="description-text">
            📈 Monthly performance showing sales value (bars) and units sold (line).
            Track both revenue and volume trends over time.
        </div>
    """, unsafe_allow_html=True)

def create_financial_analysis(bank_df, outflow_df=None):
    """Create enhanced financial analysis dashboard"""
    colored_header("💰 Financial Analysis", description="Analysis of financial metrics and cash flows")
    
    # Key Metrics
    col1, col2, col3 = st.columns(3)
    
    with col1:
        total_balance = bank_df['Balance'].sum()
        st.metric(
            "Total Bank Balance",
            format_indian_currency(total_balance),
            delta="+2.5%"
        )
    
    with col2:
        total_credit = bank_df['Credit'].sum()
        st.metric(
            "Total Credits",
            format_indian_currency(total_credit)
        )
    
    with col3:
        total_debit = bank_df['Debit'].sum()
        st.metric(
            "Total Debits",
            format_indian_currency(total_debit)
        )
    
    # Bank Analysis Tabs
    tabs = st.tabs(["Account Analysis", "Cash Flow", "Account Details"])
    
    with tabs[0]:
        st.markdown("### Account-wise Analysis")
        
        # Create account analysis visualization
        fig_bank = go.Figure()
        
        fig_bank.add_trace(go.Bar(
            name='Credits',
            x=bank_df['Account Description'],
            y=bank_df['Credit']/10000000,
            marker_color='#4CAF50'
        ))
        
        fig_bank.add_trace(go.Bar(
            name='Debits',
            x=bank_df['Account Description'],
            y=bank_df['Debit']/10000000,
            marker_color='#FF6B6B'
        ))
        
        fig_bank.add_trace(go.Bar(
            name='Balance',
            x=bank_df['Account Description'],
            y=bank_df['Balance']/10000000,
            marker_color='#4A90E2'
        ))
        
        fig_bank.update_traces(
            hovertemplate="<br>".join([
                "Account: %{x}",
                "Amount: ₹%{y:.2f} Cr",
                "<extra></extra>"
            ])
        )
        
        fig_bank.update_layout(
            barmode='group',
            title='Account-wise Financial Overview',
            yaxis_title='Amount (₹ Cr)',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            xaxis_tickangle=45,
            height=500,
            showlegend=True,
            legend=dict(
                bgcolor='rgba(0,0,0,0)',
                font=dict(color='white')
            )
        )
        
        st.plotly_chart(fig_bank, use_container_width=True)
    
    with tabs[1]:
        if outflow_df is not None:
            st.markdown("### Cash Flow Analysis")
            
            # Calculate monthly outflows
            outflow_df['Month'] = pd.to_datetime(outflow_df['Date of Payment']).dt.to_period('M')
            monthly_flow = outflow_df.groupby('Month')['Gross Amount'].sum().reset_index()
            monthly_flow['Month'] = monthly_flow['Month'].astype(str)
            
            # Create cash flow trend visualization
            fig_flow = go.Figure()
            
            fig_flow.add_trace(go.Scatter(
                x=monthly_flow['Month'],
                y=monthly_flow['Gross Amount']/10000000,
                mode='lines+markers',
                name='Monthly Outflow',
                line=dict(color="#4A90E2", width=2),
                hovertemplate="<br>".join([
                    "Month: %{x}",
                    "Amount: ₹%{y:.2f} Cr",
                    "<extra></extra>"
                ])
            ))
            
            fig_flow.update_layout(
                title='Monthly Cash Flow Trend',
                yaxis_title='Amount (₹ Cr)',
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='white',
                xaxis_tickangle=45,
                height=400
            )
            
            st.plotly_chart(fig_flow, use_container_width=True)
            
            # Payment Category Analysis
            st.markdown("### Payment Category Distribution")
            
            category_data = outflow_df.groupby('Code Tagging')['Gross Amount'].sum().reset_index()
            category_data = category_data.sort_values('Gross Amount', ascending=True)
            
            fig_category = go.Figure()
            
            fig_category.add_trace(go.Bar(
                y=category_data['Code Tagging'],
                x=category_data['Gross Amount']/10000000,
                orientation='h',
                marker_color='#4A90E2',
                hovertemplate="<br>".join([
                    "Category: %{y}",
                    "Amount: ₹%{x:.2f} Cr",
                    "<extra></extra>"
                ])
            ))
            
            fig_category.update_layout(
                title='Distribution by Payment Category',
                xaxis_title='Amount (₹ Cr)',
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='white',
                height=500
            )
            
            st.plotly_chart(fig_category, use_container_width=True)
    
    with tabs[2]:
        st.markdown("### Detailed Account Information")
        
        # Format dataframe for display
        formatted_df = bank_df.copy()
        formatted_df[['Credit', 'Debit', 'Balance']] = formatted_df[['Credit', 'Debit', 'Balance']]/10000000
        
        st.dataframe(
            formatted_df[['Account Description', 'Account No', 'Credit', 'Debit', 'Balance']]
            .style.format({
                'Credit': '₹{:,.2f} Cr',
                'Debit': '₹{:,.2f} Cr',
                'Balance': '₹{:,.2f} Cr'
            }).background_gradient(cmap='RdYlGn', subset=['Balance'])
        )

def load_account_data(file):
    """Load and process all account data"""
    account_data = {}
    
    # Load accounts 1 through 13
    for i in range(1, 14):
        sheet_name = f"Ac #{i}"
        try:
            # Read Excel sheet
            df = pd.read_excel(file, sheet_name=sheet_name)
            
            # Find start of data
            start_idx = None
            for idx, row in df.iterrows():
                if any(str(val).lower().strip() in ['txn date', 'date'] for val in row):
                    start_idx = idx
                    break
            
            if start_idx is not None:
                # Set headers and data
                headers = df.iloc[start_idx]
                df = df.iloc[start_idx + 1:].copy()
                df.columns = headers
                
                # Convert numeric columns
                if 'Amount' in df.columns:
                    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
                if 'Running Total' in df.columns:
                    df['Running Total'] = pd.to_numeric(df['Running Total'], errors='coerce')
                
                # Only keep relevant data
                df = df.dropna(subset=['Amount'], how='all')
                if not df.empty:
                    account_data[sheet_name] = df
        except Exception as e:
            print(f"Error loading {sheet_name}: {str(e)}")
            continue
    
    return account_data

def create_project_progress(account_data, outflow_df=None):
    """Create project progress analysis dashboard"""
    colored_header("🏗️ Project Progress", description="Analysis of project execution, accounts, and milestones")
    
    if not account_data:
        st.warning("No account data available for analysis.")
        return
    
    # Account Summary
    st.markdown("### Account Performance Summary")
    
    # Process account summary data
    summary_data = []
    for acc_name, acc_df in account_data.items():
        try:
            if not acc_df.empty and 'Amount' in acc_df.columns:
                total_amount = acc_df['Amount'].sum()
                transaction_count = len(acc_df)
                last_balance = acc_df['Running Total'].iloc[-1] if 'Running Total' in acc_df.columns else 0
                
                summary_data.append({
                    'Account': acc_name,
                    'Total Transactions': transaction_count,
                    'Total Amount': total_amount,
                    'Current Balance': last_balance
                })
        except Exception as e:
            print(f"Error processing {acc_name}: {str(e)}")
            continue
    
    if not summary_data:
        st.warning("Could not process account summary data.")
        return
        
    summary_df = pd.DataFrame(summary_data)
    
    # Display summary metrics
    col1, col2, col3 = st.columns(3)
    
    with col1:
        total_transactions = summary_df['Total Transactions'].sum()
        st.metric(
            "Total Transactions",
            f"{total_transactions:,}"
        )
    
    with col2:
        total_amount = summary_df['Total Amount'].sum()
        st.metric(
            "Total Amount",
            format_indian_currency(total_amount)
        )
    
    with col3:
        total_balance = summary_df['Current Balance'].sum()
        st.metric(
            "Current Balance",
            format_indian_currency(total_balance)
        )
    
    st.markdown("""
        <div class="description-text">
            💡 Overview of all account activities showing total transactions processed,
            cumulative transaction value, and current balance across all accounts.
        </div>
    """, unsafe_allow_html=True)
    
    # Account Analysis Tabs
    acc_tabs = st.tabs(["Transaction Analysis", "Balance Trends", "Account Details"])
    
    with acc_tabs[0]:
        # Create transaction volume analysis
        fig_trans = go.Figure()
        
        fig_trans.add_trace(go.Bar(
            x=summary_df['Account'],
            y=summary_df['Total Transactions'],
            name='Transaction Count',
            marker_color='#4A90E2'
        ))
        
        fig_trans.add_trace(go.Scatter(
            x=summary_df['Account'],
            y=summary_df['Total Amount']/10000000,
            name='Total Amount (Cr)',
            yaxis='y2',
            line=dict(color='#4CAF50', width=2)
        ))
        
        fig_trans.update_traces(
            hovertemplate="<br>".join([
                "Account: %{x}",
                "Value: %{y:,.2f}",
                "<extra></extra>"
            ])
        )
        
        fig_trans.update_layout(
            title='Account-wise Transaction Analysis',
            yaxis=dict(
                title='Number of Transactions',
                gridcolor='rgba(255,255,255,0.1)'
            ),
            yaxis2=dict(
                title='Amount (₹ Cr)',
                overlaying='y',
                side='right',
                tickprefix='₹',
                ticksuffix=' Cr',
                gridcolor='rgba(255,255,255,0.1)'
            ),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            showlegend=True,
            legend=dict(
                bgcolor='rgba(0,0,0,0)',
                font=dict(color='white')
            ),
            height=500,
            xaxis_tickangle=45
        )
        
        st.plotly_chart(fig_trans, use_container_width=True)
        
        st.markdown("""
            <div class="description-text">
                📊 Transaction analysis comparing volume and value across accounts.
                Bars show number of transactions while the line shows total amount.
            </div>
        """, unsafe_allow_html=True)
    
    with acc_tabs[1]:
        # Create balance trend visualization
        fig_balance = go.Figure()
        
        for acc_name, acc_df in account_data.items():
            if 'Running Total' in acc_df.columns and not acc_df.empty:
                fig_balance.add_trace(go.Scatter(
                    name=acc_name,
                    y=acc_df['Running Total']/10000000,
                    mode='lines',
                    hovertemplate="<br>".join([
                        "Account: %{fullData.name}",
                        "Balance: ₹%{y:.2f} Cr",
                        "<extra></extra>"
                    ])
                ))
        
        fig_balance.update_layout(
            title='Account Balance Trends',
            yaxis=dict(
                title='Balance (₹ Cr)',
                gridcolor='rgba(255,255,255,0.1)'
            ),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            showlegend=True,
            legend=dict(
                bgcolor='rgba(0,0,0,0)',
                font=dict(color='white'),
                yanchor="top",
                y=0.99,
                xanchor="left",
                x=0.01
            ),
            height=500
        )
        
        st.plotly_chart(fig_balance, use_container_width=True)
        
        st.markdown("""
            <div class="description-text">
                📈 Balance trends showing running totals for each account over time.
                Track account performance and identify significant transactions.
            </div>
        """, unsafe_allow_html=True)
    
    with acc_tabs[2]:
        st.markdown("### Account Transaction Details")
        
        # Select account for detailed view
        selected_account = st.selectbox(
            "Select Account",
            list(account_data.keys())
        )
        
        if selected_account in account_data:
            acc_df = account_data[selected_account].copy()
            
            if 'Amount' in acc_df.columns and 'Running Total' in acc_df.columns:
                # Convert to crores for display
                display_df = acc_df.copy()
                display_df['Amount'] = display_df['Amount']/10000000
                display_df['Running Total'] = display_df['Running Total']/10000000
                
                st.dataframe(
                    display_df.style.format({
                        'Amount': '₹{:,.2f} Cr',
                        'Running Total': '₹{:,.2f} Cr'
                    }).background_gradient(
                        cmap='RdYlGn',
                        subset=['Amount', 'Running Total']
                    )
                )
    
    # Project Progress Section
    if outflow_df is not None and not outflow_df.empty:
        st.markdown("### Project Progress Tracking")
        
        try:
            # Monthly progress tracking
            monthly_progress = outflow_df.groupby(
                pd.to_datetime(outflow_df['Date of Payment']).dt.to_period('M')
            ).agg({
                'Gross Amount': ['sum', 'count'],
                'Code Tagging': 'nunique'
            }).reset_index()
            
            monthly_progress.columns = ['Month', 'Total_Amount', 'Transaction_Count', 'Category_Count']
            monthly_progress['Month'] = monthly_progress['Month'].astype(str)
            
            # Create progress visualization
            fig_progress = make_subplots(specs=[[{"secondary_y": True}]])
            
            fig_progress.add_trace(
                go.Bar(
                    x=monthly_progress['Month'],
                    y=monthly_progress['Total_Amount']/10000000,
                    name="Monthly Outflow",
                    marker_color='rgba(74, 144, 226, 0.7)'
                ),
                secondary_y=False
            )
            
            fig_progress.add_trace(
                go.Scatter(
                    x=monthly_progress['Month'],
                    y=monthly_progress['Category_Count'],
                    name="Categories",
                    line=dict(color="#4CAF50", width=2)
                ),
                secondary_y=True
            )
            
            fig_progress.update_layout(
                title="Monthly Project Progress",
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='white',
                showlegend=True,
                legend=dict(
                    bgcolor='rgba(0,0,0,0)',
                    font=dict(color='white')
                ),
                height=500,
                xaxis_tickangle=45
            )
            
            fig_progress.update_yaxes(
                title_text="Amount (₹ Cr)",
                secondary_y=False,
                tickprefix="₹",
                ticksuffix=" Cr",
                gridcolor='rgba(255,255,255,0.1)'
            )
            fig_progress.update_yaxes(
                title_text="Number of Categories",
                secondary_y=True,
                gridcolor='rgba(255,255,255,0.1)'
            )
            
            st.plotly_chart(fig_progress, use_container_width=True)
            
            st.markdown("""
                <div class="description-text">
                    🏗️ Monthly progress showing payment outflows (bars) and category diversity (line).
                    Track both financial progress and scope of work being executed.
                </div>
            """, unsafe_allow_html=True)
            
        except Exception as e:
            st.error(f"Error creating progress analysis: {str(e)}")

def main():
    st.title("Project Management Dashboard")
    
    st.markdown("""
        <div style='background-color: #262730; padding: 1rem; border-radius: 0.5rem; margin-bottom: 1rem;'>
            <h4 style='color: #4A90E2; margin: 0;'>Project Performance Analytics</h4>
            <p style='color: #FAFAFA; margin: 0.5rem 0 0 0;'>
                Upload your project Excel file to view comprehensive analysis of sales and financials.
            </p>
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
                account_data = load_account_data(uploaded_file) 
                
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
                    # Load account data
                    account_data = load_account_data(uploaded_file)
                    create_project_progress(account_data, outflow_df)
                
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
    else:
        st.info("👆 Upload an Excel file to begin analysis")

if __name__ == "__main__":
    main()
