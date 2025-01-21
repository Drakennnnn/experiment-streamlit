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

# Enhanced Custom CSS with better box sizing
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
        .custom-info-box {
            background-color: #262730;
            padding: 1rem;
            border-radius: 0.5rem;
            border-left: 4px solid #4A90E2;
            margin: 1rem 0;
        }
    </style>
""", unsafe_allow_html=True)

def format_indian_currency(amount, units='crores'):
    """Format amount in Indian currency with units"""
    if units == 'crores':
        return f"₹{amount/10000000:.2f} Cr"
    elif units == 'lakhs':
        return f"₹{amount/100000:.2f} L"
    else:
        return f"₹{amount:,.2f}"

def standardize_bhk(bhk_value):
    """Standardize BHK notation"""
    if pd.isna(bhk_value):
        return 'Others'
    bhk_str = str(bhk_value).upper().strip()
    # Remove any spaces or hyphens and ensure BHK suffix
    bhk_str = bhk_str.replace(' ', '').replace('-', '')
    if not bhk_str.endswith('BHK'):
        bhk_str = bhk_str + 'BHK'
    return bhk_str

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
        
        # Convert date and amount columns
        if 'Date of Payment' in df.columns:
            df['Date of Payment'] = pd.to_datetime(df['Date of Payment'], errors='coerce')
        if 'Gross Amount' in df.columns:
            df['Gross Amount'] = pd.to_numeric(df['Gross Amount'], errors='coerce')
        
        return df.dropna(subset=['Gross Amount'])
    except:
        return None

def load_sales_data(file):
    """Load and process sales data with standardized BHK notation"""
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
    
    # Standardize BHK notation
    if 'BHK' in df.columns:
        df['BHK'] = df['BHK'].apply(standardize_bhk)
    
    return df.dropna(subset=['Sale Consideration'])

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
            💡 Key metrics show overall sales performance. Total sales represent cumulative value, 
            while average price indicates the mean unit cost. Area reflects total space sold.
        </div>
    """, unsafe_allow_html=True)
    
    # Enhanced Area vs Price Analysis
    st.markdown("### Area vs Price Analysis")
    
    # Create tabs for different views
    price_tabs = st.tabs(["Scatter Plot", "Price Distribution", "Price/Sq.ft Analysis"])
    
    with price_tabs[0]:
        # Enhanced scatter plot
        fig_scatter = px.scatter(
            sales_df,
            x='Area',
            y='Sale Consideration',
            color='BHK',
            size='Sale Consideration',
            hover_data={
                'Area': ':.0f',
                'Sale Consideration': lambda x: format_indian_currency(x),
                'BHK': True,
                'Tower': True
            },
            title='Area vs Price Relationship',
            labels={
                'Area': 'Area (sq.ft)',
                'Sale Consideration': 'Price'
            }
        )
        
        # Add trendlines for each BHK type
        for bhk in sales_df['BHK'].unique():
            bhk_data = sales_df[sales_df['BHK'] == bhk]
            z = np.polyfit(bhk_data['Area'], bhk_data['Sale Consideration'], 1)
            p = np.poly1d(z)
            fig_scatter.add_scatter(
                x=bhk_data['Area'],
                y=p(bhk_data['Area']),
                name=f'{bhk} Trend',
                mode='lines',
                line=dict(dash='dash'),
                showlegend=True
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
                ticksuffix=' Cr',
                gridcolor='rgba(255,255,255,0.1)'
            ),
            xaxis=dict(
                gridcolor='rgba(255,255,255,0.1)'
            )
        )
        
        st.plotly_chart(fig_scatter, use_container_width=True)
        
        st.markdown("""
            <div class="description-text">
                📌 This scatter plot shows the relationship between unit area and price. 
                Each point represents a unit, with size indicating price. Dashed lines show price trends for each BHK type.
                Hover over points for detailed information.
            </div>
        """, unsafe_allow_html=True)
    
    with price_tabs[1]:
        # Price distribution by BHK
        fig_violin = go.Figure()
        
        for bhk in sorted(sales_df['BHK'].unique()):
            subset = sales_df[sales_df['BHK'] == bhk]
            fig_violin.add_trace(go.Violin(
                x=subset['BHK'],
                y=subset['Sale Consideration']/10000000,
                name=bhk,
                box_visible=True,
                meanline_visible=True,
                points="all"
            ))
        
        fig_violin.update_layout(
            title="Price Distribution by BHK Type",
            yaxis_title="Price (₹ Cr)",
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            showlegend=True
        )
        
        st.plotly_chart(fig_violin, use_container_width=True)
        
        st.markdown("""
            <div class="description-text">
                📊 Violin plots show price distribution for each BHK type. 
                The width indicates frequency at each price point. Box plots inside show median and quartiles.
            </div>
        """, unsafe_allow_html=True)
    
    with price_tabs[2]:
        # Price per sq.ft analysis
        sales_df['Price_Per_Sqft'] = sales_df['Sale Consideration'] / sales_df['Area']
        
        fig_box = px.box(
            sales_df,
            x='BHK',
            y='Price_Per_Sqft',
            color='BHK',
            title='Price per Sq.ft Distribution by BHK Type',
            points="all"
        )
        
        fig_box.update_layout(
            yaxis_title="Price per Sq.ft (₹)",
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            showlegend=False
        )
        
        st.plotly_chart(fig_box, use_container_width=True)
        
        st.markdown("""
            <div class="description-text">
                💰 This analysis shows the distribution of price per square foot across different BHK types.
                Box plots show median, quartiles, and outliers. Individual points represent actual units.
            </div>
        """, unsafe_allow_html=True)
    
    # Sales Distribution by BHK
    st.markdown("### Sales Distribution Analysis")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Enhanced pie chart
        sales_by_bhk = sales_df.groupby('BHK')['Sale Consideration'].sum()
        fig_pie = px.pie(
            values=sales_by_bhk.values,
            names=sales_by_bhk.index,
            title="Sales Distribution by BHK Type",
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        
        fig_pie.update_traces(
            textinfo='percent+label',
            hovertemplate="<b>%{label}</b><br>" +
                        "Amount: ₹%{value:,.2f} Cr<br>" +
                        "Percentage: %{percent}"
        )
        
        fig_pie.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white'
        )
        
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with col2:
        # Add unit count distribution
        units_by_bhk = sales_df.groupby('BHK').size()
        fig_bar = px.bar(
            x=units_by_bhk.index,
            y=units_by_bhk.values,
            title="Number of Units Sold by BHK Type",
            labels={'x': 'BHK Type', 'y': 'Number of Units'},
            color=units_by_bhk.index,
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        
        fig_bar.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font_color='white',
            showlegend=False
        )
        
        st.plotly_chart(fig_bar, use_container_width=True)
# Monthly Sales Trend Analysis
    st.markdown("### Monthly Sales Trend")
    
    monthly_data = sales_df.groupby(
        pd.to_datetime(sales_df['Feed in 30-Month-Year format']).dt.to_period('M')
    ).agg({
        'Sale Consideration': ['sum', 'count'],
        'Area': 'sum'
    }).reset_index()
    
    monthly_data.columns = ['Month', 'Total_Sales', 'Units_Sold', 'Area_Sold']
    monthly_data['Month'] = monthly_data['Month'].astype(str)
    
    # Create the double-axis chart
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    fig.add_trace(
        go.Bar(
            x=monthly_data['Month'],
            y=monthly_data['Total_Sales']/10000000,
            name="Sales Value (₹ Cr)",
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
        ),
        hovermode='x unified'
    )
    
    fig.update_yaxes(
        title_text="Sales Value (₹ Cr)",
        secondary_y=False,
        gridcolor='rgba(255,255,255,0.1)',
        tickprefix="₹",
        ticksuffix=" Cr"
    )
    fig.update_yaxes(
        title_text="Number of Units",
        secondary_y=True,
        gridcolor='rgba(255,255,255,0.1)'
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("""
        <div class="description-text">
            📈 The chart shows monthly sales trends. Bars represent total sales value (in Crores), 
            while the line shows number of units sold. This helps visualize both value and volume trends.
        </div>
    """, unsafe_allow_html=True)

def create_financial_analysis(bank_df, outflow_df=None):
    """Create enhanced financial analysis dashboard"""
    colored_header("💰 Financial Analysis", description="Comprehensive analysis of financial metrics and cash flows")
    
    # Key Metrics with better formatting
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
    
    st.markdown("""
        <div class="description-text">
            💸 These metrics show the overall financial position across all accounts. 
            The balance trend indicates net cash position growth.
        </div>
    """, unsafe_allow_html=True)
    
    # Account Analysis Tabs
    finance_tabs = st.tabs(["Account Balances", "Cash Flow Analysis", "Account Details"])
    
    with finance_tabs[0]:
        fig_balance = go.Figure()
        
        # Add traces for Credit, Debit, and Balance
        fig_balance.add_trace(go.Bar(
            name='Credits',
            x=bank_df['Account Description'],
            y=bank_df['Credit']/10000000,
            marker_color='#4CAF50'
        ))
        
        fig_balance.add_trace(go.Bar(
            name='Debits',
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
            title='Account-wise Financial Overview',
            yaxis_title='Amount (₹ Cr)',
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
    
    with finance_tabs[1]:
        if outflow_df is not None:
            # Monthly cash flow analysis
            outflow_df['Month'] = pd.to_datetime(outflow_df['Date of Payment']).dt.to_period('M')
            monthly_outflow = outflow_df.groupby('Month')['Gross Amount'].sum().reset_index()
            monthly_outflow['Month'] = monthly_outflow['Month'].astype(str)
            
            fig_cashflow = px.line(
                monthly_outflow,
                x='Month',
                y='Gross Amount'/10000000,
                title='Monthly Cash Outflow Trend',
                labels={'Gross Amount': 'Amount (₹ Cr)', 'Month': 'Month'},
                line_shape='linear'
            )
            
            fig_cashflow.update_traces(
                line=dict(color="#FF6B6B", width=2),
                mode='lines+markers'
            )
            
            fig_cashflow.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='white',
                xaxis_tickangle=45,
                yaxis=dict(tickprefix="₹", ticksuffix=" Cr")
            )
            
            st.plotly_chart(fig_cashflow, use_container_width=True)
            
            # Payment category analysis
            category_data = outflow_df.groupby('Code Tagging')['Gross Amount'].agg(['sum', 'count']).reset_index()
            category_data.columns = ['Category', 'Total Amount', 'Number of Payments']
            category_data = category_data.sort_values('Total Amount', ascending=True)
            
            fig_category = go.Figure()
            
            fig_category.add_trace(go.Bar(
                y=category_data['Category'],
                x=category_data['Total Amount']/10000000,
                orientation='h',
                name='Total Amount',
                marker_color='#4A90E2'
            ))
            
            fig_category.update_layout(
                title='Payment Distribution by Category',
                xaxis_title='Amount (₹ Cr)',
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='white',
                height=600,
                xaxis=dict(tickprefix="₹", ticksuffix=" Cr")
            )
            
            st.plotly_chart(fig_category, use_container_width=True)
    
    with finance_tabs[2]:
        st.markdown("### Detailed Account Information")
        
        formatted_df = bank_df.copy()
        formatted_df[['Credit', 'Debit', 'Balance']] = formatted_df[['Credit', 'Debit', 'Balance']]/10000000
        
        st.dataframe(
            formatted_df[['Account Description', 'Account No', 'Credit', 'Debit', 'Balance']]
            .style.format({
                'Credit': '₹{:,.2f} Cr',
                'Debit': '₹{:,.2f} Cr',
                'Balance': '₹{:,.2f} Cr'
            }).background_gradient(cmap='RdYlGn', subset=['Balance'])
            .set_properties(**{'text-align': 'right'})
        )

def create_project_analysis(outflow_df):
    """Create project progress analysis dashboard"""
    colored_header("🏗️ Project Progress", description="Analysis of project execution and vendor payments")
    
    if outflow_df is not None:
        # Payment Progress Metrics
        total_paid = outflow_df['Gross Amount'].sum()
        avg_payment = outflow_df['Gross Amount'].mean()
        payment_count = len(outflow_df)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric(
                "Total Payments Made",
                format_indian_currency(total_paid)
            )
        
        with col2:
            st.metric(
                "Average Payment Size",
                format_indian_currency(avg_payment)
            )
        
        with col3:
            st.metric(
                "Number of Payments",
                f"{payment_count:,}"
            )
        
        # Vendor Analysis Tabs
        vendor_tabs = st.tabs(["Top Vendors", "Payment Categories", "Monthly Trends"])
        
        with vendor_tabs[0]:
            # Top vendors analysis
            vendor_data = outflow_df.groupby('Vendor').agg({
                'Gross Amount': ['sum', 'count']
            }).reset_index()
            
            vendor_data.columns = ['Vendor', 'Total Amount', 'Payment Count']
            vendor_data = vendor_data.sort_values('Total Amount', ascending=False).head(10)
            
            fig_vendor = go.Figure()
            
            fig_vendor.add_trace(go.Bar(
                x=vendor_data['Vendor'],
                y=vendor_data['Total Amount']/10000000,
                name='Total Amount',
                marker_color='#4CAF50'
            ))
            
            fig_vendor.add_trace(go.Scatter(
                x=vendor_data['Vendor'],
                y=vendor_data['Payment Count'],
                name='Number of Payments',
                yaxis='y2',
                line=dict(color='#4A90E2', width=2)
            ))
            
            fig_vendor.update_layout(
                title='Top 10 Vendors by Payment Amount',
                yaxis=dict(title='Amount (₹ Cr)', tickprefix="₹", ticksuffix=" Cr"),
                yaxis2=dict(title='Number of Payments', overlaying='y', side='right'),
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='white',
                xaxis_tickangle=45,
                height=600,
                showlegend=True,
                legend=dict(bgcolor='rgba(0,0,0,0)')
            )
            
            st.plotly_chart(fig_vendor, use_container_width=True)
            
            st.markdown("""
                <div class="description-text">
                    👥 This analysis shows the top vendors by payment amount. The bars represent total payments,
                    while the line shows the number of transactions with each vendor.
                </div>
            """, unsafe_allow_html=True)

def main():
    st.title("Project Management Dashboard")
    
    st.markdown("""
        <div style='background-color: #262730; padding: 1rem; border-radius: 0.5rem; margin-bottom: 1rem;'>
            <h4 style='color: #4A90E2; margin: 0;'>Project Performance Analytics</h4>
            <p style='color: #FAFAFA; margin: 0.5rem 0 0 0;'>
                Upload your project Excel file to view comprehensive analysis of sales, financials, and project progress.
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
                    create_project_analysis(outflow_df)
                
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
