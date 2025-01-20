import streamlit as st
import pandas as pd

# Title
st.title("Financial Performance Dashboard")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    st.success("File uploaded successfully!")
    # Load the Excel file
    excel_data = pd.ExcelFile(uploaded_file, engine="openpyxl")
    sheet_names = excel_data.sheet_names

    # Select sheet to analyze
    selected_sheet = st.selectbox("Select a sheet", sheet_names)

    # Load selected sheet
    df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

    st.subheader(f"Preview of {selected_sheet}")
    st.write(df.head())

    # Sheet-specific analysis
    if selected_sheet == "Sales & Collection MIS":
        st.subheader("Sales Analysis")
        if 'Pre-Funding' in df.columns and 'As On Date' in df.columns:
            df['As On Date'] = pd.to_datetime(df['As On Date'])
            st.line_chart(df.set_index('As On Date')['Pre-Funding'])
            st.metric(label="Total Pre-Funding Sales", value=df['Pre-Funding'].sum())
        else:
            st.warning("Relevant columns not found for visualization.")

    elif selected_sheet == "Bank Balances":
        st.subheader("Bank Balance Overview")
        if 'Balance' in df.columns:
            st.bar_chart(df[['Balance']])
            st.metric(label="Total Balance", value=df['Balance'].sum())
        else:
            st.warning("Balance column not found.")

    elif selected_sheet == "Project Outflow Statement":
        st.subheader("Expense Overview")
        if 'Amount' in df.columns:
            st.area_chart(df[['Amount']])
            st.metric(label="Total Expenses", value=df['Amount'].sum())
        else:
            st.warning("Amount column not found.")
    
    elif selected_sheet == "Resi Sales MIS":
        st.subheader("Residential Sales Insights")
        if 'Name of Customer' in df.columns and 'Status' in df.columns:
            st.write(df.groupby('Status')['Name of Customer'].count().reset_index())
            st.bar_chart(df['Status'].value_counts())
        else:
            st.warning("Required columns not found.")

    else:
        st.info("Select a different sheet for visualization.")

    # Download processed data
    st.download_button(label="Download Processed Data",
                       data=df.to_csv().encode('utf-8'),
                       file_name=f"{selected_sheet}_processed.csv",
                       mime="text/csv")
