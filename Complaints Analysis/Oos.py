import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from io import BytesIO
import re


st.title("📊 Complaints Dashboard")


@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name='Internal Requests')
    df = df.rename(columns={
        'Main Category Name': 'Main_Category',
        'Product Name': 'Product',
        'Ticket Created At eet': 'Report_Time',
        'FP Name': 'Facility',
        'Picker Name': 'Picker'
    })
    
    
    def extract_reporter(subject):
        if pd.isna(subject):
            return "Unknown"
        match = re.search(r'reported by (.+)$', str(subject))
        return match.group(1).strip() if match else "Unknown"
    
    df['Reporter'] = df['Ticket Subject'].apply(extract_reporter)

    df['Report_Time'] = pd.to_datetime(df['Report_Time'], errors='coerce')
    df = df.dropna(subset=['Report_Time'])
    df['Day'] = df['Report_Time'].dt.day_name()
    df['Hour'] = df['Report_Time'].dt.hour
    df['Date'] = df['Report_Time'].dt.date
    df['Main_Category'] = df['Main_Category'].fillna('Undefined')
    df['Main_Category'] = df['Main_Category'].replace('', 'Undefined')
    
    return df

# File uploader for global use
uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx"])
if not uploaded_file:
    st.info("👋 Please upload an Excel file to get started")
    st.stop()

df = load_data(uploaded_file)

# Sidebar filters
st.sidebar.header("Filters")
date_range = st.sidebar.date_input("Select date range", [df['Date'].min(), df['Date'].max()])
category_filter = st.sidebar.multiselect("Select category", options=df['Main_Category'].unique(), default=df['Main_Category'].unique())

# Apply filters
filtered_df = df[
    (df['Date'] >= date_range[0]) & (df['Date'] <= date_range[1]) &
    (df['Main_Category'].isin(category_filter))
]

st.subheader("Summary")
col1, col2, col3 = st.columns(3)
col1.metric("Total complaints", len(filtered_df))
col2.metric("Categories", len(filtered_df['Main_Category'].unique()))
col3.metric("Reporters", len(filtered_df['Reporter'].unique()))
st.write(f"**Date range:** {date_range[0]} to {date_range[1]}")

st.subheader(" Category Reporting Breakdown")

# Calculate category stats
category_report = filtered_df.groupby('Main_Category').agg(
    Report_Count=('Main_Category', 'count'),
    Reporters=('Reporter', lambda x: x.value_counts().to_dict())
).reset_index()

# Calculate percentage of total reports
total_reports = len(filtered_df)
category_report['Percentage'] = (category_report['Report_Count'] / total_reports * 100).round(1)

# Format reporter information
def format_reporters(reporter_dict):
    sorted_reporters = sorted(reporter_dict.items(), key=lambda x: x[1], reverse=True)
    return ", ".join([f"{reporter} ({count})" for reporter, count in sorted_reporters])

category_report['Top_Reporters'] = category_report['Reporters'].apply(format_reporters)

# Display category table
st.dataframe(category_report[['Main_Category', 'Report_Count', 'Percentage', 'Top_Reporters']]
             .rename(columns={
                 'Main_Category': 'Category',
                 'Report_Count': 'Complaints',
                 'Top_Reporters': 'Reporters (Count)'
             })
             .sort_values('Complaints', ascending=False)
             .style.format({'Percentage': '{:.1f}%'}),
             height=400)

# Visualization
st.subheader("Category Analysis")

tab1, tab2, tab3 = st.tabs(["Bar Chart", "Pie Chart", "Reporter Analysis"])

with tab1:
    # Bar chart
    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(data=category_report, x='Percentage', y='Main_Category', 
                palette='viridis', order=category_report.sort_values('Percentage', ascending=False).Main_Category)
    plt.title('Complaints Percentage by Category')
    plt.xlabel('Percentage of Total Complaints')
    plt.ylabel('Category')
    st.pyplot(fig)

with tab2:
    # Pie chart
    fig2, ax2 = plt.subplots()
    category_counts = filtered_df['Main_Category'].value_counts()
    ax2.pie(category_counts.values, labels=category_counts.index, autopct='%1.1f%%', 
            startangle=90, colors=sns.color_palette('viridis', len(category_counts)))
    ax2.axis('equal')
    plt.title('Category Distribution')
    st.pyplot(fig2)

with tab3:

    st.write("**Top Reporters per Category**")
    
    
    for category in category_report.sort_values('Report_Count', ascending=False).itertuples():
        with st.expander(f"{category.Main_Category} ({category.Report_Count} reports)"):
            # Create reporter dataframe
            reporter_df = pd.DataFrame({
                'Reporter': list(category.Reporters.keys()),
                'Reports': list(category.Reporters.values())
            }).sort_values('Reports', ascending=False)
            
            # Calculate reporter percentage within category
            reporter_df['% of Category'] = (reporter_df['Reports'] / category.Report_Count * 100).round(1)
            
            st.dataframe(reporter_df.style.format({'% of Category': '{:.1f}%'}))

# Table of all complaints
st.subheader("All Complaints")
st.dataframe(filtered_df[['Date', 'Main_Category', 'Product', 'Reporter', 'Facility']].reset_index(drop=True))

# Export to Excel
st.subheader("Download Data")
st.write("Export filtered data and category summary")

def to_excel(dfs, sheet_names):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for i, df in enumerate(dfs):
            df.to_excel(writer, index=False, sheet_name=sheet_names[i])
    processed_data = output.getvalue()
    return processed_data

if st.button("📥 Generate Excel Report"):
    # Create two dataframes: summary and detailed complaints
    summary_df = category_report[['Main_Category', 'Report_Count', 'Percentage', 'Top_Reporters']]
    detailed_df = filtered_df[['Date', 'Report_Time', 'Main_Category', 'Product', 'Reporter', 'Facility']]
    
    excel_file = to_excel(
        [summary_df, detailed_df], 
        ['Category Summary', 'Detailed Complaints']
    )
    
    st.download_button(
        label="Download Excel Report",
        data=excel_file,
        file_name="complaints_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )