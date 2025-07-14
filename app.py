import streamlit as st
import pandas as pd
import plotly.express as px

# Load and cache data
@st.cache_data
def load_data():
    df = pd.read_csv("Supermart Grocery Sales - Retail Analytics Dataset.csv")
    df['Order Date'] = pd.to_datetime(df['Order Date'], errors='coerce')
    df['Year'] = df['Order Date'].dt.year
    df['Month'] = df['Order Date'].dt.strftime('%B')
    return df

df = load_data()

# -------------------------------------------
# Sidebar filters
# -------------------------------------------
st.sidebar.title("Filter Options")

regions = df['Region'].dropna().unique().tolist()
categories = df['Category'].dropna().unique().tolist()
years = sorted(df['Year'].dropna().unique().tolist())

selected_regions = st.sidebar.multiselect("Select Region(s)", regions, default=regions)
selected_years = st.sidebar.multiselect("Select Year(s)", years, default=years)
selected_categories = st.sidebar.multiselect("Select Category(s)", categories, default=categories)

# Filter the dataframe
filtered_df = df[
    (df['Region'].isin(selected_regions)) &
    (df['Year'].isin(selected_years)) &
    (df['Category'].isin(selected_categories))
]

# -------------------------------------------
# Dashboard Title
# -------------------------------------------
st.title("🛒 Supermart Sales Dashboard")

# -------------------------------------------
# Region-wise Sales Chart
# -------------------------------------------
region_sales = filtered_df.groupby('Region')['Sales'].sum().reset_index()
fig_region = px.bar(region_sales, x='Region', y='Sales',
                    title="Region-wise Sales",
                    color='Sales', color_continuous_scale='Blues')
st.plotly_chart(fig_region, use_container_width=True)

# -------------------------------------------
# Monthly Sales Trend
# -------------------------------------------
monthly_sales = filtered_df.groupby('Month')['Sales'].sum().reindex(
    ['January', 'February', 'March', 'April', 'May', 'June',
     'July', 'August', 'September', 'October', 'November', 'December']
).reset_index()

fig_month = px.line(monthly_sales, x='Month', y='Sales',
                    title="Monthly Sales Trend", markers=True)
st.plotly_chart(fig_month, use_container_width=True)

# -------------------------------------------
# Category-wise Sales Pie Chart
# -------------------------------------------
category_sales = filtered_df.groupby('Category')['Sales'].sum().reset_index()
fig_cat = px.pie(category_sales, names='Category', values='Sales',
                 title="Category-wise Sales Distribution",
                 hole=0.4)
st.plotly_chart(fig_cat, use_container_width=True)

# -------------------------------------------
# Optional: Show raw data
# -------------------------------------------
if st.checkbox("Show Raw Data"):
    st.subheader("Filtered Data Table")
    st.write(filtered_df)

# -------------------------------------------
# 📥 Download filtered data as Excel
# -------------------------------------------
import io

output = io.BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    filtered_df.to_excel(writer, index=False, sheet_name='FilteredSales')
    # ❌ DO NOT call writer.save() here
    # It is handled automatically by the context manager
processed_data = output.getvalue()

st.download_button(
    label="📥 Download Filtered Data as Excel",
    data=processed_data,
    file_name='filtered_sales_data.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)
