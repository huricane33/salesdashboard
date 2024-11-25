import streamlit as st
import pandas as pd
import numpy as np  # Ensure numpy is imported
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns

# Title of the dashboard
st.title("Comprehensive Sales Dashboard")

# File uploader
uploaded_file = st.file_uploader("Upload your Sales Data file (Excel format)", type=["xlsx"])

if uploaded_file:
    # Read the Excel file
    excel_data = pd.ExcelFile(uploaded_file)
    sheet_names = excel_data.sheet_names

    # Load data from the selected sheet
    sheet = st.selectbox("Select a sheet:", sheet_names)
    raw_data = excel_data.parse(sheet, header=[0, 1])

    # Flatten multi-level headers
    raw_data.columns = ['_'.join(col).strip() if col[1] else col[0] for col in raw_data.columns]

    # Dynamically detect the Kelompok Barang column
    category_column = next((col for col in raw_data.columns if "Kelompok" in col and "Barang" in col), None)
    if not category_column:
        st.error("Kelompok Barang column not found!")
    else:
        # Reshape data
        reshaped_data = pd.melt(
            raw_data,
            id_vars=[category_column],
            var_name="Month_Store",
            value_name="Sales"
        )

        # Split Month and Store
        reshaped_data[['Month', 'Store']] = reshaped_data['Month_Store'].str.extract(r'(\d+_\w+)_([a-zA-Z]+)')
        reshaped_data.dropna(subset=['Month', 'Store'], inplace=True)

        # Add Group column (e.g., BZR, GRC, FRS)
        reshaped_data['Group'] = reshaped_data[category_column].str[:3].str.upper()

        # Sidebar Filters
        st.sidebar.header("Filters")
        selected_groups = st.sidebar.multiselect("Select Groups (e.g., BZR, GRC, FRS):",
                                                 reshaped_data['Group'].unique(), reshaped_data['Group'].unique())
        selected_months = st.sidebar.multiselect("Select Months:", reshaped_data['Month'].unique(),
                                                 reshaped_data['Month'].unique())
        selected_stores = st.sidebar.multiselect("Select Stores:", reshaped_data['Store'].unique(),
                                                 reshaped_data['Store'].unique())

        # Updated Kelompok Barang Filter with search and comparison
        st.sidebar.header("Kelompok Barang Filters")
        selected_categories = st.sidebar.multiselect(
            "Search and Compare Kelompok Barang:",
            options=reshaped_data[category_column].unique(),
            default=[reshaped_data[category_column].unique()[0]]  # Default to the first item
        )

        # Filter data for all components
        filtered_data = reshaped_data[
            (reshaped_data['Group'].isin(selected_groups)) &
            (reshaped_data['Month'].isin(selected_months)) &
            (reshaped_data['Store'].isin(selected_stores))
        ]

        # Filtered Kelompok Barang data for comparison
        kelompok_data = reshaped_data[
            (reshaped_data[category_column].isin(selected_categories)) &
            (reshaped_data['Month'].isin(selected_months)) &
            (reshaped_data['Store'].isin(selected_stores))
        ]

        # Aggregations
        group_sales = filtered_data.groupby(['Group', 'Month'])['Sales'].sum().reset_index()
        store_comparison = filtered_data.groupby(['Month', 'Store'])['Sales'].sum().reset_index()

        # 1. Total Group Sales Overview
        st.header("Total Group Sales Overview")

        # Line chart for group sales trends
        group_sales_chart = px.line(
            group_sales,
            x="Month",
            y="Sales",
            color="Group",
            title="Total Sales by Group over Months",
            labels={"Sales": "Total Sales", "Month": "Month"}
        )
        group_sales_chart.update_traces(hovertemplate="Total Sales: %{y:,.0f}<br>Month: %{x}")
        st.plotly_chart(group_sales_chart)

        # Detailed table for group sales
        st.subheader("Detailed Group Sales by Month")
        group_sales_table = group_sales.pivot_table(
            values="Sales",
            index="Group",
            columns="Month",
            aggfunc="sum",
            fill_value=0
        )
        st.dataframe(group_sales_table)

        # Display sum table if more than one month is selected
        if len(selected_months) > 1:
            st.subheader("Sum Table for Total Group Sales")

            # Calculate total sales per group
            sum_table = group_sales_table.sum(axis=1).reset_index()
            sum_table.columns = ["Group", "Total Sales"]

            # Add Grand Total row
            grand_total = pd.DataFrame([["Grand Total", sum_table["Total Sales"].sum()]], columns=["Group", "Total Sales"])
            sum_table_with_total = pd.concat([sum_table, grand_total], ignore_index=True)

            # Display the sum table with total
            st.dataframe(sum_table_with_total)

        # 2. Month-to-Month Store Comparison
        st.header("Month-to-Month Comparison Between Stores")
        store_comparison_chart = px.bar(
            store_comparison,
            x="Month",
            y="Sales",
            color="Store",
            barmode="group",
            title="Store Sales Comparison",
            labels={"Sales": "Total Sales", "Month": "Month"}
        )
        store_comparison_chart.update_traces(hovertemplate="Total Sales: %{y:,.0f}<br>Month: %{x}")
        st.plotly_chart(store_comparison_chart)

        # 3. Month-to-Month Sales for All Kelompok Barang (Detailed View)
        st.header("Month-to-Month Sales for All Kelompok Barang (Detailed View)")

        # Pivot table for sales by Kelompok Barang, Store, and Month
        kelompok_month_comparison = filtered_data.pivot_table(
            values="Sales",
            index=[category_column, 'Store', 'Group'],
            columns="Month",
            aggfunc="sum",
            fill_value=0
        )

        # Calculate month-to-month changes
        kelompok_month_changes = kelompok_month_comparison.diff(axis=1)

        # Combine sales and changes into one DataFrame, assign names to levels
        detailed_combined_table = pd.concat(
            [kelompok_month_comparison, kelompok_month_changes],
            keys=["Sales", "Change"],
            axis=1,
            names=['Type', 'Month']
        )

        # Reset index
        detailed_combined_table.reset_index(inplace=True)

        # Flatten the MultiIndex columns
        detailed_combined_table.columns = [
            '_'.join([str(i) for i in col if str(i) != '']).strip('_') if isinstance(col, tuple) else col
            for col in detailed_combined_table.columns.values
        ]

        # Define index and data columns
        index_cols = [category_column, 'Store', 'Group']
        data_cols = [col for col in detailed_combined_table.columns if col not in index_cols]

        # Calculate total sales across all months for ranking
        sales_columns = [col for col in data_cols if "Sales_" in col]
        detailed_combined_table['Total Sales'] = detailed_combined_table[sales_columns].sum(axis=1)

        # Add rank based on total sales within each group
        detailed_combined_table['Rank'] = (
            detailed_combined_table.groupby('Group')['Total Sales']
            .rank(ascending=False, method='min')
        )

        # Rename columns for better readability
        detailed_combined_table.rename(
            columns={category_column: "Kelompok Barang", "Store": "Store", "Group": "Group"},
            inplace=True
        )

        # Sort by Group and Rank
        detailed_combined_table.sort_values(['Group', 'Rank'], inplace=True)

        # Display the table
        st.write("**Detailed Sales and Month-to-Month Changes by Kelompok Barang and Store**")
        st.dataframe(detailed_combined_table)

        # 4. Comparison of Selected Kelompok Barang
        st.header("Comparison of Selected Kelompok Barang")

        # Modify the comparison chart to use different colors for each Store
        comparison_chart = px.bar(
            kelompok_data,
            x="Month",
            y="Sales",
            color="Store",  # Changed from category_column to "Store"
            barmode="group",
            title="Sales Comparison for Selected Kelompok Barang",
            labels={"Sales": "Total Sales", "Month": "Month", "Store": "Store"}
        )
        comparison_chart.update_traces(hovertemplate="Total Sales: %{y:,.0f}<br>Month: %{x}<br>")
        st.plotly_chart(comparison_chart)

        # 5. Kelompok Barang Visualization
        st.header(f"Sales Visualization for Selected Kelompok Barang")
        pie_chart = px.pie(
            kelompok_data,
            names="Store",
            values="Sales",
            title="Sales Distribution for Selected Kelompok Barang"
        )
        pie_chart.update_traces(hovertemplate="Sales: %{value:,.0f}<br>Store: %{label}")
        st.plotly_chart(pie_chart)

        # 6. Trend Analysis for Selected Kelompok Barang
        st.header("Trend Analysis for Selected Kelompok Barang")

        if kelompok_data.empty:
            st.write("No data available for the selected Kelompok Barang and filters.")
        else:
            # Prepare data for trend analysis
            trend_data = kelompok_data.groupby(['Month', category_column])['Sales'].sum().reset_index()

            # Convert 'Month' to datetime using the format '%d_%b'
            trend_data['Month'] = pd.to_datetime(trend_data['Month'], format='%d_%b', errors='coerce')

            # Sort the data by 'Month'
            trend_data.sort_values('Month', inplace=True)

            # Convert 'Month' back to string format for display, e.g., '01_Jan'
            trend_data['Month_Display'] = trend_data['Month'].dt.strftime('%d_%b')

            # Create line chart
            trend_chart = px.line(
                trend_data,
                x='Month_Display',
                y='Sales',
                color=category_column,
                title='Sales Trend for Selected Kelompok Barang',
                labels={
                    'Sales': 'Total Sales',
                    'Month_Display': 'Month',
                    category_column: 'Kelompok Barang'
                },
                markers=True
            )

            trend_chart.update_layout(
                xaxis_title='Month',
                yaxis_title='Total Sales',
                legend_title='Kelompok Barang',
                title_font_size=20
            )

            trend_chart.update_traces(
                customdata=trend_data[[category_column]],
                hovertemplate="Total Sales: %{y:,.0f}<br>Month: %{x}<br>Kelompok Barang: %{customdata[0]}"
            )

            st.plotly_chart(trend_chart, use_container_width=True)
        # 7. Top/Bottom Performers
        st.header("Top/Bottom Performers")

        # Top Performers
        top_performers = filtered_data.groupby([category_column])['Sales'].sum().nlargest(10).reset_index()
        st.subheader("Top 10 Kelompok Barang")
        st.dataframe(top_performers)

        # Bottom Performers (dynamically skip 0 values)
        all_performers = filtered_data.groupby([category_column])['Sales'].sum().reset_index()
        bottom_performers = all_performers[all_performers['Sales'] > 0].nsmallest(10, 'Sales')  # Exclude 0 and get next lowest

        if not bottom_performers.empty:
            st.subheader("Bottom 10 Kelompok Barang")
            st.dataframe(bottom_performers)
        else:
            st.subheader("Bottom 10 Kelompok Barang")
            st.write("No bottom performers with non-zero sales.")

else:
    st.info("Please upload an Excel file to proceed.")