import streamlit as st
import pandas as pd
import numpy as np  # Ensure numpy is imported
import plotly.express as px

# Title of the dashboard
st.title("Comprehensive Sales Dashboard")

# -------------------- Main Code --------------------

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

        # Ensure 'Sales' column is numeric
        reshaped_data['Sales'] = pd.to_numeric(reshaped_data['Sales'], errors='coerce')
        reshaped_data.dropna(subset=['Sales'], inplace=True)

        # Sidebar Filters with Expander and Instructions
        st.sidebar.info("Use the filters below to customize the data displayed in the dashboard.")
        with st.sidebar.expander("Filters", expanded=True):
            st.header("General Filters")
            selected_groups = st.multiselect(
                "Select Groups (e.g., BZR, GRC, FRS):",
                options=reshaped_data['Group'].unique(),
                default=reshaped_data['Group'].unique(),
                help="Filter data by selecting one or more groups."
            )
            selected_months = st.multiselect(
                "Select Months:",
                options=reshaped_data['Month'].unique(),
                default=reshaped_data['Month'].unique(),
                help="Filter data by selecting one or more months."
            )
            selected_stores = st.multiselect(
                "Select Stores:",
                options=reshaped_data['Store'].unique(),
                default=reshaped_data['Store'].unique(),
                help="Filter data by selecting one or more stores."
            )

            st.header("Kelompok Barang Filters")
            selected_categories = st.multiselect(
                "Search and Compare Kelompok Barang:",
                options=reshaped_data[category_column].unique(),
                default=[reshaped_data[category_column].unique()[0]],  # Default to the first item
                help="Select Kelompok Barang to focus on specific product groups."
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

        # Convert 'Month' to datetime for sorting
        group_sales['Month'] = pd.to_datetime(group_sales['Month'], format='%d_%b', errors='coerce')
        group_sales.dropna(subset=['Month'], inplace=True)
        group_sales.sort_values('Month', inplace=True)
        group_sales['Month'] = group_sales['Month'].dt.strftime('%d_%b')

        store_comparison['Month'] = pd.to_datetime(store_comparison['Month'], format='%d_%b', errors='coerce')
        store_comparison.dropna(subset=['Month'], inplace=True)
        store_comparison.sort_values('Month', inplace=True)
        store_comparison['Month'] = store_comparison['Month'].dt.strftime('%d_%b')

        # Create a colorblind-friendly palette
        color_palette = px.colors.qualitative.Safe

        # Create tabs for different sections
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "Group Sales Overview",
            "Store Comparison",
            "Detailed View",
            "Kelompok Barang Comparison",
            "Kelompok Barang Visualization",
            "Sales Trend",
            "Top/Bottom Performers"
        ])

        # -------------------- 1. Group Sales Overview --------------------
        with tab1:
            st.header("Detailed Group Sales by Month")

            # Create pivot table for group sales
            group_sales_table = group_sales.pivot_table(
                values="Sales",
                index="Group",
                columns="Month",
                aggfunc="sum",
                fill_value=0
            )

            # Calculate month-to-month absolute differences
            group_sales_diff = group_sales_table.diff(axis=1)

            # Compute the total sales per month and append as a new row
            total_sales_row = group_sales_table.sum(axis=0)
            total_sales_row.name = 'Grand Total'
            group_sales_table_with_total = pd.concat([group_sales_table, total_sales_row.to_frame().T])

            # Similarly for differences
            total_diff_row = group_sales_diff.sum(axis=0)
            total_diff_row.name = 'Grand Total'
            group_sales_diff_with_total = pd.concat([group_sales_diff, total_diff_row.to_frame().T])

            # Checkbox to show percentage differences
            show_percentage = st.checkbox("Show Percentage Differences", value=False)

            if show_percentage:
                # Calculate month-to-month percentage differences
                group_sales_pct_change = group_sales_table.pct_change(axis=1) * 100

                # Compute percentage change for totals
                total_pct_change_row = group_sales_table_with_total.pct_change(axis=1).iloc[-1] * 100
                total_pct_change_row.name = 'Grand Total'
                group_sales_pct_change_with_total = pd.concat(
                    [group_sales_pct_change, total_pct_change_row.to_frame().T]
                )

                # Combine sales, absolute differences, and percentage differences into one DataFrame
                group_sales_combined = pd.concat(
                    [group_sales_table_with_total, group_sales_diff_with_total, group_sales_pct_change_with_total],
                    keys=["Sales", "Difference", "Percent Change"],
                    axis=1
                )

                # Reset column names and index
                group_sales_combined.columns.names = ['Type', 'Month']
                group_sales_combined.reset_index(inplace=True)

                # Flatten the MultiIndex columns
                group_sales_combined.columns = [
                    f"{col[0]}_{col[1]}" if col[0] != 'Group' else 'Group' for col in group_sales_combined.columns
                ]

                # Fill NaN and infinite values
                group_sales_combined.replace([np.inf, -np.inf], np.nan, inplace=True)
                group_sales_combined.fillna(0, inplace=True)

                # Function to format percentage changes with arrows
                def format_percentage_with_arrows(val):
                    try:
                        val_num = float(val)
                        arrow = '↑' if val_num > 0 else '↓' if val_num < 0 else ''
                        return f"{val_num:,.2f}% {arrow}"
                    except:
                        return val

                # Format the numbers
                for col in group_sales_combined.columns[1:]:
                    if "Percent Change" in col:
                        group_sales_combined[col] = group_sales_combined[col].apply(format_percentage_with_arrows)
                    else:
                        group_sales_combined[col] = group_sales_combined[col].apply(lambda x: f"{x:,.0f}")

                # Display the combined table
                st.dataframe(group_sales_combined)

            else:
                # Combine sales and differences into one DataFrame
                group_sales_combined = pd.concat(
                    [group_sales_table_with_total, group_sales_diff_with_total],
                    keys=["Sales", "Difference"],
                    axis=1
                )

                # Reset column names and index
                group_sales_combined.columns.names = ['Type', 'Month']
                group_sales_combined.reset_index(inplace=True)

                # Flatten the MultiIndex columns
                group_sales_combined.columns = [
                    f"{col[0]}_{col[1]}" if col[0] != 'Group' else 'Group' for col in group_sales_combined.columns
                ]

                # Fill NaN values
                group_sales_combined.fillna(0, inplace=True)

                # Format the numbers
                for col in group_sales_combined.columns[1:]:
                    group_sales_combined[col] = group_sales_combined[col].apply(lambda x: f"{x:,.0f}")

                # Display the combined table
                st.dataframe(group_sales_combined)

            # Line chart for group sales
            st.subheader("Total Sales by Group Over Months")
            group_sales_chart = px.line(
                group_sales,
                x="Month",
                y="Sales",
                color="Group",
                title="Total Sales by Group Over Months",
                labels={"Sales": "Total Sales", "Month": "Month"},
                markers=True,
                color_discrete_sequence=color_palette
            )
            group_sales_chart.update_layout(
                xaxis_title='Month',
                yaxis_title='Total Sales',
                legend_title='Group',
                hovermode='x unified'
            )
            group_sales_chart.update_traces(hovertemplate="Group: %{legendgroup}<br>Month: %{x}<br>Total Sales: %{y:,.0f}")
            st.plotly_chart(group_sales_chart, use_container_width=True)

        # -------------------- 2. Store Comparison --------------------
        with tab2:
            st.header("Month-to-Month Comparison Between Stores")

            # Bar chart for store sales comparison
            store_comparison_chart = px.bar(
                store_comparison,
                x="Month",
                y="Sales",
                color="Store",
                barmode="group",
                title="Store Sales Comparison",
                labels={"Sales": "Total Sales", "Month": "Month"},
                color_discrete_sequence=color_palette
            )
            store_comparison_chart.update_traces(hovertemplate="Store: %{legendgroup}<br>Month: %{x}<br>Total Sales: %{y:,.0f}")
            store_comparison_chart.update_layout(
                xaxis_title='Month',
                yaxis_title='Total Sales',
                legend_title='Store',
                hovermode='x unified'
            )
            st.plotly_chart(store_comparison_chart, use_container_width=True)

            # Option to show or hide the detailed data table
            show_table = st.checkbox("Show Detailed Data Table with Month-to-Month Changes", value=False)

            if show_table:
                st.subheader("Detailed Data with Month-to-Month Changes")

                # Create pivot table for sales data
                store_sales_table = store_comparison.pivot_table(
                    values="Sales",
                    index="Store",
                    columns="Month",
                    aggfunc="sum",
                    fill_value=0
                )

                # Ensure months are sorted chronologically
                # Convert columns to datetime for sorting
                store_sales_table.columns = pd.to_datetime(store_sales_table.columns, format='%d_%b', errors='coerce')
                store_sales_table = store_sales_table.reindex(sorted(store_sales_table.columns), axis=1)
                store_sales_table.columns = store_sales_table.columns.strftime('%d_%b')

                # Calculate month-to-month differences
                store_sales_diff = store_sales_table.diff(axis=1)

                # Combine sales and differences into one DataFrame
                store_sales_combined = pd.concat(
                    [store_sales_table, store_sales_diff],
                    keys=["Sales", "Difference"],
                    axis=1
                )

                # Flatten the MultiIndex columns
                store_sales_combined.columns.names = ['Type', 'Month']
                store_sales_combined.reset_index(inplace=True)
                store_sales_combined.columns = [
                    f"{col[0]}_{col[1]}" if col[0] != 'Store' else 'Store' for col in store_sales_combined.columns
                ]

                # Optionally format the numbers
                store_sales_combined.fillna(0, inplace=True)
                st.dataframe(
                    store_sales_combined.style.format("{:,.0f}", subset=store_sales_combined.columns[1:])
                )

        # -------------------- 3. Detailed View --------------------
        with tab3:
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

        # -------------------- 4. Kelompok Barang Comparison --------------------
        with tab4:
            st.header("Comparison of Selected Kelompok Barang")

            # Convert 'Month' to datetime for sorting
            kelompok_data['Month'] = pd.to_datetime(kelompok_data['Month'], format='%d_%b', errors='coerce')
            kelompok_data.dropna(subset=['Month'], inplace=True)
            kelompok_data.sort_values('Month', inplace=True)
            kelompok_data['Month_Display'] = kelompok_data['Month'].dt.strftime('%d_%b')

            # Modify the comparison chart to use different colors for each Store
            comparison_chart = px.bar(
                kelompok_data,
                x="Month_Display",
                y="Sales",
                color="Store",
                barmode="group",
                title="Sales Comparison for Selected Kelompok Barang",
                labels={"Sales": "Total Sales", "Month_Display": "Month", "Store": "Store"},
                color_discrete_sequence=color_palette
            )
            comparison_chart.update_traces(hovertemplate="Store: %{legendgroup}<br>Month: %{x}<br>Total Sales: %{y:,.0f}")
            comparison_chart.update_layout(
                xaxis_title='Month',
                yaxis_title='Total Sales',
                legend_title='Store',
                hovermode='x unified'
            )
            st.plotly_chart(comparison_chart, use_container_width=True)

        # -------------------- 5. Kelompok Barang Visualization --------------------
        with tab5:
            st.header(f"Sales Visualization for Selected Kelompok Barang")
            pie_chart = px.pie(
                kelompok_data,
                names="Store",
                values="Sales",
                title="Sales Distribution for Selected Kelompok Barang",
                color_discrete_sequence=color_palette
            )
            pie_chart.update_traces(hovertemplate="Store: %{label}<br>Sales: %{value:,.0f} (%{percent})")
            st.plotly_chart(pie_chart, use_container_width=True)

        # -------------------- 6. Sales Trend --------------------
        with tab6:
            st.header("Sales Trend for Selected Kelompok Barang by Store")

            if kelompok_data.empty:
                st.write("No data available for the selected Kelompok Barang and filters.")
            else:
                # Prepare data for trend analysis
                trend_data = kelompok_data.groupby(['Month', 'Store', category_column])['Sales'].sum().reset_index()

                # Convert 'Month' to datetime for sorting
                trend_data['Month'] = pd.to_datetime(trend_data['Month'], format='%d_%b', errors='coerce')
                trend_data.dropna(subset=['Month'], inplace=True)
                trend_data.sort_values('Month', inplace=True)
                trend_data['Month_Display'] = trend_data['Month'].dt.strftime('%d_%b')

                # Create line chart
                trend_chart = px.line(
                    trend_data,
                    x='Month_Display',
                    y='Sales',
                    color='Store',
                    line_group='Store',
                    facet_col=category_column,
                    facet_col_wrap=2,
                    title='Sales Trend for Selected Kelompok Barang by Store',
                    labels={
                        'Sales': 'Total Sales',
                        'Month_Display': 'Month',
                        'Store': 'Store',
                        category_column: 'Kelompok Barang'
                    },
                    markers=True,
                    color_discrete_sequence=color_palette
                )

                trend_chart.update_layout(
                    xaxis_title='Month',
                    yaxis_title='Total Sales',
                    legend_title='Store',
                    title_font_size=20,
                    hovermode='x unified',
                    height=600  # Adjust the height as needed
                )

                trend_chart.update_traces(
                    customdata=trend_data[[category_column]],
                    hovertemplate="Store: %{legendgroup}<br>Month: %{x}<br>Total Sales: %{y:,.0f}<br>Kelompok Barang: %{customdata[0]}"
                )

                st.plotly_chart(trend_chart, use_container_width=True)

        # -------------------- 7. Top/Bottom Performers --------------------
        with tab7:
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