import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime

# Set Streamlit page configuration
st.set_page_config(layout="wide", page_title="Comprehensive Sales & Stock Dashboard")

# Title of the Dashboard
st.title("Comprehensive Sales & Stock Dashboard")

# File uploader in the main area
uploaded_file = st.file_uploader("Upload your Sales Data file (Excel format)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Load and process data
        with st.spinner('Loading and processing data...'):
            # Read the Excel file and load the first sheet automatically
            excel_data = pd.ExcelFile(uploaded_file)
            first_sheet = excel_data.sheet_names[0]
            raw_data = pd.read_excel(uploaded_file, sheet_name=first_sheet)

            # Define required columns
            required_cols = ["Grouping", "Penjualan", "HPP", "Gross Margin", "Store Name", "Month", "year", "Stock Value"]
            raw_data.columns = raw_data.columns.str.strip()  # Remove any leading/trailing spaces

            # Check for required columns (case-insensitive)
            raw_data_lower = raw_data.columns.str.lower()
            required_cols_lower = [col.lower() for col in required_cols]
            if not all(col in raw_data_lower for col in required_cols_lower):
                st.error(f"The uploaded sheet must contain the following columns: {required_cols}")
            else:
                # Rename columns to standard names (case-insensitive)
                rename_dict = {col.lower(): col for col in raw_data.columns}
                raw_data.rename(columns=rename_dict, inplace=True)

                # Convert numeric columns
                numeric_cols = ["Penjualan", "HPP", "Gross Margin", "Stock Value"]
                for col in numeric_cols:
                    # Remove thousand separators '.' and replace decimal ',' with '.' if necessary
                    raw_data[col] = raw_data[col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                    raw_data[col] = pd.to_numeric(raw_data[col], errors='coerce')

                # Drop rows with invalid numeric values
                raw_data.dropna(subset=numeric_cols, inplace=True)

                # Calculate Margin %
                raw_data['Margin %'] = (raw_data['Gross Margin'] / raw_data['Penjualan']) * 100

                # Create a Date column
                try:
                    raw_data['Date'] = pd.to_datetime(raw_data['year'].astype(int).astype(str) + '-' + raw_data['Month'],
                                                     format='%Y-%B', errors='coerce')
                    # If parsing failed (all NaT), try abbreviated month names
                    if raw_data['Date'].isna().all():
                        raw_data['Date'] = pd.to_datetime(raw_data['year'].astype(int).astype(str) + '-' + raw_data['Month'],
                                                         format='%Y-%b', errors='coerce')
                except Exception as e:
                    st.error(f"Error parsing dates: {e}")
                    raw_data['Date'] = pd.NaT

                # Drop rows with invalid Date
                raw_data.dropna(subset=['Date'], inplace=True)

                # Sort raw_data by Date
                raw_data.sort_values('Date', inplace=True)

                # If a Group column isn't present, derive it (e.g., first 3 chars of Grouping)
                if 'Group' not in raw_data.columns:
                    raw_data['Group'] = raw_data['Grouping'].astype(str).str[:3].str.upper()

                # Combine GRC and FRS into GRC+FRS
                raw_data['Group'] = raw_data['Group'].replace({'GRC': 'GRC+FRS', 'FRS': 'GRC+FRS'})
                # Filter only GRC+FRS and BZR
                raw_data = raw_data[raw_data['Group'].isin(['GRC+FRS', 'BZR'])]

                # Create a Month_Display column
                raw_data['Month_Display'] = raw_data['Date'].dt.strftime('%b %Y')

        st.success('Data loaded and processed successfully!')

        # Sidebar Filters
        st.sidebar.header("Filters")
        with st.sidebar.expander("General Filters", expanded=True):
            # Divisions Filter (Now only GRC+FRS and BZR)
            selected_groups = st.multiselect(
                "Select Divisions (GRC+FRS, BZR):",
                options=sorted(raw_data['Group'].unique()),
                default=sorted(raw_data['Group'].unique()),
                help="Choose one or more divisions to filter the sales data accordingly."
            )

            # Years Filter
            selected_years = st.multiselect(
                "Select Years:",
                options=sorted(raw_data['year'].unique()),
                default=sorted(raw_data['year'].unique()),
                help="Select the years you want to include in the analysis."
            )

            # Months Filter
            unique_months = sorted(raw_data['Month'].dropna().unique(), key=lambda x: datetime.strptime(x, '%B').month)
            selected_months = st.multiselect(
                "Select Months:",
                options=unique_months,
                default=unique_months,  # By default, all months are selected
                help="Filter the data by selecting specific months."
            )

            # Stores Filter
            selected_stores = st.multiselect(
                "Select Stores:",
                options=sorted(raw_data['Store Name'].unique()),
                default=sorted(raw_data['Store Name'].unique()),
                help="Choose the stores you want to include in the dashboard."
            )

        with st.sidebar.expander("Grouping Filters", expanded=True):
            unique_categories = sorted(raw_data['Grouping'].unique())
            default_cat = [unique_categories[0]] if unique_categories else []
            selected_categories = st.multiselect(
                "Search and Compare Grouping:",
                options=unique_categories,
                default=default_cat,
                help="Select one or more 'Grouping' categories to compare their sales performance."
            )

        # Apply General Filters
        filtered_data = raw_data[
            (raw_data['Group'].isin(selected_groups)) &
            (raw_data['year'].isin(selected_years)) &
            (raw_data['Month'].isin(selected_months)) &
            (raw_data['Store Name'].isin(selected_stores))
        ].copy()

        # Sort filtered_data by Date
        filtered_data.sort_values('Date', inplace=True)

        # Apply Grouping Filters
        kelompok_data = raw_data[
            (raw_data['Grouping'].isin(selected_categories)) &
            (raw_data['year'].isin(selected_years)) &
            (raw_data['Month'].isin(selected_months)) &
            (raw_data['Store Name'].isin(selected_stores))
        ].copy()

        # Sort kelompok_data by Date
        kelompok_data.sort_values('Date', inplace=True)

        if filtered_data.empty:
            st.warning("No data available after applying the selected filters.")
        else:
            # Aggregations
            group_sales = filtered_data.groupby(['Group', 'Date'])['Penjualan'].sum().reset_index()
            store_comparison = filtered_data.groupby(['Date', 'Store Name'])['Penjualan'].sum().reset_index()

            # Sort by date
            group_sales.sort_values('Date', inplace=True)
            store_comparison.sort_values('Date', inplace=True)

            # Add Month_Display for aggregated data
            group_sales['Month_Display'] = group_sales['Date'].dt.strftime('%b %Y')
            store_comparison['Month_Display'] = store_comparison['Date'].dt.strftime('%b %Y')

            if not kelompok_data.empty:
                kelompok_data['Date'] = pd.to_datetime(kelompok_data['Date'], errors='coerce')
                kelompok_data.dropna(subset=['Date'], inplace=True)
                kelompok_data.sort_values('Date', inplace=True)
                kelompok_data['Month_Display'] = kelompok_data['Date'].dt.strftime('%b %Y')

            # Define a colorblind-friendly palette
            color_palette = px.colors.qualitative.Safe

            # Create Tabs
            tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
                "Group Sales Overview",
                "Store Comparison",
                "Detailed View per Category",
                "Grouping BarChart",
                "Grouping PieChart",
                "Sales Trend",
                "Top/Bottom Performers",
                "Gross Margin Analysis",
                "Stock Value Analysis"
            ])

            # -------------------- 1. Group Sales Overview (Tab 1) --------------------
            with tab1:
                st.header("Detailed Group Sales by Month")
                st.markdown("""
                    This section provides a detailed overview of sales by group for each month.
                    You can toggle between viewing absolute sales figures, percentage differences, 
                    and contributions to the grand total.
                """)

                if group_sales.empty:
                    st.write("No Group Sales data available.")
                else:
                    # Create a pivot table for group sales
                    group_sales_table = group_sales.pivot_table(
                        values="Penjualan",
                        index="Group",
                        columns="Month_Display",
                        aggfunc="sum",
                        fill_value=0
                    )

                    # Ensure columns are ordered chronologically
                    group_sales_table = group_sales_table.reindex(
                        sorted(group_sales_table.columns, key=lambda x: datetime.strptime(x, '%b %Y')),
                        axis=1
                    )

                    # Calculate differences
                    group_sales_diff = group_sales_table.diff(axis=1)

                    # Compute total row
                    total_sales_row = group_sales_table.sum(axis=0)
                    total_sales_row.name = 'Grand Total'
                    group_sales_table_with_total = pd.concat([group_sales_table, total_sales_row.to_frame().T])

                    total_diff_row = group_sales_diff.sum(axis=0)
                    total_diff_row.name = 'Grand Total'
                    group_sales_diff_with_total = pd.concat([group_sales_diff, total_diff_row.to_frame().T])

                    show_percentage = st.checkbox("Show Percentage Differences", value=False, key='group_pct')
                    show_contribution = st.checkbox("Show Contribution to Grand Total", value=False,
                                                    key='group_contribution')

                    if show_contribution:
                        # Contribution to grand total
                        grand_total_sales = group_sales_table_with_total.loc['Grand Total']
                        group_contribution = (group_sales_table_with_total.div(grand_total_sales) * 100).round(2)
                        group_contribution = group_contribution.reset_index()

                        # Identify numeric columns excluding 'Group'
                        numeric_cols = group_contribution.select_dtypes(include=['number']).columns.tolist()

                        # Ensure all numeric columns are of float type
                        for col in numeric_cols:
                            group_contribution[col] = pd.to_numeric(group_contribution[col], errors='coerce')

                        # Optionally, handle NaN values resulting from conversion
                        group_contribution.fillna(0, inplace=True)

                        # Apply percentage formatting only to numeric columns
                        group_contribution_style = group_contribution.style.format({
                            col: "{:.2f}%" for col in numeric_cols
                        })

                        st.dataframe(group_contribution_style)

                    elif show_percentage:
                        # Percentage change
                        group_sales_pct_change = group_sales_table.pct_change(axis=1) * 100
                        total_pct_change_row = group_sales_table_with_total.pct_change(axis=1).iloc[-1] * 100
                        total_pct_change_row.name = 'Grand Total'
                        group_sales_pct_change_with_total = pd.concat(
                            [group_sales_pct_change, total_pct_change_row.to_frame().T]
                        )

                        group_sales_combined = pd.concat(
                            [group_sales_table_with_total, group_sales_diff_with_total,
                             group_sales_pct_change_with_total],
                            keys=["Sales", "Difference", "Percent Change"],
                            axis=1
                        )

                        group_sales_combined.columns.names = ['Type', 'Month']
                        group_sales_combined.reset_index(inplace=True)

                        group_sales_combined.columns = [
                            f"{col[0]}_{col[1]}" if col[0] != 'Group' else 'Group' for col in
                            group_sales_combined.columns
                        ]


                        def format_percentage_with_arrows(val):
                            try:
                                val_num = float(val)
                                arrow = '↑' if val_num > 0 else '↓' if val_num < 0 else ''
                                return f"{val_num:,.2f}% {arrow}"
                            except:
                                return val


                        for col in group_sales_combined.columns[1:]:
                            if "Percent Change" in col:
                                group_sales_combined[col] = group_sales_combined[col].apply(
                                    lambda x: format_percentage_with_arrows(x) if x != 0 else x
                                )
                            else:
                                try:
                                    group_sales_combined[col] = group_sales_combined[col].apply(
                                        lambda x: f"{float(x):,.0f}" if x != 0 else x
                                    )
                                except:
                                    pass

                        st.dataframe(group_sales_combined)

                    else:
                        group_sales_combined = pd.concat(
                            [group_sales_table_with_total, group_sales_diff_with_total],
                            keys=["Sales", "Difference"],
                            axis=1
                        )

                        group_sales_combined.columns.names = ['Type', 'Month']
                        group_sales_combined.reset_index(inplace=True)

                        group_sales_combined.columns = [
                            f"{col[0]}_{col[1]}" if col[0] != 'Group' else 'Group' for col in
                            group_sales_combined.columns
                        ]

                        group_sales_combined.fillna(0, inplace=True)
                        for col in group_sales_combined.columns[1:]:
                            group_sales_combined[col] = group_sales_combined[col].apply(
                                lambda x: f"{float(x):,.0f}" if x != 0 else x)

                        st.dataframe(group_sales_combined)

                    # Line chart for group sales
                    if not group_sales.empty:
                        st.subheader("Total Sales by Group Over Months")
                        line_data = group_sales.copy()
                        line_data['Month_Display'] = pd.Categorical(
                            line_data['Month_Display'],
                            categories=sorted(line_data['Month_Display'].unique(),
                                              key=lambda x: datetime.strptime(x, '%b %Y')),
                            ordered=True
                        )

                        fig = px.line(
                            line_data,
                            x="Month_Display",
                            y="Penjualan",
                            color="Group",
                            title="Total Sales by Group Over Months",
                            labels={"Penjualan": "Total Sales", "Month_Display": "Month"},
                            color_discrete_sequence=px.colors.qualitative.Safe
                        )

                        fig.update_traces(mode='lines+markers')
                        fig.update_layout(
                            xaxis_title='Month',
                            yaxis_title='Total Sales',
                            legend_title='Group',
                            hovermode='x unified'
                        )
                        fig.update_traces(
                            hovertemplate="Group: %{legendgroup}<br>Month: %{x}<br>Total Sales: %{y:,.0f}"
                        )

                        st.plotly_chart(fig, use_container_width=True)

            # -------------------- 2. Store Comparison (Tab 2) --------------------
            with tab2:
                st.header("Month-to-Month Comparison Between Stores")
                st.markdown("""
                    Compare sales performance across different stores on a monthly basis.
                    This visualization helps in identifying top-performing stores and tracking their growth.
                """)

                if store_comparison.empty or 'Month_Display' not in store_comparison.columns:
                    st.write("No Store Comparison data available.")
                else:
                    # Bar chart for store comparison
                    fig_store = px.bar(
                        store_comparison,
                        x="Month_Display",
                        y="Penjualan",
                        color="Store Name",
                        barmode="group",
                        title="Store Sales Comparison",
                        labels={"Penjualan": "Total Sales", "Month_Display": "Month"},
                        color_discrete_sequence=px.colors.qualitative.Safe
                    )
                    fig_store.update_traces(hovertemplate="Month: %{x}<br>Total Sales: %{y:,.0f}")
                    fig_store.update_layout(
                        xaxis_title='Month',
                        yaxis_title='Total Sales',
                        legend_title='Store Name',
                        hovermode='x unified'
                    )
                    st.plotly_chart(fig_store, use_container_width=True)

                    # Checkbox to show the detailed data table
                    show_table = st.checkbox("Show Detailed Data Table with Month-to-Month Changes", value=False,
                                             key='store_table')
                    if show_table:
                        st.subheader("Detailed Data with Month-to-Month Changes")

                        # Pivot table for sales by store and month
                        pivot_store = store_comparison.pivot_table(
                            values="Penjualan",
                            index="Store Name",
                            columns="Month_Display",
                            aggfunc="sum",
                            fill_value=0
                        )

                        # Ensure columns are ordered chronologically
                        pivot_store = pivot_store.reindex(
                            sorted(pivot_store.columns, key=lambda x: datetime.strptime(x, '%b %Y')),
                            axis=1
                        )

                        # Calculate month-to-month differences
                        if len(pivot_store.columns) > 1:
                            store_diff = pivot_store.diff(axis=1)
                            include_difference = True
                        else:
                            store_diff = pd.DataFrame(index=pivot_store.index)
                            include_difference = False

                        # Add Grand Total row
                        sales_total = pivot_store.sum(axis=0)
                        pivot_store_with_total = pd.concat(
                            [pivot_store, pd.DataFrame([sales_total], index=["Grand Total"])]
                        )

                        if include_difference:
                            diff_total = store_diff.sum(axis=0)
                            store_diff_with_total = pd.concat(
                                [store_diff, pd.DataFrame([diff_total], index=["Grand Total"])]
                            )

                            # Combine sales and differences
                            combined_store = pd.concat(
                                [pivot_store_with_total, store_diff_with_total],
                                keys=["Sales", "Difference"],
                                axis=1
                            )
                        else:
                            # If no differences, just display sales
                            combined_store = pivot_store_with_total.copy()
                            combined_store.columns = pd.MultiIndex.from_arrays(
                                [["Sales"] * len(combined_store.columns), combined_store.columns],
                                names=["Type", "Month"]
                            )

                        combined_store.columns.names = ['Type', 'Month']
                        combined_store.reset_index(inplace=True)

                        # Flatten MultiIndex columns
                        combined_store.columns = [
                            f"{col[0]}_{col[1]}" if col[0] != 'Store Name' else 'Store Name'
                            for col in combined_store.columns
                        ]

                        # Use Styler to format numbers with thousand separators
                        format_dict = {
                            col: "{:,.0f}" for col in combined_store.columns if
                            col != 'Store Name' and col.startswith(('Sales_', 'Difference_'))
                        }
                        combined_store_style = combined_store.style.format(format_dict)

                        # Display the styled DataFrame
                        st.dataframe(combined_store_style)



            # -------------------- 3. Detailed View per Category (Tab 3) --------------------
            with tab3:
                st.header("Month-to-Month Sales for All Grouping (Detailed View per Category)")
                st.markdown("""
                    Dive deep into the sales data for each grouping across different stores and divisions.
                    This detailed view allows for comprehensive analysis and comparison.
                """)

                if filtered_data.empty:
                    st.write("No data available for Detailed View per Category.")
                else:
                    # Pivot table for sales by Grouping, Store, and Group
                    detail_pivot = filtered_data.pivot_table(
                        values="Penjualan",
                        index=["Grouping", "Store Name", "Group"],
                        columns="Date",
                        aggfunc="sum",
                        fill_value=0
                    )

                    # Sort columns by date
                    detail_pivot = detail_pivot.reindex(sorted(detail_pivot.columns), axis=1)

                    # Convert columns back to Month_Display
                    month_display_cols = [d.strftime('%b %Y') for d in detail_pivot.columns]
                    detail_pivot.columns = month_display_cols

                    # Check how many months we have
                    if len(detail_pivot.columns) < 2:
                        # Only one month, no change or pct change possible
                        include_changes = False
                        include_pct_changes = False
                    else:
                        # Calculate changes (value difference)
                        detail_changes = detail_pivot.diff(axis=1)
                        include_changes = not detail_changes.isna().all().all()

                        # Calculate percent changes
                        detail_pct_change = detail_pivot.pct_change(axis=1) * 100
                        include_pct_changes = not detail_pct_change.isna().all().all()

                    # Always include Sales
                    keys = ["Sales"]
                    all_dfs = [detail_pivot]

                    if include_changes:
                        all_dfs.append(detail_changes)
                        keys.append("Change")
                    if include_pct_changes:
                        all_dfs.append(detail_pct_change)
                        keys.append("Percent Change")

                    detailed_combined_table = pd.concat(all_dfs, keys=keys, axis=1, names=['Type', 'Month'])
                    detailed_combined_table.reset_index(inplace=True)

                    # Flatten MultiIndex columns
                    detailed_combined_table.columns = [
                        '_'.join([str(i) for i in col if str(i) != '']).strip('_') if isinstance(col, tuple) else col
                        for col in detailed_combined_table.columns.values
                    ]

                    # Calculate total sales for ranking
                    sales_cols = [c for c in detailed_combined_table.columns if c.startswith("Sales_")]
                    # Ensure numeric types
                    for col in sales_cols:
                        detailed_combined_table[col] = pd.to_numeric(detailed_combined_table[col], errors='coerce')
                    detailed_combined_table['Total Sales'] = detailed_combined_table[sales_cols].sum(axis=1)

                    # Rank by group
                    detailed_combined_table['Rank'] = detailed_combined_table.groupby('Group')['Total Sales'].rank(
                        ascending=False, method='min')

                    # Sort by Group and Rank
                    detailed_combined_table.sort_values(['Group', 'Rank'], inplace=True)

                    # Format numeric columns using Styler
                    style_dict_detail = {col: "{:,.0f}" for col in detailed_combined_table.columns if
                                         col.startswith('Sales_') or col.startswith('Change_') or
                                         col == 'Total Sales'}
                    style_dict_detail.update({
                        col: "{:.2f}%" for col in detailed_combined_table.columns if col.startswith('Percent Change_')
                    })
                    style_dict_detail['Group'] = '{}'
                    style_dict_detail['Store Name'] = '{}'

                    detailed_combined_style = detailed_combined_table.style.format(style_dict_detail)

                    st.dataframe(detailed_combined_style)

            # -------------------- 4. Grouping BarChart (Tab 4) --------------------
            with tab4:
                st.header("Grouping Comparison")
                st.markdown("""
                    Compare sales performance across different 'Grouping' categories.
                    Visualizations adjust based on the number of categories selected.
                """)

                if kelompok_data.empty or 'Month_Display' not in kelompok_data.columns:
                    st.write("No data available for the selected Grouping.")
                else:
                    if len(selected_categories) == 1:
                        single_cat = selected_categories[0]
                        st.subheader(f"Sales Comparison for {single_cat}")

                        comparison_chart = px.bar(
                            kelompok_data,
                            x="Month_Display",
                            y="Penjualan",
                            color="Store Name",
                            barmode="group",
                            title=f"Sales Comparison for {single_cat}",
                            labels={"Penjualan": "Total Sales", "Month_Display": "Month", "Store Name": "Store"},
                            color_discrete_sequence=color_palette
                        )
                        comparison_chart.update_traces(hovertemplate="Month: %{x}<br>Total Sales: %{y:,.0f}")
                        comparison_chart.update_layout(
                            xaxis_title='Month',
                            yaxis_title='Total Sales',
                            legend_title='Store',
                            hovermode='x unified'
                        )
                        st.plotly_chart(comparison_chart, use_container_width=True)
                    else:
                        st.subheader("Sales Comparison for Selected Grouping")

                        comparison_chart = px.bar(
                            kelompok_data,
                            x="Month_Display",
                            y="Penjualan",
                            color="Store Name",
                            barmode="group",
                            facet_col="Grouping",
                            facet_col_wrap=2,
                            title="Sales Comparison for Selected Grouping",
                            labels={"Penjualan": "Total Sales", "Month_Display": "Month", "Store Name": "Store",
                                    "Grouping": "Grouping"},
                            color_discrete_sequence=color_palette
                        )
                        comparison_chart.update_traces(hovertemplate="Month: %{x}<br>Total Sales: %{y:,.0f}")
                        comparison_chart.update_layout(
                            xaxis_title='Month',
                            yaxis_title='Total Sales',
                            legend_title='Store',
                            hovermode='x unified',
                            title_font_size=20,
                            height=600
                        )
                        st.plotly_chart(comparison_chart, use_container_width=True)

            # -------------------- 5. Grouping PieChart (Tab 5) --------------------
            with tab5:
                st.header("Comparison of Grouping by PieChart")
                st.markdown("""
                    Visualize the sales distribution of selected 'Grouping' across different stores using a pie chart.
                """)

                if kelompok_data.empty:
                    st.write("No data available for the selected Grouping.")
                else:
                    pie_chart = px.pie(
                        kelompok_data,
                        names="Store Name",
                        values="Penjualan",
                        title="Sales Distribution for Selected Grouping",
                        color_discrete_sequence=color_palette
                    )
                    pie_chart.update_traces(hovertemplate="Store: %{label}<br>Sales: %{value:,.0f} (%{percent})")
                    st.plotly_chart(pie_chart, use_container_width=True)

            # -------------------- 6. Sales Trend (Tab 6) --------------------
            with tab6:
                st.header("Sales Trend for Selected Grouping by Store")
                st.markdown("""
                    Analyze the sales trends over time for selected 'Grouping' across different stores.
                    Faceted line charts provide a clear view of each category's performance.
                """)

                if kelompok_data.empty or 'Month_Display' not in kelompok_data.columns:
                    st.write("No data available for the selected Grouping.")
                else:
                    trend_data = kelompok_data.groupby(['Date', 'Store Name', 'Grouping'])['Penjualan'].sum().reset_index()
                    trend_data.sort_values('Date', inplace=True)
                    if not trend_data.empty:
                        trend_data['Month_Display'] = trend_data['Date'].dt.strftime('%b %Y')

                    if trend_data.empty or 'Month_Display' not in trend_data.columns:
                        st.write("No data to display for trend.")
                    else:
                        trend_chart = px.line(
                            trend_data,
                            x='Month_Display',
                            y='Penjualan',
                            color='Store Name',
                            facet_col='Grouping',
                            facet_col_wrap=2,
                            title='Sales Trend for Selected Grouping by Store',
                            labels={'Penjualan': 'Total Sales', 'Month_Display': 'Month', 'Store Name': 'Store', 'Grouping': 'Grouping'},
                            color_discrete_sequence=color_palette
                        )

                        # Add markers to the trend chart
                        trend_chart.update_traces(mode='lines+markers')

                        trend_chart.update_layout(
                            xaxis_title='Month',
                            yaxis_title='Total Sales',
                            legend_title='Store',
                            title_font_size=20,
                            hovermode='x unified',
                            height=600
                        )

                        trend_chart.update_traces(
                            hovertemplate="Month: %{x}<br>Total Sales: %{y:,.0f}<br>Grouping: %{legendgroup}"
                        )

                        st.plotly_chart(trend_chart, use_container_width=True)

            # -------------------- 7. Top/Bottom Performers (Tab 7) --------------------
            with tab7:
                st.header("Top/Bottom Performers")
                st.markdown("""
                    Identify the top 10 and bottom 10 performing 'Grouping' based on total sales.
                    This helps in recognizing high-performing categories and those that may need attention.
                """)

                all_performers = filtered_data.groupby('Grouping')['Penjualan'].sum().reset_index()
                top_performers = all_performers.nlargest(10, 'Penjualan')
                bottom_performers = all_performers[all_performers['Penjualan'] > 0].nsmallest(10, 'Penjualan')

                st.subheader("Top 10 Grouping")
                # Use Styler for formatting
                top_performers_style = top_performers.style.format({
                    'Penjualan': "{:,.0f}"
                })
                st.dataframe(top_performers_style)

                st.subheader("Bottom 10 Grouping")
                if bottom_performers.empty:
                    st.write("No bottom performers with non-zero sales.")
                else:
                    # Use Styler for formatting
                    bottom_performers_style = bottom_performers.style.format({
                        'Penjualan': "{:,.0f}"
                    })
                    st.dataframe(bottom_performers_style)

            # -------------------- 8. Gross Margin Analysis (Tab 8) --------------------
            with tab8:
                st.header("Gross Margin Analysis")
                st.markdown("""
                    Analyze the gross margin to understand profitability across divisions and stores.
                    This section includes total gross margin, average margin percentage, and growth rates.
                """)

                # Recalculate Gross Margin to ensure correctness
                filtered_data['Gross Margin'] = filtered_data['Penjualan'] - filtered_data['HPP']

                # Combine 'GRC' and 'FRS' into 'GRC+FRS' if not already combined
                filtered_data['Group'] = filtered_data['Group'].replace({'GRC': 'GRC+FRS', 'FRS': 'GRC+FRS'})

                # Total Gross Margin and Correct Average Margin %
                if not filtered_data.empty:
                    total_gross_margin = filtered_data['Gross Margin'].sum()
                    total_penjualan = filtered_data['Penjualan'].sum()
                    avg_margin_percent = (total_gross_margin / total_penjualan) * 100 if total_penjualan != 0 else 0
                else:
                    total_gross_margin = 0
                    avg_margin_percent = 0

                # Display Metrics
                col1, col2 = st.columns(2)
                col1.metric("Total Gross Margin", f"{total_gross_margin:,.0f}")
                col2.metric("Average Margin %", f"{avg_margin_percent:.2f}%")

                # Additional KPI: Gross Margin Growth Rate
                latest_month = filtered_data['Date'].max()
                previous_month = latest_month - pd.DateOffset(months=1)

                latest_gm = filtered_data[filtered_data['Date'] == latest_month]['Gross Margin'].sum()
                previous_gm = filtered_data[filtered_data['Date'] == previous_month]['Gross Margin'].sum()

                if previous_gm > 0:
                    gm_growth_rate = ((latest_gm - previous_gm) / previous_gm) * 100
                else:
                    gm_growth_rate = np.nan

                if not np.isnan(gm_growth_rate):
                    st.metric("Gross Margin Growth Rate", f"{gm_growth_rate:.2f}%", delta=f"{gm_growth_rate:.2f}%")
                else:
                    st.metric("Gross Margin Growth Rate", "N/A", delta="N/A")

                # Gross Margin Percentage by Division
                st.subheader("Gross Margin Percentage by Division")
                gm_by_division = filtered_data.groupby('Group').agg(
                    {'Gross Margin': 'sum', 'Penjualan': 'sum'}).reset_index()
                gm_by_division['Gross Margin %'] = (gm_by_division['Gross Margin'] / gm_by_division['Penjualan']) * 100
                gm_by_division['Gross Margin %'] = gm_by_division['Gross Margin %'].fillna(0)  # Handle division by zero
                gm_by_division_sorted = gm_by_division.sort_values('Gross Margin %', ascending=False)

                fig_gm_division = px.bar(
                    gm_by_division_sorted,
                    x='Group',
                    y='Gross Margin %',
                    title="Gross Margin Percentage by Division",
                    labels={'Group': 'Division', 'Gross Margin %': 'Gross Margin Percentage (%)'},
                    color='Group',
                    color_discrete_sequence=color_palette
                )
                fig_gm_division.update_traces(hovertemplate="Division: %{x}<br>Gross Margin %: %{y:.2f}%")
                fig_gm_division.update_layout(
                    xaxis_title='Division',
                    yaxis_title='Gross Margin Percentage (%)',
                    legend_title='Division',
                    hovermode='x unified'
                )
                st.plotly_chart(fig_gm_division, use_container_width=True)

                # Gross Margin Percentage by Store
                st.subheader("Gross Margin Percentage by Store")
                gm_by_store = filtered_data.groupby('Store Name').agg(
                    {'Gross Margin': 'sum', 'Penjualan': 'sum'}).reset_index()
                gm_by_store['Gross Margin %'] = (gm_by_store['Gross Margin'] / gm_by_store['Penjualan']) * 100
                gm_by_store['Gross Margin %'] = gm_by_store['Gross Margin %'].fillna(0)  # Handle division by zero
                gm_by_store_sorted = gm_by_store.sort_values('Gross Margin %', ascending=False)

                fig_gm_store = px.bar(
                    gm_by_store_sorted,
                    x='Store Name',
                    y='Gross Margin %',
                    title="Gross Margin Percentage by Store",
                    labels={'Store Name': 'Store', 'Gross Margin %': 'Gross Margin Percentage (%)'},
                    color='Store Name',
                    color_discrete_sequence=color_palette
                )
                fig_gm_store.update_traces(hovertemplate="Store: %{x}<br>Gross Margin %: %{y:.2f}%")
                fig_gm_store.update_layout(
                    xaxis_title='Store',
                    yaxis_title='Gross Margin Percentage (%)',
                    legend_title='Store',
                    hovermode='x unified'
                )
                st.plotly_chart(fig_gm_store, use_container_width=True)

                # Detailed Gross Margin Data by Store and Grouping
                st.markdown("---")  # Separator for better UI
                show_detailed_store_table = st.checkbox(
                    "Show Detailed Gross Margin Data by Store and Grouping",
                    help="View detailed gross margin metrics categorized by each store and product group."
                )

                if show_detailed_store_table:
                    st.subheader("Detailed Gross Margin Data by Store and Grouping")
                    detailed_gm_store = filtered_data.groupby(['Store Name', 'Grouping']).agg(
                        {
                            'Gross Margin': 'sum',
                            'Penjualan': 'sum'
                        }
                    ).reset_index()

                    detailed_gm_store['Gross Margin %'] = (detailed_gm_store['Gross Margin'] / detailed_gm_store['Penjualan']) * 100
                    detailed_gm_store['Gross Margin %'] = detailed_gm_store['Gross Margin %'].fillna(0)

                    detailed_gm_store.rename(columns={
                        'Store Name': 'Store Name',
                        'Grouping': 'Grouping',
                        'Gross Margin': 'Gross Margin Value',
                        'Gross Margin %': 'Gross Margin Percentage (%)'
                    }, inplace=True)

                    # Use Styler for formatting
                    detailed_gm_store['Gross Margin Value'] = detailed_gm_store['Gross Margin Value']
                    detailed_gm_store['Gross Margin Percentage (%)'] = detailed_gm_store['Gross Margin Percentage (%)']

                    detailed_gm_store = detailed_gm_store.sort_values(by=['Gross Margin Value'], ascending=False)

                    detailed_gm_store_style = detailed_gm_store.style.format({
                        'Gross Margin Value': "{:,.0f}",
                        'Gross Margin Percentage (%)': "{:.2f}%"
                    }).highlight_max(axis=0)

                    st.dataframe(detailed_gm_store_style)

                show_detailed_division_table = st.checkbox(
                    "Show Detailed Gross Margin Data by Division, Store, Month, and Year",
                    help="View detailed gross margin metrics categorized by Division, Store, Month, and Year."
                )

                if show_detailed_division_table:
                    st.subheader("Detailed Gross Margin Data by Division, Store, Month, and Year")
                    # Ensure relevant columns are numeric
                    filtered_data['Gross Margin'] = pd.to_numeric(filtered_data['Gross Margin'], errors='coerce')
                    filtered_data['Penjualan'] = pd.to_numeric(filtered_data['Penjualan'], errors='coerce')

                    detailed_gm_division = filtered_data.groupby(['Group', 'Store Name', 'year', 'Month']).agg(
                        {
                            'Gross Margin': 'sum',
                            'Penjualan': 'sum'
                        }
                    ).reset_index()

                    detailed_gm_division['Gross Margin %'] = (detailed_gm_division['Gross Margin'] / detailed_gm_division['Penjualan']) * 100
                    detailed_gm_division['Gross Margin %'] = detailed_gm_division['Gross Margin %'].fillna(0)

                    detailed_gm_division.rename(columns={
                        'Group': 'Division',
                        'Store Name': 'Store Name',
                        'year': 'Year',
                        'Month': 'Month',
                        'Gross Margin': 'Gross Margin Value',
                        'Gross Margin %': 'Gross Margin Percentage (%)'
                    }, inplace=True)

                    # Use Styler for formatting
                    detailed_gm_division['Gross Margin Value'] = detailed_gm_division['Gross Margin Value']
                    detailed_gm_division['Gross Margin Percentage (%)'] = detailed_gm_division['Gross Margin Percentage (%)']

                    detailed_gm_division = detailed_gm_division.sort_values(
                        by=['Division', 'Store Name', 'Year', 'Month'],
                        ascending=[True, True, True, True]
                    )

                    detailed_gm_division_style = detailed_gm_division.style.format({
                        'Gross Margin Value': "{:,.0f}",
                        'Gross Margin Percentage (%)': "{:.2f}%"
                    }).highlight_max(axis=0)

                    st.dataframe(detailed_gm_division_style)

            # -------------------- 9. Stock Value Analysis (Tab 9) --------------------
            with tab9:
                st.header("Stock Value Analysis")
                st.markdown("""
                    Analyze the stock value data over time across different divisions and stores.
                    This helps in understanding inventory value trends, top/bottom stock value categories, and more.
                """)

                if filtered_data.empty:
                    st.write("No data available for Stock Value Analysis.")
                else:
                    # -------------------- Aggregate Stock Data --------------------
                    st.subheader("Total Stock Value by Group Over Months")
                    stock_data = filtered_data.groupby(['Group', 'Date'])['Stock Value'].sum().reset_index()
                    stock_data['Month_Display'] = stock_data['Date'].dt.strftime('%b %Y')

                    # Ensure consistent date parsing and chronological ordering
                    stock_data['Month_Display'] = pd.Categorical(
                        stock_data['Month_Display'],
                        categories=sorted(stock_data['Month_Display'].unique(),
                                          key=lambda x: datetime.strptime(x, '%b %Y')),
                        ordered=True
                    )

                    # -------------------- Line Chart of Stock Value Over Months by Group --------------------
                    if not stock_data.empty:
                        fig_stock = px.line(
                            stock_data,
                            x="Month_Display",
                            y="Stock Value",
                            color="Group",
                            title="Total Stock Value by Group Over Months",
                            labels={"Stock Value": "Total Stock Value", "Month_Display": "Month"},
                            color_discrete_sequence=px.colors.qualitative.Safe
                        )

                        fig_stock.update_traces(mode='lines+markers')
                        fig_stock.update_layout(
                            xaxis_title='Month',
                            yaxis_title='Stock Value',
                            legend_title='Group',
                            hovermode='x unified'
                        )

                        fig_stock.update_traces(
                            hovertemplate="Group: %{legendgroup}<br>Month: %{x}<br>Stock Value: %{y:,.0f}"
                        )

                        st.plotly_chart(fig_stock, use_container_width=True)

                    # -------------------- Top/Bottom Stock Value Categories (Grouping) --------------------
                    st.subheader("Top 10 Grouping by Average Stock Value")
                    stock_by_grouping_avg = filtered_data.groupby('Grouping')['Stock Value'].mean().reset_index()

                    top_stock_avg = stock_by_grouping_avg.nlargest(10, 'Stock Value')
                    top_stock_avg_style = top_stock_avg.rename(
                        columns={'Stock Value': 'Average Stock Value'}).style.format({
                        'Average Stock Value': "{:,.0f}"
                    })
                    st.dataframe(top_stock_avg_style)

                    st.subheader("Bottom 10 Grouping by Average Stock Value")
                    bottom_stock_avg = stock_by_grouping_avg[stock_by_grouping_avg['Stock Value'] > 0].nsmallest(10,
                                                                                                                 'Stock Value')

                    if bottom_stock_avg.empty:
                        st.write("No bottom performers with non-zero average stock value.")
                    else:
                        bottom_stock_avg_style = bottom_stock_avg.rename(
                            columns={'Stock Value': 'Average Stock Value'}).style.format({
                            'Average Stock Value': "{:,.0f}"
                        })
                        st.dataframe(bottom_stock_avg_style)

                    # -------------------- Detailed Stock Value by Store and Month --------------------
                    st.subheader("Detailed Stock Value by Store and Month")
                    store_stock_pivot = filtered_data.pivot_table(
                        values="Stock Value",
                        index="Store Name",
                        columns="Month_Display",
                        aggfunc="sum",
                        fill_value=0
                    )

                    store_stock_pivot = store_stock_pivot.reindex(
                        sorted(store_stock_pivot.columns, key=lambda x: datetime.strptime(x, '%b %Y')),
                        axis=1
                    )

                    store_stock_diff = store_stock_pivot.diff(axis=1).fillna(0)

                    combined_store_stock = pd.concat(
                        [store_stock_pivot, store_stock_diff],
                        keys=["Stock Value", "Difference"],
                        axis=1
                    )

                    combined_store_stock.columns.names = ['Type', 'Month']
                    combined_store_stock.reset_index(inplace=True)

                    combined_store_stock.columns = [
                        f"{col[0]}_{col[1]}" if col[0] != 'Store Name' else 'Store Name' for col in
                        combined_store_stock.columns
                    ]

                    combined_store_stock_style = combined_store_stock.style.format({
                        **{col: "{:,.0f}" for col in combined_store_stock.columns if
                           col.startswith('Stock Value_') or col.startswith('Difference_')},
                        'Store Name': '{}'
                    }).highlight_max(axis=0)
                    st.dataframe(combined_store_stock_style)

                    # -------------------- Compare Sales and Stock for Each Month --------------------
                    st.subheader("Comparison of Sales and Stock Value by Month and Group")
                    st.markdown("""
                        This table allows you to compare both Sales and Stock Value side-by-side for each month, helping you understand
                        how inventory levels relate to sales performance over time.
                    """)

                    comparison_basis = st.selectbox(
                        "Select Comparison Basis:",
                        options=["Division", "Store", "Grouping"],
                        help="Choose whether to compare Sales and Stock Value by Division, Store, or Grouping."
                    )

                    grouping_col = (
                        "Group" if comparison_basis == "Division"
                        else "Store Name" if comparison_basis == "Store"
                        else "Grouping"
                    )

                    sales_pivot_compare = filtered_data.pivot_table(
                        values="Penjualan",
                        index=grouping_col,
                        columns="Month_Display",
                        aggfunc="sum",
                        fill_value=0
                    )

                    stock_pivot_compare = filtered_data.pivot_table(
                        values="Stock Value",
                        index=grouping_col,
                        columns="Month_Display",
                        aggfunc="sum",
                        fill_value=0
                    )

                    all_months_compare = sorted(sales_pivot_compare.columns.union(stock_pivot_compare.columns),
                                                key=lambda x: datetime.strptime(x, '%b %Y'))

                    sales_pivot_compare = sales_pivot_compare.reindex(columns=all_months_compare, fill_value=0)
                    stock_pivot_compare = stock_pivot_compare.reindex(columns=all_months_compare, fill_value=0)

                    sales_pivot_compare.reset_index(inplace=True)
                    stock_pivot_compare.reset_index(inplace=True)

                    combined_sales_stock = pd.merge(
                        sales_pivot_compare,
                        stock_pivot_compare,
                        on=grouping_col,
                        how='outer',
                        suffixes=('_Sales', '_Stock')
                    ).fillna(0)

                    for month in all_months_compare:
                        sales_col = f"{month}_Sales"
                        stock_col = f"{month}_Stock"
                        pct_col = f"Stock%_{month}"
                        combined_sales_stock[pct_col] = combined_sales_stock.apply(
                            lambda row: (row[stock_col] / row[sales_col] * 100) if row[sales_col] != 0 else np.nan,
                            axis=1
                        )
                        combined_sales_stock[pct_col] = pd.to_numeric(combined_sales_stock[pct_col], errors='coerce')

                    combined_sales_stock_display = combined_sales_stock.copy()
                    for month in all_months_compare:
                        pct_col = f"Stock%_{month}"
                        combined_sales_stock_display[pct_col] = combined_sales_stock_display[pct_col].apply(
                            lambda x: f"{x:,.2f}%" if pd.notnull(x) else "N/A"
                        )

                    format_dict_sales_stock = {col: "{:,.0f}" for col in combined_sales_stock.columns if
                                               col.endswith('_Sales') or col.endswith('_Stock')}
                    format_dict_sales_stock.update(
                        {col: "{}" for col in combined_sales_stock.columns if col.startswith('Stock%_')})
                    format_dict_sales_stock[grouping_col] = '{}'

                    combined_sales_stock_style = combined_sales_stock_display.style.format(format_dict_sales_stock)

                    st.dataframe(combined_sales_stock_style)

                    # -------------------- Download Option for Comparison Table --------------------
                    csv = combined_sales_stock.to_csv(index=False)
                    st.download_button(
                        label="Download Comparison Table as CSV",
                        data=csv,
                        file_name='sales_stock_comparison.csv',
                        mime='text/csv',
                    )


    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")

else:
    st.info("Please upload an Excel file to proceed.")