import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import numpy_financial as npf
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.datavalidation import DataValidation

# Page Title
st.title("NBA Team Underwriting Dashboard")

def generate_quarters(start_year, end_year):
    quarters = []
    for year in range(start_year, end_year + 1):
        for quarter in range(1, 5):
            quarters.append(f"{quarter}Q{year % 100:02d}")
    return quarters

quarter_options = generate_quarters(2025, 2040)

# Team Selection Dropdown
team = st.selectbox("Select Team", options=["Select Team", "Boston Celtics"], index=0)

# Set financial inputs based on the selected team
if team == "Boston Celtics":
    starting_revenue = 390
    starting_debt = 325
    starting_enterprise_value = 5660

    st.sidebar.header("Inputs")

    ownership_stake = st.sidebar.number_input("Desired Ownership Stake (%)", min_value=1.0, value=5.0, step=0.5)
    starting_enterprise_value = st.sidebar.number_input(
        "Starting Enterprise Value ($M)", min_value=0.0, value=float(starting_enterprise_value), step=10.0
    )
    starting_debt = st.sidebar.number_input(
        "Starting Debt ($M)", min_value=0.0, value=float(starting_debt), step=5.0
    )
    starting_equity = starting_enterprise_value - starting_debt
    starting_revenue = st.sidebar.number_input(
        "Starting Revenue ($M)", min_value=0.0, value=float(starting_revenue), step=10.0
    )

    st.sidebar.text(f"Starting TEV: ${starting_enterprise_value:.0f}M")

    entry_quarter = st.sidebar.selectbox("Entry Quarter", options=quarter_options, index=quarter_options.index("2Q25"))
    exit_quarter = st.sidebar.selectbox("Exit Quarter", options=quarter_options, index=quarter_options.index("2Q32"))
    desired_moic = st.sidebar.number_input("Desired MOIC (x)", min_value=1.0, value=2.5, step=0.1)

    def quarter_to_year(quarter):
        try:
            quarter_num = int(quarter[0])
            if quarter_num not in [1, 2, 3, 4]:
                raise ValueError("Quarter must be between 1 and 4.")
            year = int("20" + quarter[2:])
            return year + (quarter_num - 1) / 4
        except (ValueError, IndexError):
            st.error("Invalid quarter format. Use the format '1Q25'.")
        return None

    entry_year = quarter_to_year(entry_quarter)
    exit_year = quarter_to_year(exit_quarter)
    holding_period_years = exit_year - entry_year

    # Initial TEV/Revenue multiple
    entry_tev_revenue = starting_enterprise_value / starting_revenue

    def reset_to_average():
        return 11.9

    def reset_to_entry():
        return entry_tev_revenue

    def reset_to_comps():
        return 13.1

    new_tev_revenue = entry_tev_revenue
    
    # Adjust layout with minimal gap
    col_set_buttons = st.columns([1, 1, 1], gap="small")

    with col_set_buttons[0]:
        if st.button("Entry Multiple"):
            new_tev_revenue = reset_to_entry()

    with col_set_buttons[1]:
        if st.button("League Avg Multiple"):
            new_tev_revenue = reset_to_average()

    with col_set_buttons[2]:
        if st.button("Closest Comps Multiple"):
            new_tev_revenue = reset_to_comps()

        
    st.header("TEV/Revenue Adjustment")
    new_tev_revenue = st.slider("Adjust Exit TEV/Revenue Multiple", min_value=5.0, max_value=20.0, value=new_tev_revenue, step=0.1)

    # Graph 1
    fig_tev_revenue = go.Figure()
    fig_tev_revenue.add_trace(go.Bar(
        x=["Entry", "NBA Average", "Comps", "Exit"],
        y=[entry_tev_revenue, 11.9, 13.1, new_tev_revenue],
        marker_color=["#007A33", "gray", "navy", "#007A33"],
        text=[f"{entry_tev_revenue:.1f}", "11.9", "13.1", f"{new_tev_revenue:.1f}"],
        textposition="outside"
    ))
    fig_tev_revenue.update_layout(
        title="TEV/Revenue Comparison",
        yaxis_title="TEV/Revenue Multiple",
        template="plotly_white",
        barmode="group",
        height=500,
        width=600
    )

    def solve_for_revenue_growth():
        def objective(revenue_growth):
            exit_revenue = starting_revenue * ((1 + revenue_growth / 100) ** holding_period_years)
            exit_tev = exit_revenue * new_tev_revenue
            moic = exit_tev / starting_enterprise_value
            return abs(moic - desired_moic)

        from scipy.optimize import minimize
        result = minimize(objective, x0=[10], bounds=[(0, 30)])
        return result.x[0]

    required_revenue_growth = solve_for_revenue_growth()

    # Graph 2
    comparables = {"Warriors": 14, "Knicks": 12, "Lakers": 9}
    fig_growth = go.Figure()
    fig_growth.add_trace(go.Bar(
        x=list(comparables.keys()) + ["Required"],
        y=list(comparables.values()) + [required_revenue_growth],
        marker_color=["#FFC72C", "#006BB6", "#552583", "#007A33"],
        text=[f"{val:.1f}" for val in list(comparables.values()) + [required_revenue_growth]],
        textposition="outside"
    ))
    fig_growth.update_layout(
        title="Required Revenue Growth vs Comparables",
        yaxis_title="Revenue Growth (%)",
        template="plotly_white",
        barmode="group",
        height=500,
        width=600
    )

    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(fig_tev_revenue, use_container_width=True)
    with col2:
        st.plotly_chart(fig_growth, use_container_width=True)

    projected_revenue = [starting_revenue * ((1 + required_revenue_growth / 100) ** year) for year in range(int(holding_period_years) + 1)]
    entry_cash_flow = ownership_stake / 100 * starting_enterprise_value
    exit_cash_flow = ownership_stake / 100 * (projected_revenue[-1] * new_tev_revenue)

    cash_flows = [
        -entry_cash_flow
    ] + [0] * (int(holding_period_years) - 1) + [exit_cash_flow]

    def calculate_irr(cash_flows):
        return npf.irr(cash_flows) * 100

    irr = calculate_irr(cash_flows)
    moic = exit_cash_flow / abs(entry_cash_flow)

    def generate_styled_table(df):
        table_html = '<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">'
        table_html += '<tr style="background-color: #0056b3; color: white; font-weight: bold; text-align: center;">'
        for col in df.columns:
            table_html += f'<th style="padding: 10px; border: 1px solid #ddd;">{col}</th>'
        table_html += '</tr>'
        for _, row in df.iterrows():
            # Set both rows to white background
            table_html += '<tr style="background-color: white; text-align: center;">'
            for cell in row:
                table_html += f'<td style="padding: 10px; border: 1px solid #ddd;">{cell}</td>'
            table_html += '</tr>'
        table_html += '</table>'
        return table_html


    projections_styled = pd.DataFrame({
        "($M)": ["Revenue", "Cash Flow"],
        **{f"{int(entry_year) + i}": [f"{projected_revenue[i]:,.2f}", f"{cash_flows[i]:,.2f}"] for i in range(len(projected_revenue))}
    })

    # Generate the styled HTML table
    html_table = generate_styled_table(projections_styled)

    def generate_quarters_table_html_horizontal(df):
        html = '<table style="width: 100%; border-collapse: collapse;">'
        html += '<thead><tr style="background-color: #0056b3; color: white; text-align: center;">'
        html += ''.join(f'<th style="padding: 8px; border: 1px solid #ddd;">{col}</th>' for col in df.columns)
        html += '</tr></thead><tbody>'
        for _, row in df.iterrows():
            html += '<tr>'
            html += ''.join(f'<td style="padding: 8px; border: 1px solid #ddd; text-align: center;">{cell}</td>' for cell in row)
            html += '</tr>'
        html += '</tbody></table>'
        return html

    # Toggle for displaying quarterly data
    # Toggle for displaying quarterly data
    st.subheader("Projected Financials")
    # Add a dropdown to toggle between "Years" and "Quarters"
    view_option = st.selectbox("View By", options=["Years", "Quarters"], index=0)

    if view_option == "Quarters":
        # Generate quarters data for the expanded view
        quarter_labels = []
        revenue_row = []
        cash_flow_row = []

        for year_idx, annual_revenue in enumerate(projected_revenue):
            year = int(entry_year) + year_idx
            for q in range(1, 5):
                quarter_label = f"{q}Q{str(year)[-2:]}"
                quarter_labels.append(quarter_label)

                # Divide annual revenue by 4 for each quarter
                revenue_row.append(f"{annual_revenue / 4:.2f}")

                # Cash flow logic
                if quarter_label == entry_quarter:
                    # Entry Quarter: Investment (negative cash flow)
                    cash_flow_row.append(f"{cash_flows[0]:.2f}")
                elif quarter_label == exit_quarter:
                    # Exit Quarter: Return (positive cash flow)
                    cash_flow_row.append(f"{cash_flows[-1]:.2f}")
                else:
                    # Intermediate Quarters: No cash flow
                    cash_flow_row.append("0.00")

        # Create a DataFrame for the quarters table
        quarters_table = pd.DataFrame(
            [revenue_row, cash_flow_row],
            columns=quarter_labels,
            index=["Revenue ($M)", "Cash Flow ($M)"]
        ).reset_index()

        # Generate the styled HTML table
        html_quarters_table = generate_quarters_table_html_horizontal(quarters_table)
        st.markdown(html_quarters_table, unsafe_allow_html=True)

    else:
        # Show the original yearly table
        st.markdown(html_table, unsafe_allow_html=True)


    st.subheader("Investment Summary")
        
    investment_summary = pd.DataFrame({
        "Metric": [
            "Investment Amount ($M)",
            "Exit Amount ($M)",
            "IRR (%)",
            "MOIC (x)",
            "Entry Multiple",
            "Exit Multiple",
            "Required Avg. Revenue Growth"
        ],
        "Value": [
            f"${entry_cash_flow:,.0f}M",
            f"${exit_cash_flow:,.0f}M",
            f"{irr:.1f}%",
            f"{moic:.1f}x",
            f"{entry_tev_revenue:.1f}x",
            f"{new_tev_revenue:.1f}x",
            f"{required_revenue_growth:.1f}%"
        ]
    })
    
    def generate_summary_table_html(df):
        html = '<table style="width: 100%; border-collapse: collapse;">'
        html += '<thead><tr style="background-color: #0056b3; color: white; text-align: left;">'
        html += ''.join(f'<th style="padding: 8px; border: 1px solid #ddd;">{col}</th>' for col in df.columns)
        html += '</tr></thead><tbody>'
        for _, row in df.iterrows():
            html += '<tr>'
            html += ''.join(f'<td style="padding: 8px; border: 1px solid #ddd;">{cell}</td>' for cell in row)
            html += '</tr>'
        html += '</tbody></table>'
        return html


    # Display the table
    st.markdown(generate_summary_table_html(investment_summary), unsafe_allow_html=True)

        # Function to style Excel headers
    def style_headers(ws):
        header_fill = PatternFill(start_color="0056b3", end_color="0056b3", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        header_alignment = Alignment(horizontal="center", vertical="center")
        
        for cell in ws[1]:  # First row is the header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment

    # Function to adjust column widths
    def adjust_column_widths(ws):
        for col in ws.columns:
            max_length = max(len(str(cell.value) if cell.value else "") for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max_length + 2

    # Export to Excel with Two Sheets
    def export_to_excel(projections_df, summary_df):
        # Create a workbook
        wb = Workbook()

        # First Sheet: Projections
        ws1 = wb.active
        ws1.title = "Projections"
        
        # Add data to Projections sheet as numbers
        for row in dataframe_to_rows(projections_df, index=False, header=True):
            ws1.append(row)

        # Convert all cells (except headers) to numeric format if possible
        for row in ws1.iter_rows(min_row=2, min_col=2):
            for cell in row:
                try:
                    cell.value = float(cell.value)  # Convert to numeric
                except (ValueError, TypeError):
                    pass  # Leave as is if not convertible

        # Style headers and adjust column widths
        style_headers(ws1)
        adjust_column_widths(ws1)

        # Add dynamic revenue formulas to Projections sheet
        for col_idx in range(3, len(projections_df.columns) + 1):  # Start from the second year onward
            col_letter = ws1.cell(1, col_idx).column_letter
            prev_col_letter = ws1.cell(1, col_idx - 1).column_letter
            # Revenue formula: Previous year's revenue * (1 + growth rate from Summary sheet)
            ws1.cell(2, col_idx).value = f"={prev_col_letter}2*(1+(Summary!$B$6)/100)"

        # Add the input revenue directly for the first column
        ws1.cell(2, 2).value = float(projections_df.iloc[0, 1])  # Input revenue in B2

        # Add TEV formulas
        ws1.cell(4, 2).value = f"=B2*Summary!$B$4"  # Entry TEV = Revenue * Entry Multiple
        ws1.cell(4, ws1.max_column).value = f"={ws1.cell(2, ws1.max_column).coordinate}*Summary!$B$5"  # Exit TEV

        # Update label for TEV row
        ws1.cell(4, 1).value = "Implied Enterprise Value"

        # Second Sheet: Summary
        ws2 = wb.create_sheet(title="Summary")
        
        # Clean summary data: Remove `%`, `M`, and `x` suffixes
        summary_cleaned = summary_df.copy()
        summary_cleaned["Value"] = summary_cleaned["Value"].replace(
            {r"[^\d\.]": ""}, regex=True  # Remove non-numeric characters
        ).astype(float)  # Convert to float

        # Add data to Summary sheet
        for row in dataframe_to_rows(summary_cleaned, index=False, header=True):
            ws2.append(row)

        # Style headers and adjust column widths
        style_headers(ws2)
        adjust_column_widths(ws2)

        # Add IRR and MOIC formulas directly to the Summary sheet
        cash_flow_range = f"'Projections'!B3:{ws1.cell(3, len(projections_df.columns)).column_letter}3"  # Cash flow row
        ws2.cell(len(summary_cleaned) + 2, 1).value = "IRR (%)"
        ws2.cell(len(summary_cleaned) + 2, 2).value = f"=IRR({cash_flow_range})"
        ws2.cell(len(summary_cleaned) + 3, 1).value = "MOIC (x)"
        ws2.cell(len(summary_cleaned) + 3, 2).value = f"='Projections'!{ws1.cell(3, len(projections_df.columns)).coordinate}/ABS('Projections'!{ws1.cell(3, 2).coordinate})"

        ws2.delete_rows(4, 2)

        # Save to BytesIO for Streamlit download
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return output

    # Prepare DataFrames for Export
    if view_option == "Years":
        projections_df = projections_styled  # Use the annual projections
    else:
        projections_df = pd.DataFrame(
            [revenue_row, cash_flow_row],
            columns=quarter_labels,
            index=["Revenue ($M)", "Cash Flow ($M)"]
        ).reset_index()

    summary_df = investment_summary  # Use the summary DataFrame

    # Export Button in Streamlit
    st.download_button(
        label="Download Excel File",
        data=export_to_excel(projections_df, summary_df),
        file_name="CelticsModel_v1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
