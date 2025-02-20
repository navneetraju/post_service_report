from io import BytesIO

import pandas as pd
import xlsxwriter


def create_report_pivot_table(df: pd.DataFrame, label: str):
    """ Generates the pivot table for Over Production Analysis """
    df.dropna(subset=['srvcrsname'], inplace=True)
    df['total_cost'] = df['served_prtncount'] * df['costprice']

    # Pivot Table for Over Production
    over_production_df = df[df['srvcrsname'].str.startswith("**")]
    pivot_over_production_df = over_production_df.pivot_table(index='srvcrsname', values='total_cost', aggfunc='sum')

    pivot_over_production_df.loc['Over Production'] = pivot_over_production_df.sum()
    pivot_over_production_df.index = pivot_over_production_df.index.str.replace("**", "")
    pivot_df = pivot_over_production_df.round(2)

    # Recipe Cost
    recipe_cost_df = df[~df['srvcrsname'].str.startswith("**") & (df['srvcrsname'] != 'Bars')]
    # total_recipe_cost = recipe_cost_df['total_cost'].sum()

    pivot_df['Percentage'] = round(pivot_df['total_cost'] / pivot_df.loc['Over Production']['total_cost'], 2)
    # pivot_df['Total Recipe Cost'] = round(total_recipe_cost, 2)

    pivot_df = pivot_df[["total_cost", "Percentage"]]
    pivot_df = pivot_df.rename(columns={"total_cost": "Cost"})
    pivot_df.index.name = label
    pivot_df.reset_index(inplace=True)

    pivot_df.rename(columns={"index": label}, inplace=True)
    pivot_df[label] = pivot_df[label].replace({"Thrown": "Waste"})
    desired_order = ["Reused", "Waste", "Donated", "Over Production"]
    pivot_df = pivot_df.set_index(label).reindex(desired_order).reset_index()
    pivot_df = pivot_df.set_index(label)
    pivot_df = pivot_df.reindex(["Reused", "Waste", "Donated", "Over Production"])
    pivot_df.reset_index(inplace=True)

    return pivot_df


def generate_exec_summary(evk_pivot, irc_pivot, uv_pivot):
    summary = pd.DataFrame()
    summary["Residential (All Units)"] = evk_pivot.iloc[:, 1] + irc_pivot.iloc[:, 1] + uv_pivot.iloc[:, 1]
    summary.index = ["Reused", "Waste", "Donated", "Over Production"]
    summary.reset_index(inplace=True)
    return summary


def generate_report(evk_df: pd.DataFrame, irc_df: pd.DataFrame, uv_df: pd.DataFrame):
    """ Writes all three tables (EVK, IRC, UV) into a single Excel sheet with proper formatting, including a chart. """

    # Create the pivot tables and executive summary
    evk_pivot = create_report_pivot_table(evk_df, "EVK")
    irc_pivot = create_report_pivot_table(irc_df, "IRC")
    uv_pivot = create_report_pivot_table(uv_df, "UV")
    executive_summary = generate_exec_summary(evk_pivot, irc_pivot, uv_pivot)

    # Create an in-memory buffer
    output = BytesIO()

    # Create a workbook and add a worksheet
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Over Production Summary")

    # Write Executive Summary
    worksheet.write(0, 0, "Executive Summary", workbook.add_format({'bold': True, 'font_size': 14}))
    headers = ["Residential (All Units)", "Total"]
    worksheet.write_row(1, 0, headers, workbook.add_format(
        {'bold': True, 'bg_color': '#2F75B5', 'font_color': 'white', 'align': 'center'}))

    for row_idx, row in enumerate(executive_summary.itertuples(index=False), start=2):
        worksheet.write(row_idx, 0, row[0])  # Label
        worksheet.write(row_idx, 1, row[1], workbook.add_format({'num_format': '"$"#,##0.00'}))  # Total

    # Create a new summary table for Over Production across all halls (EVK, IRC, UV)
    worksheet.write(7, 0, "Over Production Summary", workbook.add_format({'bold': True, 'font_size': 14}))
    worksheet.write_row(8, 0, ["Hall", "Total Cost"], workbook.add_format(
        {'bold': True, 'bg_color': '#2F75B5', 'font_color': 'white', 'align': 'center'}))

    # Write data for the new summary
    halls = ["EVK", "IRC", "UV"]
    over_production_totals = [
        evk_pivot.loc[3, "Cost"],
        irc_pivot.loc[3, "Cost"],
        uv_pivot.loc[3, "Cost"],
    ]

    for i, (hall, total) in enumerate(zip(halls, over_production_totals), start=9):
        worksheet.write(i, 0, hall)
        worksheet.write(i, 1, total, workbook.add_format({'num_format': '"$"#,##0.00'}))  # Total Cost

    # Add a bar chart for the new summary
    chart = workbook.add_chart({'type': 'pie'})

    # Consolidate the data for the chart
    chart.add_series({
        'name': 'Total Cost',
        'categories': ['Over Production Summary', 9, 0, 11, 0],  # Categories (Halls)
        'values': ['Over Production Summary', 9, 1, 11, 1],  # Total Cost values
        'data_labels': {'value': True}
    })

    # Customize chart axes and title
    chart.set_title({'name': 'Over Production Analysis (EVK, IRC, UV)'})
    chart.set_x_axis({'name': 'Residential Halls'})
    chart.set_y_axis({'name': 'Total Cost'})
    chart.set_size({'width': 540, 'height': 360})

    # Insert the chart into the worksheet
    worksheet.insert_chart('G6', chart)

    # Write tables for EVK, IRC, UV
    def write_table(dataframe, label, start_row):
        worksheet.write(start_row, 0, label, workbook.add_format({'bold': True, 'font_size': 14}))
        start_row += 1
        worksheet.write_row(start_row, 0, dataframe.columns, workbook.add_format({'bold': True, 'bg_color': '#E6E6E6'}))
        for row_idx, row in enumerate(dataframe.itertuples(index=False), start=start_row + 1):
            for col_idx, value in enumerate(row):
                if col_idx == 1:  # Format 'Total Recipe Cost' and 'Cost' columns
                    worksheet.write(row_idx, col_idx, value, workbook.add_format({'num_format': '"$"#,##0.00'}))
                elif col_idx == 2:  # Format 'Percentage' column
                    worksheet.write(row_idx, col_idx, value, workbook.add_format({'num_format': '0%'}))
                else:
                    worksheet.write(row_idx, col_idx, value)

    current_row = 14
    write_table(evk_pivot, "EVK", current_row)
    current_row += len(evk_pivot) + 2
    write_table(irc_pivot, "IRC", current_row)
    current_row += len(irc_pivot) + 2
    write_table(uv_pivot, "UV", current_row)

    # Auto-adjust column widths
    for col_num in range(worksheet.dim_colmax + 1):
        worksheet.set_column(col_num, col_num, 20)  # Adjust column widths to 20

    # Close the workbook
    workbook.close()

    # Return the in-memory Excel file
    output.seek(0)
    return output
    """ Writes all three tables (EVK, IRC, UV) into a single Excel sheet with proper formatting, including a chart. """

    # Create the pivot tables and executive summary
    evk_pivot = create_report_pivot_table(evk_df, "EVK")
    irc_pivot = create_report_pivot_table(irc_df, "IRC")
    uv_pivot = create_report_pivot_table(uv_df, "UV")
    executive_summary = generate_exec_summary(evk_pivot, irc_pivot, uv_pivot)

    # Set up the file path and workbook
    file_path = output_directory / "Over_Production_Summary.xlsx"
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet("Over Production Summary")

    # Write Executive Summary
    worksheet.write(0, 0, "Executive Summary", workbook.add_format({'bold': True, 'font_size': 14}))
    headers = ["Residential (All Units)", "Total", "Percentage"]
    worksheet.write_row(1, 0, headers, workbook.add_format(
        {'bold': True, 'bg_color': '#2F75B5', 'font_color': 'white', 'align': 'center'}))

    for row_idx, row in enumerate(executive_summary.itertuples(index=False), start=2):
        worksheet.write(row_idx, 0, row[0])  # Label
        worksheet.write(row_idx, 1, row[1], workbook.add_format({'num_format': '"$"#,##0.00'}))  # Total
        worksheet.write(row_idx, 2, row[2], workbook.add_format({'num_format': '0%'}))  # Percentage

    # Create a new summary table for Over Production across all halls (EVK, IRC, UV)
    worksheet.write(7, 0, "Over Production Summary", workbook.add_format({'bold': True, 'font_size': 14}))
    worksheet.write_row(8, 0, ["Hall", "Total Cost", "Percentage"], workbook.add_format(
        {'bold': True, 'bg_color': '#2F75B5', 'font_color': 'white', 'align': 'center'}))

    # Write data for the new summary
    halls = ["EVK", "IRC", "UV"]
    over_production_totals = [
        evk_pivot.loc[0, "Cost"],
        irc_pivot.loc[0, "Cost"],
        uv_pivot.loc[0, "Cost"],
    ]
    percentages = [
        evk_pivot.loc[0, "Percentage"],
        irc_pivot.loc[0, "Percentage"],
        uv_pivot.loc[0, "Percentage"],
    ]

    for i, (hall, total, percentage) in enumerate(zip(halls, over_production_totals, percentages), start=9):
        worksheet.write(i, 0, hall)
        worksheet.write(i, 1, total, workbook.add_format({'num_format': '"$"#,##0.00'}))  # Total Cost
        worksheet.write(i, 2, percentage, workbook.add_format({'num_format': '0%'}))  # Percentage

    # Add a bar chart for the new summary
    chart = workbook.add_chart({'type': 'column'})

    # Consolidate the data for the chart
    chart.add_series({
        'name': 'Total Cost',
        'categories': ['Over Production Summary', 9, 0, 11, 0],  # Categories (Halls)
        'values': ['Over Production Summary', 9, 1, 11, 1],  # Total Cost values
        'data_labels': {'value': True}
    })
    chart.add_series({
        'name': 'Percentage',
        'categories': ['Over Production Summary', 9, 0, 11, 0],  # Categories (Halls)
        'values': ['Over Production Summary', 9, 2, 11, 2],  # Percentage values
        'y2_axis': True,
        'data_labels': {'value': True}
    })

    # Customize chart axes and title
    chart.set_title({'name': 'Over Production Analysis (EVK, IRC, UV)'})
    chart.set_x_axis({'name': 'Residential Halls'})
    chart.set_y_axis({'name': 'Total Cost'})
    chart.set_y2_axis({'name': 'Percentage'})

    chart.set_size({'width': 720, 'height': 480})

    # Insert the chart into the worksheet
    worksheet.insert_chart('G6', chart)

    # Write tables for EVK, IRC, UV
    def write_table(dataframe, label, start_row):
        worksheet.write(start_row, 0, label, workbook.add_format({'bold': True, 'font_size': 14}))
        start_row += 1
        worksheet.write_row(start_row, 0, dataframe.columns, workbook.add_format({'bold': True, 'bg_color': '#E6E6E6'}))
        for row_idx, row in enumerate(dataframe.itertuples(index=False), start=start_row + 1):
            for col_idx, value in enumerate(row):
                if col_idx == 1:  # Format 'Total Recipe Cost' and 'Cost' columns
                    worksheet.write(row_idx, col_idx, value, workbook.add_format({'num_format': '"$"#,##0.00'}))
                elif col_idx == 3:  # Format 'Percentage' column
                    worksheet.write(row_idx, col_idx, value, workbook.add_format({'num_format': '0%'}))
                else:
                    worksheet.write(row_idx, col_idx, value)

    current_row = 14
    write_table(evk_pivot, "EVK", current_row)
    current_row += len(evk_pivot) + 2
    write_table(irc_pivot, "IRC", current_row)
    current_row += len(irc_pivot) + 2
    write_table(uv_pivot, "UV", current_row)

    # Auto-adjust column widths
    for col_num in range(worksheet.dim_colmax + 1):
        worksheet.set_column(col_num, col_num, 20)  # Adjust column widths to 20

    # Close the workbook
    workbook.close()

    print(
        f"✅ Excel file '{file_path}' created with Executive Summary, Over Production Summary, all tables, and a chart!")
