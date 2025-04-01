from io import BytesIO

import pandas as pd
import xlsxwriter


def create_report_pivot_table(df: pd.DataFrame, label: str):
    """ Generates the pivot table for Over Production Analysis """
    df.dropna(subset=['srvcrsname'], inplace=True)

    # Pivot Table for Over Production
    over_production_df = df[df['srvcrsname'].str.startswith("**")]
    pivot_over_production_df = over_production_df.pivot_table(index='srvcrsname', values='Total_Cost', aggfunc='sum')

    pivot_over_production_df.loc['Over Production'] = pivot_over_production_df.sum()
    pivot_over_production_df.index = pivot_over_production_df.index.str.replace("**", "")
    pivot_df = pivot_over_production_df.round(2)

    pivot_df['Percentage'] = round(pivot_df['Total_Cost'] / pivot_df.loc['Over Production']['Total_Cost'], 2)

    pivot_df = pivot_df[["Total_Cost", "Percentage"]]
    pivot_df = pivot_df.rename(columns={"Total_Cost": "Over Production"})
    pivot_df.index.name = label
    pivot_df.reset_index(inplace=True)

    pivot_df.rename(columns={"index": label}, inplace=True)
    pivot_df[label] = pivot_df[label].replace({"Thrown": "Waste"})
    desired_order = ["Reused", "Waste", "Donated", "Over Production"]
    pivot_df = pivot_df.set_index(label).reindex(desired_order).reset_index()

    return pivot_df


def generate_exec_summary(evk_pivot, irc_pivot, uv_pivot):
    summary = pd.DataFrame()
    summary["Over Production"] = evk_pivot.iloc[:, 1] + irc_pivot.iloc[:, 1] + uv_pivot.iloc[:, 1]
    summary.index = ["Reused", "Waste", "Donated", "Over Production"]
    summary["Percentage"] = summary["Over Production"] / summary.loc[
        "Over Production", "Over Production"]
    summary.reset_index(inplace=True)
    return summary


def generate_report(evk_df: pd.DataFrame, irc_df: pd.DataFrame, uv_df: pd.DataFrame):
    """ Writes all three tables (EVK, IRC, UV) into a single Excel sheet with proper formatting, including a chart. """

    # Convert `eventdate` to datetime format for all three dataframes
    for df in [evk_df, irc_df, uv_df]:
        df["eventdate"] = pd.to_datetime(df["eventdate"], errors="coerce")

    # Get the overall min and max date across all dataframes
    min_date = min(evk_df["eventdate"].min(), irc_df["eventdate"].min(), uv_df["eventdate"].min())
    max_date = max(evk_df["eventdate"].max(), irc_df["eventdate"].max(), uv_df["eventdate"].max())

    # Format date range for report header
    date_range_string = f"{min_date.strftime('%m/%d/%Y')} - {max_date.strftime('%m/%d/%Y')}"

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

    # **Write Report Header with Date Range**
    worksheet.write(0, 2, f"Over Production Monthly Summary {date_range_string}",
                    workbook.add_format({'bold': True, 'font_size': 16}))

    # **Write Executive Summary**
    worksheet.write(2, 0, "Executive Summary", workbook.add_format({'bold': True, 'font_size': 14}))
    headers = ["Over Production", "Total", "Percentage"]
    worksheet.write_row(2, 0, headers, workbook.add_format(
        {'bold': True, 'bg_color': '#2F75B5', 'font_color': 'white', 'align': 'center'}))

    for row_idx, row in enumerate(executive_summary.itertuples(index=False), start=3):
        worksheet.write(row_idx, 0, row[0], workbook.add_format({'bg_color': '#DCE6F1'})) if row[
                                                                                                 0] != "Over Production" else \
            worksheet.write(row_idx, 0, row[0])  # Hall
        worksheet.write(row_idx, 1, row[1],
                        workbook.add_format({'num_format': '"$"#,##0.00', 'bg_color': '#DCE6F1'})) if row[
                                                                                                          0] != "Over Production" else \
            worksheet.write(row_idx, 1, row[1],
                            workbook.add_format({'num_format': '"$"#,##0.00'}))  # Total
        worksheet.write(row_idx, 2, row[2],
                        workbook.add_format({'num_format': '0%', 'bg_color': '#DCE6F1'})) if row[
                                                                                                 0] != "Over Production" else \
            worksheet.write(row_idx, 2, row[2],
                            workbook.add_format({'num_format': '0%'}))  # Percentage

    # # **Write Over Production Summary**
    worksheet.write(9, 0, "Over Production Summary", workbook.add_format({'bold': True, 'font_size': 14}))
    worksheet.write_row(10, 0, ["Hall", "Total Cost"], workbook.add_format(
        {'bold': True, 'bg_color': '#2F75B5', 'font_color': 'white', 'align': 'center'}))

    # **Write Over Production Data**
    halls = ["EVK", "IRC", "UV"]
    over_production_totals = [
        evk_pivot.loc[3, "Over Production"],
        irc_pivot.loc[3, "Over Production"],
        uv_pivot.loc[3, "Over Production"],
    ]

    for i, (hall, total) in enumerate(zip(halls, over_production_totals), start=11):
        worksheet.write(i, 0, hall)
        worksheet.write(i, 1, total, workbook.add_format({'num_format': '"$"#,##0.00'}))  # Total Cost

    # **Add a Pie Chart for Over Production**
    chart = workbook.add_chart({'type': 'pie'})
    chart.add_series({
        'name': 'Total Cost',
        'categories': ['Over Production Summary', 3, 0, 5, 0],  # Categories (Halls)
        'values': ['Over Production Summary', 3, 1, 5, 1],  # Total Cost values
        'data_labels': {'value': True}
    })

    chart.set_title({'name': 'Residential Over Production'})
    chart.set_size({'width': 540, 'height': 360})

    # Insert the chart into the worksheet
    worksheet.insert_chart('G6', chart)

    # **Write Detailed Data Tables**
    def write_table(dataframe, start_row):
        start_row += 1
        worksheet.write_row(start_row, 0, dataframe.columns, workbook.add_format({'bold': True, 'bg_color': '#E6E6E6'}))
        for row_idx, row in enumerate(dataframe.itertuples(index=False), start=start_row + 1):
            is_overproduction_row = row[0] == "Over Production"
            for col_idx, value in enumerate(row):
                if col_idx == 1:  # Format 'Total Recipe Cost' and 'Cost' columns
                    worksheet.write(row_idx, col_idx, value, workbook.add_format(
                        {'num_format': '"$"#,##0.00', 'bg_color': '#DCE6F1'})) if not is_overproduction_row else \
                        worksheet.write(row_idx, col_idx, value, workbook.add_format({'num_format': '"$"#,##0.00'}))
                elif col_idx == 2:  # Format 'Percentage' column
                    worksheet.write(row_idx, col_idx, value, workbook.add_format(
                        {'num_format': '0%', 'bg_color': '#DCE6F1'})) if not is_overproduction_row else \
                        worksheet.write(row_idx, col_idx, value, workbook.add_format({'num_format': '0%'}))
                else:
                    worksheet.write(row_idx, col_idx, value,
                                    workbook.add_format({'bg_color': '#DCE6F1'})) if not is_overproduction_row else \
                        worksheet.write(row_idx, col_idx, value)

    current_row = 9
    write_table(evk_pivot, current_row)
    current_row += len(evk_pivot) + 2
    write_table(irc_pivot, current_row)
    current_row += len(irc_pivot) + 2
    write_table(uv_pivot, current_row)

    # Auto-adjust column widths
    for col_num in range(worksheet.dim_colmax + 1):
        worksheet.set_column(col_num, col_num, 20)  # Adjust column widths to 20

    workbook.close()

    output.seek(0)
    return output