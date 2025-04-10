from io import BytesIO

import pandas as pd
import xlsxwriter


def create_report_pivot_table(df: pd.DataFrame, label: str):
    df.dropna(subset=['srvcrsname'], inplace=True)
    over_production_df = df[df['srvcrsname'].str.startswith("**")]
    pivot_over_production_df = over_production_df.pivot_table(
        index='srvcrsname', values='Total_Cost', aggfunc='sum'
    )
    pivot_over_production_df.loc['Over Production'] = pivot_over_production_df.sum()
    pivot_over_production_df.index = pivot_over_production_df.index.str.replace("**", "")
    pivot_df = pivot_over_production_df.round(2)
    pivot_df['Percentage'] = round(
        pivot_df['Total_Cost'] / pivot_df.loc['Over Production']['Total_Cost'], 2
    )
    pivot_df = pivot_df[["Total_Cost", "Percentage"]]
    pivot_df = pivot_df.rename(columns={"Total_Cost": "Over Production"})
    pivot_df.index.name = label
    pivot_df.reset_index(inplace=True)
    pivot_df[label] = pivot_df[label].replace({"Thrown": "Waste"})
    desired_order = ["Reused", "Waste", "Donated", "Over Production"]
    pivot_df = pivot_df.set_index(label).reindex(desired_order).reset_index()
    return pivot_df


def generate_exec_summary(evk_pivot, irc_pivot, uv_pivot):
    summary = pd.DataFrame()
    summary["Over Production"] = (
            evk_pivot.iloc[:, 1] + irc_pivot.iloc[:, 1] + uv_pivot.iloc[:, 1]
    )
    summary.index = ["Reused", "Waste", "Donated", "Over Production"]
    summary["Percentage"] = summary["Over Production"] / summary.loc["Over Production", "Over Production"]
    summary.reset_index(inplace=True)
    return summary


def generate_report(evk_df: pd.DataFrame, irc_df: pd.DataFrame, uv_df: pd.DataFrame):
    # Convert eventdate to datetime
    for df in [evk_df, irc_df, uv_df]:
        df["eventdate"] = pd.to_datetime(df["eventdate"], errors="coerce")

    min_date = min(evk_df["eventdate"].min(), irc_df["eventdate"].min(), uv_df["eventdate"].min())
    max_date = max(evk_df["eventdate"].max(), irc_df["eventdate"].max(), uv_df["eventdate"].max())
    date_range_string = f"{min_date.strftime('%m/%d/%Y')} - {max_date.strftime('%m/%d/%Y')}"

    # Create pivot tables and the executive summary
    evk_pivot = create_report_pivot_table(evk_df, "EVK")
    irc_pivot = create_report_pivot_table(irc_df, "IRC")
    uv_pivot = create_report_pivot_table(uv_df, "UV")
    exec_summary = generate_exec_summary(evk_pivot, irc_pivot, uv_pivot)

    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Over Production Summary")

    # -----------------------------
    # Header and Subheader
    # -----------------------------
    header_format = workbook.add_format({'font_size': 14, 'font_color': 'white', 'bg_color': '#7B1FA2'})
    worksheet.merge_range(0, 0, 0, 7,
                          f"USC Hospitality - Over Production Monthly Summary ({date_range_string})",
                          header_format)

    subheader_format = workbook.add_format(
        {'font_size': 18, 'bg_color': '#DCE6F1', 'align': 'center', 'valign': 'vcenter'})
    worksheet.merge_range(1, 0, 1, 7, "Residential (All units)", subheader_format)

    # Adjust column widths:
    # Columns 0-2 for tables, 3 as spacer, columns 4-8 for charts.
    worksheet.set_column(0, 0, 15)  # Category
    worksheet.set_column(1, 1, 18)  # Cost
    worksheet.set_column(2, 2, 15)  # Percentage
    worksheet.set_column(3, 3, 5)  # Spacer
    worksheet.set_column(4, 8, 15)  # Chart area

    # -----------------------------
    # Executive Summary Block (with chart)
    # -----------------------------
    exec_start = 3
    exec_hdr_fmt = workbook.add_format({'align': 'center', 'bg_color': '#2F75B5', 'font_color': 'white'})
    worksheet.write_row(exec_start, 0, ["Category", "Total", "Percentage"], exec_hdr_fmt)
    row = exec_start + 1
    data_fmt = workbook.add_format({'bg_color': '#DCE6F1'})
    data_cur = workbook.add_format({'num_format': '"$"#,##0.00', 'bg_color': '#DCE6F1'})
    data_pct = workbook.add_format({'num_format': '0%', 'bg_color': '#DCE6F1'})
    for tup in exec_summary.itertuples(index=False):
        cat, tot, pct = tup
        if cat == "Over Production":
            worksheet.write(row, 0, cat)
            worksheet.write(row, 1, tot, workbook.add_format({'num_format': '"$"#,##0.00'}))
            worksheet.write(row, 2, pct, workbook.add_format({'num_format': '0%'}))
        else:
            worksheet.write(row, 0, cat, data_fmt)
            worksheet.write(row, 1, tot, data_cur)
            worksheet.write(row, 2, pct, data_pct)
        row += 1
    exec_end = row - 1

    # Insert Executive Summary Doughnut Chart (no slice labels)
    exec_chart = workbook.add_chart({'type': 'pie'})
    exec_chart.add_series({
        'name': "Executive Summary",
        'categories': ['Over Production Summary', exec_start + 1, 0, exec_end - 1, 0],
        'values': ['Over Production Summary', exec_start + 1, 2, exec_end - 1, 2],
        'data_labels': {'value': True, 'num_format': '0%'},  # Show percentage labels
    })
    exec_chart.set_title({'name': 'Executive Summary'})
    # exec_chart.set_hole_size(50)
    exec_chart.set_size({'width': 340, 'height': 220})
    worksheet.insert_chart(exec_start, 4, exec_chart, {'x_offset': 25, 'y_offset': 0})

    # Set the next starting row for the detailed blocks after some spacing
    current_row = exec_end + 10

    # -----------------------------
    # Function to write a detailed hall block (for EVK, IRC, or UV)
    # -----------------------------
    def write_block(title, pivot_df, start_row, color_main, color_header, color_data):
        # Title band for the block
        block_fmt = workbook.add_format(
            {'font_size': 18, 'bg_color': color_main, 'align': 'center', 'valign': 'vcenter'})
        worksheet.merge_range(start_row, 0, start_row, 7, title, block_fmt)
        start_row += 1

        # Table header
        hdr_fmt = workbook.add_format({'align': 'center', 'bg_color': color_header})
        worksheet.write_row(start_row, 0, ["Category", "Cost", "Percentage"], hdr_fmt)
        start_row += 1

        data_start = start_row
        d_fmt = workbook.add_format({'bg_color': color_data})
        d_cur = workbook.add_format({'num_format': '"$"#,##0.00', 'bg_color': color_data})
        d_pct = workbook.add_format({'num_format': '0%', 'bg_color': color_data})

        for row_data in pivot_df.itertuples(index=False):
            cat, cost, pct = row_data
            if cat == "Over Production":
                worksheet.write(start_row, 0, cat)
                worksheet.write(start_row, 1, cost, workbook.add_format({'num_format': '"$"#,##0.00'}))
                worksheet.write(start_row, 2, pct, workbook.add_format({'num_format': '0%'}))
            else:
                worksheet.write(start_row, 0, cat, d_fmt)
                worksheet.write(start_row, 1, cost, d_cur)
                worksheet.write(start_row, 2, pct, d_pct)
            start_row += 1
        data_end = start_row - 1

        # Create a smaller doughnut chart for this block (no slice labels)
        chart = workbook.add_chart({'type': 'pie'})
        chart.add_series({
            'categories': ['Over Production Summary', data_start, 0, data_end - 1, 0],
            'values': ['Over Production Summary', data_start, 2, data_end - 1, 2],
            'data_labels': {'value': True, 'num_format': '0%'},  # Show percentage labels
        })
        # chart.set_title({'name': title})
        # chart.set_hole_size(50)
        chart.set_size({'width': 340, 'height': 220})
        # Insert the chart to the right of the table (starting at column 5)
        worksheet.insert_chart(data_start, 4, chart, {'x_offset': 25, 'y_offset': 0})

        # Return the next starting row (with extra vertical spacing)
        return start_row + 10

    # -----------------------------
    # Detailed Breakdown Blocks (for each hall)
    # -----------------------------
    # EVK Block
    current_row = write_block(
        "EVK Breakdown",
        evk_pivot,
        current_row,
        color_main="#D9D9D9",
        color_header="#FFD965",
        color_data="#FFF2CC"
    )

    # IRC Block
    current_row = write_block(
        "IRC Breakdown",
        irc_pivot,
        current_row,
        color_main="#FDEADA",
        color_header="#F4B183",
        color_data="#FCE4D6"
    )

    # UV Block
    current_row = write_block(
        "UV Breakdown",
        uv_pivot,
        current_row,
        color_main="#E4F4EA",
        color_header="#9DC3E6",
        color_data="#DEEAF6"
    )

    workbook.close()
    output.seek(0)
    return output
