import pandas as pd
import xlsxwriter


def add_report(data: pd.DataFrame, output: xlsxwriter.Workbook, worksheet):
    # Remove rows where 'srvcrsname' is in ['**Donated', '**Reused', '**Thrown']
    filtered_data = data[~data['srvcrsname'].isin(['**Donated', '**Reused', '**Thrown'])]

    # Convert 'eventdate' to datetime format and handle errors
    filtered_data['eventdate'] = pd.to_datetime(filtered_data['eventdate'], errors='coerce')

    # Drop rows where 'eventdate' could not be converted
    filtered_data = filtered_data.dropna(subset=['eventdate'])

    # Convert necessary columns to numeric and handle missing values
    filtered_data['fcst_prtncount'] = pd.to_numeric(filtered_data['fcst_prtncount'], errors='coerce').fillna(0)
    filtered_data['served_prtncount'] = pd.to_numeric(filtered_data['served_prtncount'], errors='coerce').fillna(0)
    filtered_data['costprice'] = pd.to_numeric(filtered_data['costprice'], errors='coerce').fillna(0)

    # Create new columns for cost calculations
    filtered_data['pre_service_cost'] = filtered_data['fcst_prtncount'] * filtered_data['costprice']
    filtered_data['post_service_cost'] = filtered_data['served_prtncount'] * filtered_data['costprice']

    # Group by date with unique customer counts
    unique_fcst_custcount_sum = (
        filtered_data.drop_duplicates(subset=['eventdate', 'fcst_custcount'])
        .groupby(filtered_data['eventdate'].dt.date)['fcst_custcount'].sum()
    )
    unique_sold_custcount_sum = (
        filtered_data.drop_duplicates(subset=['eventdate', 'sold_custcount'])
        .groupby(filtered_data['eventdate'].dt.date)['sold_custcount'].sum()
    )

    # Aggregate other metrics by date
    summary = filtered_data.groupby(filtered_data['eventdate'].dt.date).agg(
        pre_service_total_cost=('pre_service_cost', 'sum'),
        post_service_total_cost=('post_service_cost', 'sum')
    ).reset_index()

    # Merge with the unique customer counts
    summary = summary.merge(unique_fcst_custcount_sum, on='eventdate', how='left')
    summary = summary.merge(unique_sold_custcount_sum, on='eventdate', how='left')

    # Calculate revenue from rows with item names starting with 'RES _REVENUE_'
    revenue_data = filtered_data[filtered_data['itemname'].str.startswith('RES _REVENUE_')]
    revenue_summary = revenue_data.groupby(revenue_data['eventdate'].dt.date).agg(
        revenue=('sold_prtncount', 'sum')
    ).reset_index()

    # Merge revenue data into the summary
    summary = pd.merge(summary, revenue_summary, on='eventdate', how='left').fillna(0)

    # Rename columns for clarity
    summary.rename(columns={
        'eventdate': 'Date',
        'fcst_custcount': 'pre_service_cust_count',
        'sold_custcount': 'post_service_customer_count'
    }, inplace=True)

    summary = summary[
        ['Date', 'pre_service_cust_count', 'post_service_customer_count',
         'pre_service_total_cost', 'post_service_total_cost', 'revenue']
    ]

    summary.loc[len(summary)] = ['Totals:', summary['pre_service_cust_count'].sum(),
                                 summary['post_service_customer_count'].sum(),
                                 summary['pre_service_total_cost'].sum(),
                                 summary['post_service_total_cost'].sum(), summary['revenue'].sum()]
    summary.loc[len(summary) - 1, 'Date'] = 'Totals:'
    summary.loc[len(summary) - 1, 'pre_service_cust_count'] = summary.loc[len(summary) - 1, 'pre_service_cust_count']
    summary.loc[len(summary) - 1, 'post_service_customer_count'] = summary.loc[
        len(summary) - 1, 'post_service_customer_count']
    summary.loc[len(summary) - 1, 'pre_service_total_cost'] = summary.loc[len(summary) - 1, 'pre_service_total_cost']
    summary.loc[len(summary) - 1, 'post_service_total_cost'] = summary.loc[len(summary) - 1, 'post_service_total_cost']
    summary.loc[len(summary) - 1, 'revenue'] = summary.loc[len(summary) - 1, 'revenue']
    summary['total_cost_variance'] = (summary['post_service_total_cost'] - summary['pre_service_total_cost']) / summary[
        'pre_service_total_cost']
    summary['Day'] = summary['Date'].map(lambda x: pd.Timestamp(x).day_name() if x != 'Totals:' else 'Totals:')

    revenue_table = summary[['revenue']][:-1].copy()
    revenue_table.loc[len(revenue_table)] = revenue_table['revenue'].sum()
    revenue_table['sales_per_person'] = summary['revenue'] / summary['post_service_customer_count']
    revenue_table['cost_per_person'] = summary['post_service_total_cost'] / summary['post_service_customer_count']
    revenue_table['margin'] = revenue_table['sales_per_person'] - revenue_table['cost_per_person']
    revenue_table['margin_percentage'] = revenue_table['cost_per_person'] / revenue_table['sales_per_person']

    summary_table_to_write = summary[['Day', 'Date', 'pre_service_cust_count', 'post_service_customer_count',
                                      'pre_service_total_cost', 'post_service_total_cost', 'total_cost_variance']]

    title_format = output.add_format(
        {'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'italic': True, 'font_name': 'Arial',
         'border': 1, 'bottom': 1})
    worksheet.merge_range('A1:G1', 'Pre-Post Service Cost Summary', title_format)
    worksheet.set_row(0, 30)  # Set the height of the first row

    subtitle_format = output.add_format(
        {'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'font_name': 'Calibri', 'border': 1})
    worksheet.merge_range('A2:F2', ' (3/31/2025 to 4/6/2025) - Week 1', subtitle_format)
    worksheet.set_row(1, 20)  # Set the height of the second row

    header_format = output.add_format(
        {'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Arial',
         'bg_color': '#D9D9D9', 'italic': True, 'text_wrap': True})
    data_format = output.add_format(
        {'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Arial',
         'num_format': '#,##0'})
    bold_data_format = output.add_format(
        {'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Arial',
         'num_format': '#,##0',
         'bold': True})
    amount_format = output.add_format(
        {'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Arial',
         'num_format': '$#,##0.00'})
    bold_amount_format = output.add_format(
        {'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Arial', 'bold': True,
         'num_format': '$#,##0.00'})
    percent_format = output.add_format(
        {'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Arial',
         'num_format': '0%'})
    headers = ['Day', 'Date', 'Pre-Service Customer Count', 'Post-Service Customer Count',
               'Pre-Service Total Cost', 'Post-Service (Prepped) Total', 'Total Cost Variance']
    worksheet.write_row('A3', headers, header_format)
    for row_num, row_data in enumerate(summary_table_to_write.values, start=3):
        if row_data[0] == 'Totals:':
            worksheet.merge_range(row_num, 0, row_num, 1, row_data[0], output.add_format(
                {'font_size': 8, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_name': 'Arial',
                 'bold': True}))
            worksheet.write(row_num, 2, row_data[2], bold_data_format)
            worksheet.write(row_num, 3, row_data[3], bold_data_format)
            worksheet.write(row_num, 4, round(row_data[4], 2), bold_amount_format)
            worksheet.write(row_num, 5, round(row_data[5], 2), bold_amount_format)
            bg_color = '#FFC7CE' if row_data[6] <= -0.1 else '#C6EFCE'
            worksheet.write(row_num, 6, row_data[6], output.add_format(
                {'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Arial',
                 'num_format': '0%', 'bg_color': bg_color}))
        else:
            worksheet.write(row_num, 0, row_data[0], data_format)
            worksheet.write(row_num, 1, row_data[1].strftime('%m/%d/%Y'), data_format)
            worksheet.write(row_num, 2, row_data[2], data_format)
            worksheet.write(row_num, 3, row_data[3], data_format)
            worksheet.write(row_num, 4, round(row_data[4], 2), amount_format)
            worksheet.write(row_num, 5, round(row_data[5], 2), amount_format)
            bg_color = '#FFC7CE' if row_data[6] <= -0.1 else '#C6EFCE'
            worksheet.write(row_num, 6, row_data[6], output.add_format(
                {'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Arial',
                 'num_format': '0%', 'bg_color': bg_color}))

    worksheet.set_column_pixels('A:A', 94)
    worksheet.set_column_pixels('B:B', 82)
    worksheet.set_column_pixels('C:C', 121)
    worksheet.set_column_pixels('D:D', 130)
    worksheet.set_column_pixels('E:E', 122)
    worksheet.set_column_pixels('F:F', 129)
    worksheet.set_column_pixels('G:G', 64)

    worksheet.set_row_pixels(0, 31)

    worksheet.merge_range('A12:E12', 'Report Period  3/31/2025 - 4/6/2025 - Week 1', subtitle_format)

    worksheet.write_row('A13', ['Revenue', 'Sales Per Person', 'Cost Per Person', 'Margin ($)', 'Margin (%)'],
                        header_format)

    for idx, row in enumerate(revenue_table.values, start=13):
        if idx == len(revenue_table) + 12:

            worksheet.write(idx, 0, row[0], bold_amount_format)
            worksheet.write(idx, 1, row[1], bold_amount_format)
            worksheet.write(idx, 2, row[2], bold_amount_format)
            worksheet.write(idx, 3, row[3], bold_amount_format)
            worksheet.write(idx, 4, row[4], percent_format)
        else:
            worksheet.write(idx, 0, row[0], amount_format)
            worksheet.write(idx, 1, row[1], amount_format)
            worksheet.write(idx, 2, row[2], amount_format)
            worksheet.write(idx, 3, row[3], amount_format)
            worksheet.write(idx, 4, row[4], percent_format)
