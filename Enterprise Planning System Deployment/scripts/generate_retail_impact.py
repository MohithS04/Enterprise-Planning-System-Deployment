import xlsxwriter
import os

def create_retail_impact_report(output_path):
    workbook = xlsxwriter.Workbook(output_path)
    
    # --- Formats ---
    title_format = workbook.add_format({'bold': True, 'font_size': 18, 'font_color': '#44546A'})
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 
        'border': 1, 'align': 'center'
    })
    curr_format = workbook.add_format({'num_format': '$#,##0', 'border': 1})
    bold_curr_format = workbook.add_format({'num_format': '$#,##0', 'border': 1, 'bold': True, 'bg_color': '#D9E1F2'})
    box_format = workbook.add_format({
        'border': 2, 'bg_color': '#FCE4D6', 'text_wrap': True, 'valign': 'top'
    })
    box_title = workbook.add_format({'bold': True, 'bg_color': '#ED7D31', 'font_color': 'white', 'border': 1})

    ws = workbook.add_worksheet('Retail Impact Bridge')
    ws.hide_gridlines(2)
    ws.set_column('B:C', 20)
    ws.set_column('E:G', 25)
    
    ws.write('B2', 'Retail Operational Impact - Variance Bridge', title_format)
    
    # Data for the bridge chart
    ws.write_row('B4', ['Bridge Step', 'Impact ($)'], header_format)
    
    bridge_data = [
        ['Budgeted EBITDA', 5000000],
        ['Volume/Traffic Impact', -250000],
        ['Pricing & Mix Shift', 150000],
        ['Store-Level Cost Increases', -100000],
        ['One-Time Legal Settlement', -50000],
        ['Actual EBITDA', 4750000]
    ]
    
    for i, (label, val) in enumerate(bridge_data):
        fmt = bold_curr_format if i == 0 or i == len(bridge_data)-1 else curr_format
        ws.write_row(i+4, 1, [label, val], fmt)

    # Insert a Column Chart to act as a pseudo-waterfall/bridge chart
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': 'Operational Impact',
        'categories': ['Retail Impact Bridge', 4, 1, 9, 1], # B5:B10
        'values':     ['Retail Impact Bridge', 4, 2, 9, 2], # C5:C10
        'data_labels': {'value': True, 'num_format': '$#,##0'},
    })
    chart.set_title({'name': 'EBITDA Bridge (Budget to Actual)'})
    chart.set_legend({'none': True})
    
    # Add chart to sheet
    ws.insert_chart('B12', chart, {'x_scale': 1.5, 'y_scale': 1.2})
    
    # Callout Boxes Section
    ws.write('E4', 'Top 3 Operational Risks & Opportunities', box_title)
    
    risks_opps = (
        "1. RISK: Foot traffic in Northeast region down 5% YoY due to adverse weather events.\n\n"
        "2. OPP: Average order value (AOV) increased by $2.50 due to successful cross-sell promotional campaign.\n\n"
        "3. RISK: Hourly wage inflation at retail store level currently projecting to overrun budget by $100k next Q."
    )
    ws.merge_range('E5:G10', risks_opps, box_format)
    
    workbook.close()
    print(f"Retail Impact Report created at: {output_path}")

if __name__ == "__main__":
    out_dir = "/Users/mohithreddy/Desktop/Realtime Projects/Enterprise Planning System Deployment/04_Reporting_Suite"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "4.2_Retail_Operational_Impact.xlsx")
    create_retail_impact_report(out_path)
