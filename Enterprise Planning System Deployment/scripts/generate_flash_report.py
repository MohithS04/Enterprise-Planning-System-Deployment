import xlsxwriter
import os

def create_flash_report(output_path):
    workbook = xlsxwriter.Workbook(output_path)
    
    # --- Formats ---
    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'font_color': '#333f4f', 'align': 'center'})
    date_format = workbook.add_format({'italic': True, 'align': 'center', 'font_color': '#595959'})
    
    kpi_title = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center', 'font_size': 14})
    kpi_val = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'font_size': 18})
    
    # Traffic Lights
    green_light = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True, 'align': 'center', 'border': 1})
    yellow_light = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'bold': True, 'align': 'center', 'border': 1})
    red_light = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True, 'align': 'center', 'border': 1})
    
    section_header = workbook.add_format({'bold': True, 'bg_color': '#2F5597', 'font_color': 'white', 'font_size': 12, 'border': 1})
    box_format = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top', 'bg_color': '#F2F2F2'})

    ws = workbook.add_worksheet('Executive Flash')
    ws.hide_gridlines(2)
    # Set paper size and fit to page
    ws.set_paper(9) # A4
    ws.fit_to_pages(1, 1)
    ws.set_landscape()

    ws.set_column('A:A', 5)
    ws.set_column('B:E', 25)
    ws.set_column('F:F', 5)
    
    # Header
    ws.merge_range('B2:E2', 'EXECUTIVE FLASH REPORT', title_format)
    ws.merge_range('B3:E3', 'For the Month Ended: April 30, 2026', date_format)
    
    # 1. KPI Tiles
    ws.write('B5', 'Revenue (YTD)', kpi_title)
    ws.write('C5', 'EBITDA Margin', kpi_title)
    ws.write('D5', 'Ending Headcount', kpi_title)
    ws.write('E5', 'Cash Balance', kpi_title)
    
    curr = workbook.add_format({'num_format': '$#,##0', 'bold': True, 'border': 1, 'align': 'center', 'font_size': 16})
    pct = workbook.add_format({'num_format': '0.0%', 'bold': True, 'border': 1, 'align': 'center', 'font_size': 16})
    num = workbook.add_format({'num_format': '#,##0', 'bold': True, 'border': 1, 'align': 'center', 'font_size': 16})
    
    ws.write('B6', 15400000, curr)
    ws.write('C6', 0.111, pct)
    ws.write('D6', 224, num)
    ws.write('E6', 4200000, curr)
    
    # Traffic Lights
    ws.write('B7', 'GREEN (+2%)', green_light)
    ws.write('C7', 'YELLOW (-0.3%)', yellow_light)
    ws.write('D7', 'GREEN (On Plan)', green_light)
    ws.write('E7', 'RED (-$500k)', red_light)
    
    # 2. Chart Data setup (Hidden columns ideally, but we'll put right below for now)
    ws.write('B9', 'Month', kpi_title)
    ws.write('C9', 'Budget', kpi_title)
    ws.write('D9', 'Actual', kpi_title)
    
    months = ['Jan', 'Feb', 'Mar', 'Apr']
    budget = [3800000, 3900000, 4200000, 4000000]
    actual = [3850000, 3950000, 4100000, 3900000]
    
    for i in range(4):
        ws.write(10+i, 1, months[i])
        ws.write(10+i, 2, budget[i])
        ws.write(10+i, 3, actual[i])
        
    # Create Line Chart
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({
        'name': 'Actuals',
        'categories': ['Executive Flash', 10, 1, 13, 1],
        'values': ['Executive Flash', 10, 3, 13, 3],
        'line': {'color': '#4472C4', 'width': 2.25}
    })
    chart.add_series({
        'name': 'Budget',
        'categories': ['Executive Flash', 10, 1, 13, 1],
        'values': ['Executive Flash', 10, 2, 13, 2],
        'line': {'color': '#A5A5A5', 'dash_type': 'dash', 'width': 1.5}
    })
    chart.set_title({'name': 'YTD Revenue Trend ($)'})
    chart.set_legend({'position': 'bottom'})
    
    ws.insert_chart('B16', chart, {'x_scale': 1.6, 'y_scale': 1.2})
    
    # 3. Commentary Section
    ws.merge_range('D16:E16', 'Top 3 Wins', section_header)
    ws.merge_range('D17:E19', 
        "1. Successful rollout of new SaaS product line.\n"
        "2. Favorable variance in IT hardware costs.\n"
        "3. Employee retention up 5% over Q1.", box_format)
        
    ws.merge_range('D20:E20', 'Top 3 Risks', section_header)
    ws.merge_range('D21:E23',
        "1. Cash collection cycle extending by 4 days.\n"
        "2. Supply chain delays impacting May deliveries.\n"
        "3. Increased hourly wage competition.", box_format)
        
    ws.merge_range('D24:E24', 'Key Decisions Needed', section_header)
    ws.merge_range('D25:E27',
        "1. Approve $500k draw from revolver to pad cash balance?\n"
        "2. Delay next wave of hiring until Q3?", box_format)

    workbook.close()
    print(f"Flash Report created at: {output_path}")

if __name__ == "__main__":
    out_dir = "/Users/mohithreddy/Desktop/Realtime Projects/Enterprise Planning System Deployment/04_Reporting_Suite"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "4.4_Executive_Flash_Report.xlsx")
    create_flash_report(out_path)
