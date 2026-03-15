import xlsxwriter
import os

def create_bva_report(output_path):
    workbook = xlsxwriter.Workbook(output_path)
    
    # --- Formats ---
    title_format = workbook.add_format({'bold': True, 'font_size': 18, 'font_color': '#1F4E78'})
    subtitle_format = workbook.add_format({'bold': True, 'font_size': 12, 'italic': True})
    
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    sub_header_format = workbook.add_format({
        'bold': True, 'bg_color': '#D9E1F2', 'border': 1
    })
    
    curr_format = workbook.add_format({'num_format': '$#,##0', 'border': 1})
    bold_curr_format = workbook.add_format({'num_format': '$#,##0', 'border': 1, 'bold': True, 'bg_color': '#F2F2F2'})
    pct_format = workbook.add_format({'num_format': '0.0%', 'border': 1})
    
    text_format = workbook.add_format({'border': 1})
    note_format = workbook.add_format({'italic': True, 'font_size': 10})
    
    # Conditional Formats
    green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    
    # --- Report Sheet ---
    ws = workbook.add_worksheet('Budget vs Actual')
    ws.hide_gridlines(2)
    
    ws.set_column('A:A', 35)
    ws.set_column('B:F', 16)
    ws.set_column('G:G', 5)
    ws.set_column('H:I', 35) # For variance heatmap by dept
    
    ws.write('A1', 'Monthly Budget vs. Actual Report', title_format)
    ws.write('A2', 'For the Period Ending: April 30, 2026', subtitle_format)
    
    # 1. Executive Summary Section
    ws.write('A4', 'Executive Summary', sub_header_format)
    ws.merge_range('A5:F5', '• Revenue missed budget by $100k (2%) driven by lower volume in the Northeast region.', text_format)
    ws.merge_range('A6:F6', '• Operating Expenses ran favorable to budget by $50k due to delayed headcount hiring.', text_format)
    ws.merge_range('A7:F7', '• Net EBITDA landing at $400k (11.1% margin), slightly below the 11.4% target.', text_format)

    # 2. P&L Table
    ws.write_row('A9', ['P&L Line Item', 'Budget', 'Actual', '$ Variance', '% Variance', 'Prior Year'], header_format)
    
    def write_pl_line(row, label, budget, actual, py, is_bold=False):
        fmt = bold_curr_format if is_bold else curr_format
        lbl_fmt = sub_header_format if is_bold else text_format
        ws.write(row, 0, label, lbl_fmt)
        ws.write(row, 1, budget, fmt)
        ws.write(row, 2, actual, fmt)
        
        # Variance logic: (Assume Revenue >. Actual-Budget. Expense > Budget - Actual)
        # Using a simple (Actual - Budget) for now, will fix signs via conditional formatting
        ws.write_formula(row, 3, f'=C{row+1}-B{row+1}', fmt)
        ws.write_formula(row, 4, f'=IFERROR(D{row+1}/B{row+1}, 0)', pct_format)
        ws.write(row, 5, py, fmt)

    write_pl_line(9, 'Gross Revenue', 5000000, 4900000, 4500000, True)
    write_pl_line(10, 'Cost of Goods Sold', -2500000, -2520000, -2300000, False)
    write_pl_line(11, 'Gross Margin', 2500000, 2380000, 2200000, True)
    
    write_pl_line(12, 'Payroll & Benefits', -1200000, -1150000, -1100000, False)
    write_pl_line(13, 'Marketing & Advertising', -400000, -420000, -350000, False)
    write_pl_line(14, 'Software & IT', -200000, -200000, -180000, False)
    write_pl_line(15, 'Travel & Entertainment', -100000, -90000, -80000, False)
    write_pl_line(16, 'Total Operating Expenses', -1900000, -1860000, -1710000, True)
    
    write_pl_line(17, 'EBITDA', 600000, 520000, 490000, True)
    
    # Conditional formatting on P&L $ Variance (Column D) (Positive = Green, Negative = Red)
    ws.conditional_format('D10:D18', {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': green_fmt})
    ws.conditional_format('D10:D18', {'type': 'cell', 'criteria': '<', 'value': 0, 'format': red_fmt})

    # 3. Variance Heatmap by Department
    ws.write('H9', 'OpEx Variance by Department', header_format)
    ws.write('I9', 'Variance (%)', header_format)
    
    depts = [('Sales', 0.05), ('Marketing', -0.04), ('IT', 0.00), ('HR', 0.02), ('Finance', -0.01)]
    for i, (dept, var) in enumerate(depts):
        ws.write(i+10, 7, dept, text_format)
        ws.write(i+10, 8, var, pct_format)
        
    # Heatmap conditional formatting on I11:I15 (using a 3-color scale)
    ws.conditional_format('I11:I15', {
        'type': '3_color_scale',
        'min_color': '#F8696B', # Red (unfavorable/negative)
        'mid_color': '#FFFFFF', # White (zero)
        'max_color': '#63BE7B'  # Green (favorable/positive)
    })

    # 4. Footnotes
    ws.write('A20', 'Footnotes & Variance Explanations:', subtitle_format)
    ws.write('A21', '1. COGS variance driven by supply chain expediting fees incurred in Week 2.', note_format)
    ws.write('A22', '2. Marketing overspend due to accelerated Q3 ad campaign rollout.', note_format)

    workbook.close()
    print(f"BvA Report created at: {output_path}")

if __name__ == "__main__":
    out_dir = "/Users/mohithreddy/Desktop/Realtime Projects/Enterprise Planning System Deployment/04_Reporting_Suite"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "4.1_Monthly_BvA_Report.xlsx")
    create_bva_report(out_path)
