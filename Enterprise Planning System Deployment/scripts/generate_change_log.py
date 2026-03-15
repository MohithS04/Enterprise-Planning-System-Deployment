import xlsxwriter
import os

def create_change_log(output_path):
    workbook = xlsxwriter.Workbook(output_path)
    
    # --- Formats ---
    title_format = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#C00000'})
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center'
    })
    curr_format = workbook.add_format({'num_format': '$#,##0', 'border': 1})
    text_format = workbook.add_format({'border': 1})
    pos_var_fmt = workbook.add_format({'num_format': '$#,##0', 'border': 1, 'font_color': 'green'})
    neg_var_fmt = workbook.add_format({'num_format': '$#,##0', 'border': 1, 'font_color': 'red'})
    
    # --- Change Log Sheet ---
    ws = workbook.add_worksheet('Budget Change Log')
    ws.set_column('A:B', 20)
    ws.set_column('C:C', 18)
    ws.set_column('D:D', 20)
    ws.set_column('E:E', 25)
    ws.set_column('F:F', 30)
    
    ws.write('A1', 'Monthly Forecast & Budget Change Log', title_format)
    ws.write('A2', 'Tracks all manual adjustments and re-forecast submissions.')
    
    # 1. Summary Table
    ws.write('A4', 'Net Budget Movement by Category', header_format)
    ws_sum = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1})
    
    summary_data = [
        ['Volume/Revenue Shift', -250000],
        ['Price Changes', 150000],
        ['Timing Shifts (OpEx)', 25000],
        ['New Initiatives', -100000],
        ['Corrections', 5000]
    ]
    
    ws.write_row('A5', ['Adjustment Reason', 'Net Impact ($)'], header_format)
    for i, (reason, amt) in enumerate(summary_data):
        ws.write(i+5, 0, reason, text_format)
        ws.write(i+5, 1, amt, curr_format)
    
    ws.write(10, 0, 'Total Net Change', ws_sum)
    ws.write_formula(10, 1, '=SUM(B6:B10)', curr_format)
    
    # 2. Detailed Log Register
    ws.write_row('A13', [
        'Date/Time', 'Line Item Changed', 'Adjustment ($)', 
        'Reason Code', 'Submitting Dept', 'Approver', 'Notes / Context'
    ], header_format)
    
    log_data = [
        ['2026-04-15', 'Gross Revenue - East', -250000, 'Volume', 'Sales', 'J. Doe', 'Lower store traffic'],
        ['2026-04-15', 'Gross Revenue - West', 150000, 'Price', 'Sales', 'J. Doe', 'Price hike realized'],
        ['2026-04-18', 'Software Subscriptions', 25000, 'Timing', 'IT', 'S. Smith', 'Pushed purchase to Q3'],
        ['2026-04-20', 'Marketing Campaigns', -100000, 'New Initiative', 'Marketing', 'M. Lee', 'Approved ad spend boost'],
        ['2026-04-25', 'Office Supplies', 5000, 'Correction', 'HR', 'System', 'Reversed duplicate entry']
    ]
    
    for i, row_data in enumerate(log_data):
        row = i + 13
        ws.write(row, 0, row_data[0], text_format)
        ws.write(row, 1, row_data[1], text_format)
        
        # Color formatting for adjustment amount
        fmt = pos_var_fmt if row_data[2] > 0 else neg_var_fmt
        ws.write(row, 2, row_data[2], fmt)
        
        ws.write(row, 3, row_data[3], text_format)
        
        # Add Dropdown Validation for Reason Code
        ws.data_validation(row, 3, row, 3, {
            'validate': 'list',
            'source': ['Volume', 'Price', 'Timing', 'New Initiative', 'Correction']
        })
        
        ws.write(row, 4, row_data[4], text_format)
        ws.write(row, 5, row_data[5], text_format)
        ws.write(row, 6, row_data[6], text_format)
        
    ws.autofilter('A13:G18')

    workbook.close()
    print(f"Change Log created at: {output_path}")

if __name__ == "__main__":
    out_dir = "/Users/mohithreddy/Desktop/Realtime Projects/Enterprise Planning System Deployment/04_Reporting_Suite"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "4.3_Monthly_Budget_Change_Log.xlsx")
    create_change_log(out_path)
