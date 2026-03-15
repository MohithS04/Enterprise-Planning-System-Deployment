import xlsxwriter
import os

def create_rolling_forecast_template(output_path):
    workbook = xlsxwriter.Workbook(output_path)
    
    # --- Formatting Definitions ---
    # Global protection formats
    unlocked = workbook.add_format({'locked': False})
    
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#305496', 'font_color': 'white', 
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    sub_header_format = workbook.add_format({
        'bold': True, 'bg_color': '#D9E1F2', 'border': 1
    })
    
    # Default (locked) vs Input (unlocked)
    locked_num_format = workbook.add_format({'num_format': '#,##0', 'border': 1, 'bg_color': '#EFEFEF'})
    locked_curr_format = workbook.add_format({'num_format': '$#,##0', 'border': 1, 'bg_color': '#EFEFEF'})
    locked_pct_format = workbook.add_format({'num_format': '0.0%', 'border': 1, 'bg_color': '#EFEFEF'})
    
    input_num_format = workbook.add_format({'num_format': '#,##0', 'border': 1, 'bg_color': '#FFF2CC', 'locked': False})
    input_curr_format = workbook.add_format({'num_format': '$#,##0', 'border': 1, 'bg_color': '#FFF2CC', 'locked': False})
    input_pct_format = workbook.add_format({'num_format': '0.0%', 'border': 1, 'bg_color': '#FFF2CC', 'locked': False})
    
    formula_format = workbook.add_format({
        'num_format': '$#,##0', 'bg_color': '#E2EFDA', 'border': 1, 'bold': True
    })
    variance_format = workbook.add_format({'num_format': '$#,##0', 'border': 1})
    variance_pct_format = workbook.add_format({'num_format': '0.0%', 'border': 1})
    
    title_format = workbook.add_format({'bold': True, 'font_size': 16})
    note_format = workbook.add_format({'italic': True, 'font_size': 10, 'font_color': '#595959'})
    border_format = workbook.add_format({'border': 1})

    # --- Setup Forecasting Sheet ---
    ws = workbook.add_worksheet('Rolling 12-Month Forecast')
    ws.protect('password') # Protect the sheet
    
    ws.set_column('A:A', 35)
    ws.set_column('B:M', 12)
    ws.set_column('N:P', 15)
    
    ws.write('A1', '12-Month Rolling Forecast & Driver Inputs', title_format)
    ws.write('A2', 'Note: Yellow cells are open for forecast inputs. Gray cells are locked (Actuals).', note_format)
    
    # Months setup (Assume currently end of April, Actuals through April)
    months = ['Jan (Act)', 'Feb (Act)', 'Mar (Act)', 'Apr (Act)', 'May (Fcst)', 'Jun (Fcst)', 'Jul (Fcst)', 'Aug (Fcst)', 'Sep (Fcst)', 'Oct (Fcst)', 'Nov (Fcst)', 'Dec (Fcst)']
    for i, month in enumerate(months):
        ws.write(4, i+1, month, header_format)
    
    ws.write(4, 13, 'FY Forecast Total', header_format)
    ws.write(4, 14, 'Prior Forecast', header_format)
    ws.write(4, 15, 'Variance ($)', header_format)
    
    def write_row_data(row, label, actuals, forecast, is_driver, fmt_type):
        ws.write(row, 0, label, border_format)
        for i in range(4): # Jan-Apr Actuals
            fmt = locked_curr_format if fmt_type == 'curr' else (locked_pct_format if fmt_type == 'pct' else locked_num_format)
            ws.write(row, i+1, actuals[i], fmt)
        for i in range(4, 12): # May-Dec Forecast
            fmt = input_curr_format if fmt_type == 'curr' else (input_pct_format if fmt_type == 'pct' else input_num_format)
            ws.write(row, i+1, forecast[i-4], fmt)

    # 1. Driver Assumptions
    ws.write('A6', '1. Key Driver Assumptions', sub_header_format)
    
    write_row_data(6, 'Unit Volume Sold', [100, 110, 105, 115], [120, 130, 135, 140, 150, 160, 155, 180], True, 'num')
    write_row_data(7, 'Average Price per Unit', [50, 50, 50, 50], [55, 55, 55, 55, 60, 60, 60, 60], True, 'curr')
    
    write_row_data(8, 'Ending Headcount', [20, 20, 21, 22], [22, 23, 24, 25, 25, 26, 26, 28], True, 'num')
    write_row_data(9, 'Average Monthly Salary', [5000]*4, [5200]*8, True, 'curr')
    write_row_data(10, 'Benefits Load %', [0.25]*4, [0.26]*8, True, 'pct')
    
    # 2. Financial Forecast Calculation
    ws.write('A13', '2. Financial Forecast Summary', sub_header_format)
    
    # Revenue = Volume * Price
    ws.write('A14', 'Gross Revenue', border_format)
    for col in range(1, 13):
        col_name = xlsxwriter.utility.xl_col_to_name(col)
        ws.write_formula(f'{col_name}14', f'={col_name}7*{col_name}8', formula_format)
    ws.write_formula('N14', '=SUM(B14:M14)', formula_format)
    ws.write('O14', 850000, variance_format) # Prior Fcst
    ws.write_formula('P14', '=N14-O14', variance_format)
    
    # Payroll = Headcount * Avg Salary * (1 + Benefits Load)
    ws.write('A15', 'Total Payroll & Benefits', border_format)
    for col in range(1, 13):
        col_name = xlsxwriter.utility.xl_col_to_name(col)
        # Headcount(row9) * AvgSal(row10) * (1 + BenLoad(row11))
        ws.write_formula(f'{col_name}15', f'={col_name}9*{col_name}10*(1+{col_name}11)', formula_format)
    ws.write_formula('N15', '=SUM(B15:M15)', formula_format)
    ws.write('O15', 1500000, variance_format)
    ws.write_formula('P15', '=N15-O15', variance_format)
    
    # Variable Costs = Revenue * 15%
    ws.write('A16', 'Variable Operating Costs', border_format)
    for col in range(1, 13):
        col_name = xlsxwriter.utility.xl_col_to_name(col)
        ws.write_formula(f'{col_name}16', f'={col_name}14*0.15', formula_format)
    ws.write_formula('N16', '=SUM(B16:M16)', formula_format)
    ws.write('O16', 125000, variance_format)
    ws.write_formula('P16', '=N16-O16', variance_format)
    
    # Operating Income
    ws.write('A18', 'Net Operating Income', sub_header_format)
    for col in range(1, 13):
        col_name = xlsxwriter.utility.xl_col_to_name(col)
        ws.write_formula(f'{col_name}18', f'={col_name}14-{col_name}15-{col_name}16', formula_format)
    ws.write_formula('N18', '=SUM(B18:M18)', formula_format)
    ws.write('O18', -700000, variance_format)
    ws.write_formula('P18', '=N18-O18', variance_format)
    
    workbook.close()
    print(f"Rolling Forecast Template created at: {output_path}")

if __name__ == "__main__":
    out_dir = "/Users/mohithreddy/Desktop/Realtime Projects/Enterprise Planning System Deployment/01_Planning_Templates"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "1.2_Rolling_Forecast_Template.xlsx")
    create_rolling_forecast_template(out_path)
