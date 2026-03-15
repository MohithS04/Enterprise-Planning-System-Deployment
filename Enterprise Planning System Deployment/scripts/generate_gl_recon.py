import xlsxwriter
import os

def create_gl_recon_template(output_path):
    workbook = xlsxwriter.Workbook(output_path)
    
    # --- Formats ---
    title_format = workbook.add_format({'bold': True, 'font_size': 16})
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#375623', 'font_color': 'white', 
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
    })
    
    currency_format = workbook.add_format({'num_format': '$#,##0', 'border': 1})
    formula_format = workbook.add_format({'num_format': '$#,##0', 'bg_color': '#E2EFDA', 'border': 1, 'bold': True})
    percent_format = workbook.add_format({'num_format': '0.0%', 'border': 1})
    
    # Conditional formatting formats
    success_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    warning_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'})
    error_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True})

    # --- GL Reconciliation Control Sheet ---
    ws = workbook.add_worksheet('GL Tie-Out Control')
    ws.set_column('A:B', 30)
    ws.set_column('C:E', 15)
    ws.set_column('F:G', 18)
    ws.set_column('H:H', 25)
    
    ws.write('A1', 'Automated GL to Planning System Tie-Out', title_format)
    ws.write('A2', 'Parameters: Tolerance = ±$500 OR ±0.1% Threshold')
    
    headers = [
        'GL Account Name', 'GL Account Code', 
        'Planning System Balance', 'GL Trial Balance', 
        'Variance ($)', 'Variance (%)', 'Status Alert', 'Resolution Note'
    ]
    ws.write_row('A4', headers, header_format)
    ws.set_row(3, 30) # wrap text height
    
    # Mock Data Rows
    data = [
        ['Software Subscriptions', '6010', 50000, 50000],          # Match
        ['Travel & Entertainment', '6020', 25500, 25000],          # Out by $500 (Fail)
        ['Office Supplies', '6030', 12000, 12050],                 # Out by -$50 (Warn/Fail depending on % - let's see)
        ['Professional Services', '6040', 155000, 155100],         # Out by -$100 (< 0.1% -> Pass via tolerance)
        ['Marketing & Advertising', '6050', 300000, 290000],       # Out by $10,000 (Fail)
    ]
    
    start_row = 4
    num_rows = len(data)
    for i, row_data in enumerate(data):
        row = start_row + i
        ws.write(row, 0, row_data[0], currency_format)
        ws.write(row, 1, row_data[1], currency_format)
        
        # Balances
        ws.write(row, 2, row_data[2], currency_format) # Planning balance
        ws.write(row, 3, row_data[3], currency_format) # GL TB balance
        
        # Variances
        ws.write_formula(row, 4, f'=C{row+1}-D{row+1}', currency_format)
        ws.write_formula(row, 5, f'=IFERROR(E{row+1}/D{row+1}, 0)', percent_format)
        
        # Status Alert Logic: IF(OR(ABS(Var$) > 500, ABS(Var%) > 0.001), "FAIL - Investigate", "PASS")
        ws.write_formula(row, 6, f'=IF(OR(ABS(E{row+1})>500, ABS(F{row+1})>0.001), "FAILED - Investigate", "PASS")', currency_format)
        
        # Note
        ws.write(row, 7, '', currency_format)
        
    # Apply Conditional Formatting to the Status string
    # Range is G5:G(last_row)
    status_range = f'G5:G{start_row+num_rows}'
    
    # PASS
    ws.conditional_format(status_range, {
        'type': 'cell',
        'criteria': '==',
        'value': '"PASS"',
        'format': success_format
    })
    
    # FAILED
    ws.conditional_format(status_range, {
        'type': 'text',
        'criteria': 'containing',
        'value': 'FAILED',
        'format': error_format
    })

    workbook.close()
    print(f"GL Recon Template created at: {output_path}")

if __name__ == "__main__":
    out_dir = "/Users/mohithreddy/Desktop/Realtime Projects/Enterprise Planning System Deployment/02_Validation_and_Reconciliation"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "2.1_GL_Reconciliation_Control_Sheet.xlsx")
    create_gl_recon_template(out_path)
