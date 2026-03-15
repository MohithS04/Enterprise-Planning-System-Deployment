import xlsxwriter
import os

def create_dept_submission_template(output_path):
    workbook = xlsxwriter.Workbook(output_path)
    
    # --- Formats ---
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#203764', 'font_color': 'white', 
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    sub_header_format = workbook.add_format({
        'bold': True, 'bg_color': '#D9E1F2', 'border': 1
    })
    
    input_text_format = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1})
    input_num_format = workbook.add_format({'num_format': '#,##0', 'bg_color': '#FFF2CC', 'border': 1})
    
    read_only_format = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1})
    formula_format = workbook.add_format({'num_format': '#,##0', 'bg_color': '#E2EFDA', 'border': 1, 'bold': True})
    
    title_format = workbook.add_format({'bold': True, 'font_size': 16})
    note_format = workbook.add_format({'italic': True, 'font_size': 10, 'font_color': '#595959'})
    border_format = workbook.add_format({'border': 1})
    
    # --- Status Tracker Sheet ---
    ws_status = workbook.add_worksheet('Submission Status')
    ws_status.set_column('A:B', 30)
    
    ws_status.write('A1', 'Submission Status Tracker', title_format)
    ws_status.write('A3', 'Metric', header_format)
    ws_status.write('B3', 'Value', header_format)
    
    ws_status.write('A4', 'Department Name', border_format)
    ws_status.write('B4', 'Marketing', input_text_format)
    
    ws_status.write('A5', 'Cost Center', border_format)
    ws_status.write('B5', 'CC-200', input_text_format)
    
    ws_status.write('A6', 'Submission Status', border_format)
    # Validation for status
    ws_status.data_validation('B6', {
        'validate': 'list',
        'source': ['Not Started', 'In Progress', 'Ready for Review', 'Submitted'],
        'input_title': 'Set Status',
        'input_message': 'Select current submission stage.'
    })
    ws_status.write('B6', 'In Progress', input_text_format)
    
    ws_status.write('A7', 'Total Submitted Budget', border_format)
    ws_status.write_formula('B7', '=SUM(\'Expense Input\'!H:H)', formula_format)

    # --- Lists / Dropdowns Sheet ---
    ws_lists = workbook.add_worksheet('Lookup Lists')
    ws_lists.hide() # Hide lookup sheet from business user
    
    gl_accounts = [
        '6010 - Software Subscriptions', 
        '6020 - Travel & Entertainment',
        '6030 - Office Supplies',
        '6040 - Professional Services',
        '6050 - Marketing & Advertising'
    ]
    cost_centers = [
        'CC-100 (Sales)',
        'CC-200 (Marketing)',
        'CC-300 (IT)',
        'CC-400 (HR)'
    ]
    
    ws_lists.write_column('A1', gl_accounts)
    ws_lists.write_column('B1', cost_centers)

    # --- Expense Input Form ---
    ws_input = workbook.add_worksheet('Expense Input')
    ws_input.set_column('A:A', 35)
    ws_input.set_column('B:B', 25)
    ws_input.set_column('C:H', 15)
    
    ws_input.write('A1', 'Simplified Department Expense Budget Form', title_format)
    ws_input.write('A2', 'Instructions: Fill in yellow cells. Use dropdowns where available. Must be positive numbers.', note_format)
    
    headers = ['GL Account', 'Cost Center', 'Description / Justification', 'Q1 Spend', 'Q2 Spend', 'Q3 Spend', 'Q4 Spend', 'FY Total']
    ws_input.write_row('A4', headers, header_format)
    
    # Pre-fill some empty input rows
    for row in range(4, 24): # 20 input rows
        # Dropdown for GL
        ws_input.data_validation(row, 0, row, 0, {
            'validate': 'list',
            'source': f"='Lookup Lists'!$A$1:$A${len(gl_accounts)}",
            'input_title': 'Choose GL Account'
        })
        ws_input.write(row, 0, '', input_text_format)
        
        # Dropdown for CC
        ws_input.data_validation(row, 1, row, 1, {
            'validate': 'list',
            'source': f"='Lookup Lists'!$B$1:$B${len(cost_centers)}",
            'input_title': 'Choose Cost Center'
        })
        ws_input.write(row, 1, '', input_text_format)
        
        # Description
        ws_input.write(row, 2, '', input_text_format)
        
        # Numeric validations (Q1-Q4) -> Must be >= 0
        for col in range(3, 7):
            ws_input.data_validation(row, col, row, col, {
                'validate': 'decimal',
                'criteria': '>=',
                'value': 0,
                'input_title': 'Enter Amount',
                'input_message': 'Must be >= 0',
                'error_title': 'Invalid Input',
                'error_message': 'Amount cannot be negative.'
            })
            ws_input.write(row, col, '', input_num_format)
            
        # Total Formula
        ws_input.write_formula(row, 7, f'=SUM(D{row+1}:G{row+1})', formula_format)

    # Auto-filter on the data table
    ws_input.autofilter('A4:H24')

    workbook.close()
    print(f"Departmental Submission Template created at: {output_path}")

if __name__ == "__main__":
    out_dir = "/Users/mohithreddy/Desktop/Realtime Projects/Enterprise Planning System Deployment/01_Planning_Templates"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "1.3_Departmental_Submission_Template.xlsx")
    create_dept_submission_template(out_path)
