import xlsxwriter
import os

def create_budget_template(output_path):
    workbook = xlsxwriter.Workbook(output_path)
    
    # --- Formatting Definitions ---
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    sub_header_format = workbook.add_format({
        'bold': True, 'bg_color': '#D9E1F2', 'border': 1
    })
    currency_format = workbook.add_format({'num_format': '$#,##0', 'border': 1})
    currency_input_format = workbook.add_format({
        'num_format': '$#,##0', 'bg_color': '#FFF2CC', 'border': 1
    })
    percent_format = workbook.add_format({'num_format': '0.0%', 'border': 1})
    percent_input_format = workbook.add_format({
        'num_format': '0.0%', 'bg_color': '#FFF2CC', 'border': 1
    })
    number_input_format = workbook.add_format({
        'num_format': '#,##0', 'bg_color': '#FFF2CC', 'border': 1
    })
    text_input_format = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1})
    formula_format = workbook.add_format({
        'num_format': '$#,##0', 'bg_color': '#E2EFDA', 'border': 1, 'bold': True
    })
    variance_format = workbook.add_format({'num_format': '$#,##0', 'border': 1})
    variance_percent_format = workbook.add_format({'num_format': '0.0%', 'border': 1})
    title_format = workbook.add_format({'bold': True, 'font_size': 16})
    note_format = workbook.add_format({'italic': True, 'font_size': 10, 'font_color': '#595959'})
    border_format = workbook.add_format({'border': 1})

    # --- Setup Workflow Sheet ---
    ws_workflow = workbook.add_worksheet('Approval Workflow')
    ws_workflow.set_column('A:A', 25)
    ws_workflow.set_column('B:B', 30)
    
    ws_workflow.write('A1', 'Annual Budget Approval Workflow', title_format)
    ws_workflow.write('A3', 'Department', header_format)
    ws_workflow.write('B3', 'Status', header_format)
    
    ws_workflow.write('A4', 'Sales', border_format)
    ws_workflow.write('A5', 'Marketing', border_format)
    ws_workflow.write('A6', 'IT', border_format)
    
    # Dropdown validation for status
    for row in range(3, 6):
        ws_workflow.data_validation(row, 1, row, 1, {
            'validate': 'list',
            'source': ['Draft', 'Submitted', 'Approved'],
            'input_title': 'Select Status',
            'input_message': 'Choose the current budget status.'
        })
        ws_workflow.write(row, 1, 'Draft', text_input_format)

    # --- Setup Budget Input Sheet ---
    ws_budget = workbook.add_worksheet('FY_Budget_Input')
    ws_budget.set_column('A:A', 30)
    ws_budget.set_column('B:B', 15)
    ws_budget.set_column('C:O', 12)
    ws_budget.set_column('P:R', 15)
    
    ws_budget.write('A1', 'FY Annual Budget Input by Department / Cost Center', title_format)
    ws_budget.write('A2', 'Note: Yellow cells denote input cells. Green cells are auto-calculated.', note_format)
    
    # Headers
    headers = ['Category / Account', 'GL Code', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'FY Total', 'Prior Year Actual', 'Variance ($)', 'Variance (%)']
    for col_num, header in enumerate(headers):
        ws_budget.write(3, col_num, header, header_format)
    
    # 1. Revenue
    ws_budget.write('A5', '1. Revenue Targets', sub_header_format)
    ws_budget.write('A6', 'Product A Sales', border_format)
    ws_budget.write('B6', '4000', border_format)
    
    # Write sample inputs and formulas
    for col in range(2, 14): # Jan-Dec
        ws_budget.write(5, col, 50000, currency_input_format)
    ws_budget.write_formula('O6', '=SUM(C6:N6)', formula_format)
    ws_budget.write('P6', 550000, currency_format) # PY Actual
    ws_budget.write_formula('Q6', '=O6-P6', variance_format)
    ws_budget.write_formula('R6', '=IFERROR(Q6/P6, 0)', variance_percent_format)
    
    # 2. Headcount & Compensation
    ws_budget.write('A8', '2. Headcount & Compensation', sub_header_format)
    ws_budget.write('A9', 'Headcount (FTE)', border_format)
    ws_budget.write('B9', 'N/A', border_format)
    for col in range(2, 14):
        ws_budget.write(8, col, 10, number_input_format)
    ws_budget.write_formula('O9', '=AVERAGE(C9:N9)', formula_format) # Avg FTE
    ws_budget.write('P9', 9, border_format) # PY FTE
    ws_budget.write_formula('Q9', '=O9-P9', border_format)
    ws_budget.write_formula('R9', '=IFERROR(Q9/P9, 0)', variance_percent_format)

    ws_budget.write('A10', 'Base Salaries', border_format)
    ws_budget.write('B10', '5000', border_format)
    for col in range(2, 14):
        ws_budget.write(9, col, 80000, currency_input_format)
    ws_budget.write_formula('O10', '=SUM(C10:N10)', formula_format)
    ws_budget.write('P10', 900000, currency_format)
    ws_budget.write_formula('Q10', '=O10-P10', variance_format)
    ws_budget.write_formula('R10', '=IFERROR(Q10/P10, 0)', variance_percent_format)

    ws_budget.write('A11', 'Benefits Load %', border_format)
    ws_budget.write('B11', 'N/A', border_format)
    for col in range(2, 14):
        ws_budget.write(10, col, 0.25, percent_input_format)
    ws_budget.write_formula('O11', '=AVERAGE(C11:N11)', percent_format)
    ws_budget.write('P11', 0.24, percent_format)
    ws_budget.write_formula('Q11', '=O11-P11', percent_format)
    ws_budget.write_formula('R11', '=IFERROR(Q11/P11, 0)', variance_percent_format)

    ws_budget.write('A12', 'Total Benefits Cost', border_format)
    ws_budget.write('B12', '5100', border_format)
    for col in range(2, 14):
        col_letter = xlsxwriter.utility.xl_col_to_name(col)
        ws_budget.write_formula(f'{col_letter}13', f'={col_letter}11*{col_letter}12', formula_format)
    ws_budget.write_formula('O12', '=SUM(C12:N12)', formula_format)
    ws_budget.write('P12', 216000, currency_format)
    ws_budget.write_formula('Q12', '=O12-P12', variance_format)
    ws_budget.write_formula('R12', '=IFERROR(Q12/P12, 0)', variance_percent_format)

    # 3. Direct & Indirect Setup
    ws_budget.write('A14', '3. Operating Expenses', sub_header_format)
    ws_budget.write('A15', 'Software Subscriptions', border_format)
    ws_budget.write('B15', '6010', border_format)
    for col in range(2, 14):
        ws_budget.write(14, col, 5000, currency_input_format)
    ws_budget.write_formula('O15', '=SUM(C15:N15)', formula_format)
    ws_budget.write('P15', 55000, currency_format)
    ws_budget.write_formula('Q15', '=O15-P15', variance_format)
    ws_budget.write_formula('R15', '=IFERROR(Q15/P15, 0)', variance_percent_format)

    ws_budget.write('A16', 'Travel & Ent', border_format)
    ws_budget.write('B16', '6020', border_format)
    for col in range(2, 14):
        ws_budget.write(15, col, 2000, currency_input_format)
    ws_budget.write_formula('O16', '=SUM(C16:N16)', formula_format)
    ws_budget.write('P16', 15000, currency_format)
    ws_budget.write_formula('Q16', '=O16-P16', variance_format)
    ws_budget.write_formula('R16', '=IFERROR(Q16/P16, 0)', variance_percent_format)

    # 4. CapEx
    ws_budget.write('A18', '4. Capital Expenditures', sub_header_format)
    ws_budget.write('A19', 'Server Upgrades', border_format)
    ws_budget.write('B19', '1500', border_format)
    for col in range(2, 14):
        val = 25000 if col == 2 else 0 # Only Jan has spend
        ws_budget.write(18, col, val, currency_input_format)
    ws_budget.write_formula('O19', '=SUM(C19:N19)', formula_format)
    ws_budget.write('P19', 0, currency_format)
    ws_budget.write_formula('Q19', '=O19-P19', variance_format)
    ws_budget.write_formula('R19', '=IFERROR(Q19/P19, 0)', variance_percent_format)

    # Justification text column for CapEx
    ws_budget.write('S4', 'CapEx Justification', header_format)
    ws_budget.write('S19', 'Q1 end-of-life server replacements', text_input_format)

    workbook.close()
    print(f"Annual Budget Template created at: {output_path}")

if __name__ == "__main__":
    out_dir = "/Users/mohithreddy/Desktop/Realtime Projects/Enterprise Planning System Deployment/01_Planning_Templates"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "1.1_Annual_Budget_Template.xlsx")
    create_budget_template(out_path)
