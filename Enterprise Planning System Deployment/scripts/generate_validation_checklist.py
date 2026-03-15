import xlsxwriter
import os

def create_validation_checklist(output_path):
    workbook = xlsxwriter.Workbook(output_path)
    
    # --- Formats ---
    title_format = workbook.add_format({'bold': True, 'font_size': 16})
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#548235', 'font_color': 'white', 
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    sub_header_format = workbook.add_format({
        'bold': True, 'bg_color': '#E2EFDA', 'border': 1
    })
    border_format = workbook.add_format({'border': 1})
    input_format = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1})
    date_format = workbook.add_format({'num_format': 'mm/dd/yyyy hh:mm', 'bg_color': '#FFF2CC', 'border': 1})
    
    # --- 1. Validation Rule Library ---
    ws_rules = workbook.add_worksheet('Validation Rule Library')
    ws_rules.set_column('A:A', 25)
    ws_rules.set_column('B:B', 60)
    ws_rules.set_column('C:C', 30)
    
    ws_rules.write('A1', 'System Validation & Integrity Rules', title_format)
    ws_rules.write_row('A3', ['Check Type', 'Rule Description', 'Action if Failed'], header_format)
    
    rules = [
        ['Completeness', 'All active GL accounts must be mapped in the planning tool.', 'Alert + Missing account log generated'],
        ['Mathematical Accuracy', 'Sum of all cost centers must equal the entity total.', 'Block submission workflow'],
        ['Period Integrity', 'No actuals may be posted or edited in future (open) periods.', 'Warning flag / Hard block'],
        ['Chart of Accounts Alignment', 'Planning account description must match GL TB description.', 'Mismatch variance report'],
        ['Sign Convention', 'Revenue = Credit (-), Expense = Debit (+).', 'Auto-flip sign with audit note']
    ]
    
    for i, rule in enumerate(rules):
        ws_rules.write_row(3 + i, 0, rule, border_format)

    # --- 2. Monthly Close Checklist ---
    ws_close = workbook.add_worksheet('Monthly Close Checklist')
    ws_close.set_column('A:A', 5)
    ws_close.set_column('B:B', 40)
    ws_close.set_column('C:C', 15)
    ws_close.set_column('D:D', 20)
    ws_close.set_column('E:E', 15)
    ws_close.set_column('F:F', 20)
    ws_close.set_column('G:G', 30)
    
    ws_close.write('B1', 'FP&A Monthly Close & Reconciliation Checklist', title_format)
    ws_close.write_row('A3', ['Step', 'Task Description', 'Status', 'Preparer Initial', 'Prep Date/Time', 'Reviewer Initial', 'Exception Notes'], header_format)
    
    tasks = [
        'Export GL Trial Balance from ERP (SAP/Oracle)',
        'Import Actuals into Planning System via API/Flat File',
        'Run Automated Tie-Out Control Sheet',
        'Clear all ±$500 variances',
        'Run Sign Convention & Completeness Validation Scripts',
        'Lock Actuals Period in Rolling Forecast Module',
        'Generate Monthly Budget vs. Actual Reporting Package',
        'Distribute Executive Flash Report'
    ]
    
    for i, task in enumerate(tasks):
        row = i + 3
        ws_close.write(row, 0, i+1, border_format)
        ws_close.write(row, 1, task, border_format)
        
        # Status Dropdown
        ws_close.data_validation(row, 2, row, 2, {
            'validate': 'list',
            'source': ['Not Started', 'In Progress', 'Complete', 'Exception'],
            'input_title': 'Select Status'
        })
        ws_close.write(row, 2, 'Not Started', input_format)
        
        ws_close.write(row, 3, '', input_format) # Prep
        ws_close.write(row, 4, '', date_format)  # Date
        ws_close.write(row, 5, '', input_format) # Rev
        ws_close.write(row, 6, '', input_format) # Notes

    # --- 3. Audit Trail Log ---
    ws_audit = workbook.add_worksheet('Audit Trail Log')
    ws_audit.set_column('A:A', 20)
    ws_audit.set_column('B:B', 20)
    ws_audit.set_column('C:C', 30)
    ws_audit.set_column('D:D', 40)
    
    ws_audit.write('A1', 'System Integration Audit Trail', title_format)
    ws_audit.write_row('A3', ['Timestamp', 'User ID', 'Action Type', 'Details / Override Reason'], header_format)
    
    # Mock some audit trail lines
    audit_data = [
        ['2026-03-01 08:30:12', 'system_api', 'Data Import', 'GL Trial Balance loaded for Feb-26'],
        ['2026-03-01 09:15:44', 'jsmith', 'Manual Override', 'Adjusted CC-200 mapping due to re-org'],
        ['2026-03-02 14:00:21', 'tjones', 'Workflow Approval', 'Marketing budget submitted and approved']
    ]
    
    for i, ad in enumerate(audit_data):
        ws_audit.write_row(3 + i, 0, ad, border_format)

    workbook.close()
    print(f"Validation Checklist Template created at: {output_path}")

if __name__ == "__main__":
    out_dir = "/Users/mohithreddy/Desktop/Realtime Projects/Enterprise Planning System Deployment/02_Validation_and_Reconciliation"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "2.2_Validation_and_Close_Checklist.xlsx")
    create_validation_checklist(out_path)
