import xlsxwriter
import os

def create_consolidation_template(output_path):
    workbook = xlsxwriter.Workbook(output_path)
    
    # --- Formats ---
    title_format = workbook.add_format({'bold': True, 'font_size': 16})
    note_format = workbook.add_format({'italic': True, 'font_size': 10, 'font_color': '#595959'})
    header_format = workbook.add_format({
        'bold': True, 'bg_color': '#2F5597', 'font_color': 'white', 
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    sub_header_format = workbook.add_format({
        'bold': True, 'bg_color': '#BDD7EE', 'border': 1
    })
    currency_format = workbook.add_format({'num_format': '$#,##0', 'border': 1})
    formula_format = workbook.add_format({
        'num_format': '$#,##0', 'bg_color': '#E2EFDA', 'border': 1, 'bold': True
    })
    border_format = workbook.add_format({'border': 1})
    elimination_format = workbook.add_format({
        'num_format': '$#,##0', 'bg_color': '#FCE4D6', 'border': 1, 'font_color': 'red'
    })
    rate_format = workbook.add_format({'num_format': '0.0000', 'bg_color': '#FFF2CC', 'border': 1})
    percent_format = workbook.add_format({'num_format': '0.0%', 'border': 1})

    # --- 1. Settings & Rates ---
    ws_rates = workbook.add_worksheet('FX Rates & Settings')
    ws_rates.set_column('A:D', 20)
    
    ws_rates.write('A1', 'Global Consolidation Settings', title_format)
    ws_rates.write('A3', 'Currency', header_format)
    ws_rates.write('B3', 'Spot Rate (Period End)', header_format)
    ws_rates.write('C3', 'Average Rate (P&L)', header_format)
    
    ws_rates.write('A4', 'USD (Base)', border_format)
    ws_rates.write('B4', 1.0000, rate_format)
    ws_rates.write('C4', 1.0000, rate_format)
    
    ws_rates.write('A5', 'EUR', border_format)
    ws_rates.write('B5', 1.0850, rate_format)
    ws_rates.write('C5', 1.0720, rate_format)
    
    ws_rates.write('A6', 'GBP', border_format)
    ws_rates.write('B6', 1.2640, rate_format)
    ws_rates.write('C6', 1.2510, rate_format)
    
    ws_rates.write('A8', 'Minority Interest %', border_format)
    ws_rates.write('B8', 0.15, percent_format)

    # --- 2. Consolidated P&L ---
    ws_pl = workbook.add_worksheet('Consolidated P&L')
    ws_pl.set_column('A:A', 35)
    ws_pl.set_column('B:G', 16)
    
    ws_pl.write('A1', 'Global Consolidated Income Statement', title_format)
    
    headers = ['P&L Line Item', 'Entity 1 (US - USD)', 'Entity 2 (EU - EUR)', 'Entity 3 (UK - GBP)', 'FX Adjustments', 'Intercomp Eliminations', 'Consolidated Total']
    ws_pl.write_row('A3', headers, header_format)
    
    def write_pl_row(row, label, us_val, eu_val, uk_val, is_subtotal=False, elim_val=0):
        fmt = sub_header_format if is_subtotal else border_format
        ws_pl.write(row, 0, label, fmt)
        ws_pl.write(row, 1, us_val, currency_format)
        
        # EU values with FX applied via formula referencing the rate sheet
        # Assuming Average Rate is on 'FX Rates & Settings'!$C$5 (for EUR) and $C$6 (for GBP)
        if not is_subtotal:
            ws_pl.write(row, 2, eu_val, currency_format) 
            ws_pl.write(row, 3, uk_val, currency_format)
            
            # FX Adj = (EU_Local * (EUR_Avg_Rate - 1)) + (UK_Local * (GBP_Avg_Rate - 1))
            row_l = row + 1
            ws_pl.write_formula(row, 4, f'=(C{row_l}*(\'FX Rates & Settings\'!$C$5-1)) + (D{row_l}*(\'FX Rates & Settings\'!$C$6-1))', formula_format)
            
            # Eliminations
            ws_pl.write(row, 5, elim_val, elimination_format)
            
            # Subtotal
            ws_pl.write_formula(row, 6, f'=SUM(B{row_l}:F{row_l})', formula_format)

    ws_pl.write('A4', 'Gross Revenue', border_format)
    ws_pl.write('B4', 5000000, currency_format)
    ws_pl.write('C4', 2000000, currency_format)
    ws_pl.write('D4', 1500000, currency_format)
    ws_pl.write_formula('E4', '=(C4*(\'FX Rates & Settings\'!$C$5-1)) + (D4*(\'FX Rates & Settings\'!$C$6-1))', formula_format)
    ws_pl.write('F4', -200000, elimination_format) # Intercompany sale elim
    ws_pl.write_formula('G4', '=SUM(B4:F4)', formula_format)

    ws_pl.write('A5', 'Cost of Goods Sold', border_format)
    ws_pl.write('B5', -2000000, currency_format)
    ws_pl.write('C5', -800000, currency_format)
    ws_pl.write('D5', -600000, currency_format)
    ws_pl.write_formula('E5', '=(C5*(\'FX Rates & Settings\'!$C$5-1)) + (D5*(\'FX Rates & Settings\'!$C$6-1))', formula_format)
    ws_pl.write('F5', 200000, elimination_format) # Matching COGS elim
    ws_pl.write_formula('G5', '=SUM(B5:F5)', formula_format)
    
    ws_pl.write('A6', 'Gross Profit', sub_header_format)
    for col in range(1, 7):
        col_name = xlsxwriter.utility.xl_col_to_name(col)
        ws_pl.write_formula(f'{col_name}6', f'=SUM({col_name}4:{col_name}5)', formula_format)
        
    ws_pl.write('A7', 'Operating Expenses', border_format)
    ws_pl.write('B7', -1000000, currency_format)
    ws_pl.write('C7', -500000, currency_format)
    ws_pl.write('D7', -400000, currency_format)
    ws_pl.write_formula('E7', '=(C7*(\'FX Rates & Settings\'!$C$5-1)) + (D7*(\'FX Rates & Settings\'!$C$6-1))', formula_format)
    ws_pl.write('F7', 0, elimination_format)
    ws_pl.write_formula('G7', '=SUM(B7:F7)', formula_format)

    ws_pl.write('A8', 'Net Income', sub_header_format)
    for col in range(1, 7):
        col_name = xlsxwriter.utility.xl_col_to_name(col)
        ws_pl.write_formula(f'{col_name}8', f'={col_name}6+{col_name}7', formula_format)
        
    ws_pl.write('A9', 'Minority Interest Adj', border_format)
    ws_pl.write('B9', 0, currency_format)
    ws_pl.write('C9', 0, currency_format)
    ws_pl.write('D9', 0, currency_format)
    ws_pl.write('E9', 0, currency_format)
    ws_pl.write('F9', 0, currency_format)
    # Minority Interest = Consolidated Net Income * 15% (assuming EU entity is 85% owned)
    ws_pl.write_formula('G9', f'=-C8*\'FX Rates & Settings\'!$C$5*\'FX Rates & Settings\'!$B$8', formula_format)
    
    ws_pl.write('A10', 'Net Income Attributable to Parent', header_format)
    ws_pl.write_formula('G10', '=G8+G9', formula_format)

    workbook.close()
    print(f"Consolidation & Roll-Up Template created at: {output_path}")

if __name__ == "__main__":
    out_dir = "/Users/mohithreddy/Desktop/Realtime Projects/Enterprise Planning System Deployment/01_Planning_Templates"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "1.4_Consolidation_Template.xlsx")
    create_consolidation_template(out_path)
