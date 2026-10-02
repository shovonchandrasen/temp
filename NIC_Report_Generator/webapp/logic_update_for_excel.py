import pandas as pd
import openpyxl
import os
import glob

def get_file(pattern):
    # Case-insensitive search: [iI] matches 'I' or 'i'
    # This searches for files starting with 'inventory activity' in the current directory
    files = glob.glob(pattern)
    return files[0] if files else None

def process_inventory():
    stock_report_file = get_file('*[sS]tore [wW]ise [sS]tock*.csv')  #'Store Wise Stock Report-from-2026-02-01-to-2026-02-28.csv'
    input_file = get_file('[iI]nventory [aA]ctivity*.xlsx') #'Inventory Activity Report Feb 26-input.xlsx'
    input_file_2 = get_file('*[cC]losing [sS]tock*.xlsx')  #'Closing Stock Report Feb 26-input.xlsx'
    
    if not os.path.exists(stock_report_file) or not os.path.exists(input_file) or not os.path.exists(input_file_2):
        print("Error: Files not found.")
        return

    # 1. Load reference data
    df_stock = pd.read_csv(stock_report_file)
    df_stock['Item Name'] = df_stock['Item Name'].astype(str).str.strip()
    # Create a dictionary for fast lookup (case-insensitive)
    stock_map = {
        str(k).strip().lower(): v
        for k, v in df_stock.set_index('Item Name').to_dict('index').items()
    }
    
    # 1.1 Load reference data 2
    #df_stock_closing = pd.read_excel(input_file_2, engine='openpyxl', header=6)
    #df_stock_closing['Product name'] = df_stock_closing['Product name'].astype(str).str.strip()
    # Create a dictionary for fast lookup
    #stock_closing_map = df_stock_closing.set_index('Product name').to_dict('index')
    
    wb2= openpyxl.load_workbook(input_file_2)
    ws2 = wb2.active
    header_row2=7
    headers2 = [c.value for c in ws2[header_row2]]
    idx = headers2.index('Product name')

    # Create the map in a single comprehension
    #stock_closing_map = {
        #str(r[idx]).strip(): {headers[i]: r[i] for i in range(len(headers)) if i != idx}
        
        #for r in ws.iter_rows(min_row=8, values_only=True)
    #}
    
    stock_closing_map = {
        str(r[idx]).strip().lower(): {
            "Physical": r[idx+2]
        }
        for r in ws2.iter_rows(min_row=8, values_only=True)
        if r[idx] is not None
    }
    
    #print(stock_closing_map)

    # 2. Load Excel with openpyxl to maintain formatting
    wb = openpyxl.load_workbook(input_file)
    ws = wb.active
    
    # 3. Find header row (Assuming header=3, which is row 4 in Excel/openpyxl)
    header_row = 4
    headers = [cell.value for cell in ws[header_row]]
    try:
        prod_col_idx = headers.index('Product name') + 1 # +1 for 1-based indexing
    except ValueError:
        print("Error: 'Product name' column not found.")
        return

    # 4. Iterate through rows starting after header
    for row_num in range(header_row + 1, ws.max_row + 1):
        prod_name = str(ws.cell(row=row_num, column=prod_col_idx).value or "").strip()
        prod_key = prod_name.lower()
        
        if prod_key in stock_map:
            data = stock_map[prod_key]
            # Mapping columns relative to 'Product name' column
            # Adjust these offsets based on your file structure (prod_col_idx + N)
            ws.cell(row=row_num, column=prod_col_idx + 1).value = data.get('Opening Stock')
            ws.cell(row=row_num, column=prod_col_idx + 2).value = data.get('Goods Receipt')
            ws.cell(row=row_num, column=prod_col_idx + 3).value = data.get('Goods Return')
            ws.cell(row=row_num, column=prod_col_idx + 8).value = data.get('Inventory Deducted')
            
           
        if prod_key in stock_closing_map:
            data = stock_closing_map[prod_key]
            # Mapping columns relative to 'Product name' column
            # Adjust these offsets based on your file structure (prod_col_idx + N)
            ws.cell(row=row_num, column=prod_col_idx + 9).value = data.get('Physical')
           

    # 5. Save changes to the same file
    wb.save(input_file)
    
    try:
        prod_col_idx2 = headers2.index('Product name') + 1 # +1 for 1-based indexing
    except ValueError:
        print("Error: 'Product name' column not found for 2.")
        return
    for row_num in range(header_row2 + 1, ws2.max_row + 1):
        prod_name = str(ws2.cell(row=row_num, column=prod_col_idx2).value or "").strip()
        prod_key = prod_name.lower()
        
        if prod_key in stock_map:
            data = stock_map[prod_key]
            # Mapping columns relative to 'Product name' column
            # Adjust these offsets based on your file structure (prod_col_idx2 + N)
            ws2.cell(row=row_num, column=prod_col_idx2 + 1).value = data.get('Closing Stock')
            
            
    wb2.save(input_file_2)
    print(f"Updated {input_file} , {input_file_2} successfully while maintaining format.")

process_inventory()