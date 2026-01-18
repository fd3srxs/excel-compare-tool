import os
from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Color
from copy import copy
import io

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

def compare_excels(file1, file2):
    wb1 = load_workbook(file1, data_only=True)
    wb2 = load_workbook(file2, data_only=True)
    
    # Create a new workbook for output
    output_wb = Workbook()
    output_wb.remove(output_wb.active) # Remove default sheet
    
    red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
    green_fill = PatternFill(start_color='FF00FF00', end_color='FF00FF00', fill_type='solid')
    red_font = Font(color='FFFF0000')

    # Iterate through sheets in the first workbook
    for sheet_name in wb1.sheetnames:
        ws1 = wb1[sheet_name]
        
        # Create sheet in output
        ws_out = output_wb.create_sheet(title=sheet_name)
        
        # Check if sheet exists in second workbook
        ws2 = None
        if sheet_name in wb2.sheetnames:
            ws2 = wb2[sheet_name]
            
        # Strategy: Scan rows to find "ID" column
        id_col_idx = None # 1-based index
        header_row_idx = None # 1-based index
        
        # Scan first 20 rows (or max rows) to find header
        max_scan_rows = min(20, ws1.max_row)
        for r_idx in range(1, max_scan_rows + 1):
            row = ws1[r_idx]
            for cell in row:
                if cell.value:
                    val_str = str(cell.value).lower()
                    if "id" in val_str or "sku" in val_str or "#" in val_str:
                        id_col_idx = cell.column
                        header_row_idx = cell.row
                        break
            if id_col_idx:
                break
        
        # If ID column found and ws2 exists, build index for ws2
        # We assume the header in ws2 is at the same row index as ws1 ?? 
        # Or should we just scan ws2 too? 
        # For now, let's assume data starts after header_row_idx.
        
        ws2_index = {} # Map ID -> Row Object (or Row Index)
        if id_col_idx and ws2 and header_row_idx:
            # Iterate rows in ws2 starting from header_row_idx + 1?
            # Or just iterate all rows and skip if row index <= header_row_idx?
            # Actually, ws2 might have different header position?
            # Let's assume simplest case: structure is similar.
            
            for row in ws2.iter_rows(min_row=header_row_idx + 1):
                # Get the cell in the ID column
                # row is a tuple of cells. Index is col_idx - 1
                try:
                    # Check if the row is long enough
                    if len(row) >= id_col_idx:
                        id_cell = row[id_col_idx - 1]
                        val = id_cell.value
                        if val is not None:
                            ws2_index[val] = row
                except IndexError:
                    pass

        # Iterate through rows and columns of File 1
        for row in ws1.iter_rows():
            current_row_idx = row[0].row
            
            # Find matching row in File 2
            row2 = None
            if ws2:
                # Only try to match data rows (rows after the header)
                if header_row_idx and current_row_idx > header_row_idx:
                     if id_col_idx and ws2_index:
                        # Look up by ID
                        try:
                             if len(row) >= id_col_idx:
                                id_val = row[id_col_idx - 1].value
                                if id_val in ws2_index:
                                    row2 = ws2_index[id_val]
                        except:
                            pass
                     else:
                        # Fallback to positional match
                         try:
                            row2 = ws2[current_row_idx]
                         except IndexError:
                            row2 = None
                else:
                    # For header row or pre-header rows, try to match by position
                    if header_row_idx and current_row_idx == header_row_idx:
                         # This is the header row.
                         # Try to find header row in ws2?
                         # For now, simplistic: match same row index
                         try:
                             row2 = ws2[current_row_idx]
                         except IndexError:
                             row2 = None
                    else:
                        # Pre-header rows
                        try:
                             row2 = ws2[current_row_idx]
                        except IndexError:
                             row2 = None


            for i, cell in enumerate(row):
                # Copy value to output
                new_cell = ws_out.cell(row=cell.row, column=cell.column, value=cell.value)
                
                # Apply styles
                if header_row_idx and cell.row == header_row_idx:
                    new_cell.fill = green_fill
                elif id_col_idx and cell.column == id_col_idx:
                     # Only highlight ID column in data rows
                    if header_row_idx and cell.row > header_row_idx:
                        new_cell.fill = yellow_fill
                
                # Compare
                is_diff = False
                if ws2:
                    if row2:
                        try:
                            # Verify row2 has this column
                            # i is 0-based index of cell in row
                            if i < len(row2):
                                cell2 = row2[i]
                                if cell.value != cell2.value:
                                    is_diff = True
                            else:
                                is_diff = True # Column missing in row2
                        except:
                            is_diff = True
                    else:
                        is_diff = True # Row not found in File 2
                else:
                     is_diff = True # Sheet doesn't exist in file 2
                
                if is_diff:
                    # Make value red if different
                    new_cell.font = red_font

    return output_wb

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/compare', methods=['POST'])
def compare():
    if 'file1' not in request.files or 'file2' not in request.files:
        return 'No files uploaded', 400
    
    file1 = request.files['file1']
    file2 = request.files['file2']
    
    if file1.filename == '' or file2.filename == '':
        return 'No selected file', 400

    if file1 and file2:
        output_wb = compare_excels(file1, file2)
        
        output_stream = io.BytesIO()
        output_wb.save(output_stream)
        output_stream.seek(0)
        
        return send_file(
            output_stream,
            as_attachment=True,
            download_name='comparison_result.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

if __name__ == '__main__':
    app.run(debug=True, port=5000)
