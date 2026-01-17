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
    red_font = Font(color='FFFF0000')

    # Iterate through sheets in the first workbook
    for sheet_name in wb1.sheetnames:
        ws1 = wb1[sheet_name]
        
        # Create sheet in output
        ws_out = output_wb.create_sheet(title=sheet_name)
        
        # Check if sheet exists in second workbook
        if sheet_name in wb2.sheetnames:
            ws2 = wb2[sheet_name]
        else:
            ws2 = None
            
        # Iterate through rows and columns
        for row in ws1.iter_rows():
            for cell in row:
                # Copy value to output
                new_cell = ws_out.cell(row=cell.row, column=cell.column, value=cell.value)
                
                # Compare
                is_diff = False
                if ws2:
                    try:
                        cell2 = ws2.cell(row=cell.row, column=cell.column)
                        if cell.value != cell2.value:
                            is_diff = True
                    except:
                        is_diff = True # Cell doesn't exist in file 2
                else:
                     is_diff = True # Sheet doesn't exist in file 2
                
                if is_diff:
                    # Make value red if different
                    new_cell.font = red_font
                    # new_cell.fill = red_fill # Option: use fill instead of font? User asked for "make the value red", so font color is safer interpretation, but maybe both? Let's stick to font color as requested "make the value red".

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
