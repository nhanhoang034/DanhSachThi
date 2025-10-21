from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from io import BytesIO
import datetime
import os

app = Flask(__name__)

# Đọc data
DATA_FILE = os.path.join('static', 'data.csv')

@app.route('/')
def index():
    # Đọc CSV KHÔNG CÓ HEADER
    df = pd.read_csv(DATA_FILE, header=None, names=['Tên', 'Mã hội viên', 'Quyền'])
    members = df.to_dict(orient='records')
    return render_template('index.html', members=members)

@app.route('/export', methods=['POST'])
def export():
    selected = request.json.get('selected', [])
    exam_code = request.json.get('exam_code', '').strip().upper()
    
    if not selected or not exam_code:
        return jsonify({'error': 'Thiếu dữ liệu'}), 400
    
    # Đọc CSV KHÔNG CÓ HEADER
    df = pd.read_csv(DATA_FILE, header=None, names=['Tên', 'Mã hội viên', 'Quyền'])
    
    # Lọc theo mã được chọn và giữ nguyên thứ tự chọn
    df_sel_list = []
    for code in selected:
        row = df[df['Mã hội viên'] == code]
        if not row.empty:
            df_sel_list.append(row.iloc[0])
    
    if not df_sel_list:
        return jsonify({'error': 'Không tìm thấy học viên'}), 400
    
    # Tính cấp đăng ký dự thi
    def calculate_cap_dang_ky(quyen):
        quyen_str = str(quyen).strip()
        if quyen_str.startswith('Cấp'):
            try:
                cap_hien_tai = int(quyen_str.replace('Cấp', '').strip())
                return cap_hien_tai - 1
            except:
                return quyen_str
        else:
            return quyen_str
    
    try:
        # Tạo workbook mới
        wb = Workbook()
        ws = wb.active
        ws.title = "DST"
        
        # Thiết lập độ rộng cột
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 25
        ws.column_dimensions['F'].width = 25
        
        # Tiêu đề chính (dòng 1-2)
        ws.merge_cells('A1:F2')
        title_cell = ws['A1']
        title_cell.value = f'DANH SÁCH ĐĂNG KÝ THAM DỰ THI THĂNG CẤP ĐAI TAEKWONDO CLB_01102'
        title_cell.font = Font(name='Times New Roman', size=14, bold=True)
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Header (dòng 3)
        headers = ['STT', 'Mã kỳ thi', 'Mã Đơn vị', 'Mã CLB', 'Mã hội viên', 'Cấp đẳng đăng ký dự thi']
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_num)
            cell.value = header
            cell.font = Font(name='Times New Roman', size=11, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        
        # Ghi dữ liệu (từ dòng 4)
        for idx, row_data in enumerate(df_sel_list, start=1):
            current_row = 3 + idx
            
            # STT
            cell = ws.cell(row=current_row, column=1, value=idx)
            cell.alignment = Alignment(horizontal='center')
            
            # Mã kỳ thi
            cell = ws.cell(row=current_row, column=2, value=exam_code)
            cell.alignment = Alignment(horizontal='center')
            
            # Mã đơn vị
            cell = ws.cell(row=current_row, column=3, value='TNIN')
            cell.alignment = Alignment(horizontal='center')
            
            # Mã CLB
            cell = ws.cell(row=current_row, column=4, value='CLB_01102')
            cell.alignment = Alignment(horizontal='center')
            
            # Mã hội viên
            cell = ws.cell(row=current_row, column=5, value=row_data['Mã hội viên'])
            cell.alignment = Alignment(horizontal='center')
            
            # Cấp đăng ký dự thi
            cell = ws.cell(row=current_row, column=6, value=calculate_cap_dang_ky(row_data['Quyền']))
            cell.alignment = Alignment(horizontal='center')
            
            # Border cho tất cả các ô
            for col in range(1, 7):
                ws.cell(row=current_row, column=col).border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
        
        # Lưu vào BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
    except Exception as e:
        app.logger.error(f"Lỗi tạo Excel: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Lỗi tạo file Excel: {str(e)}'}), 500
    
    filename = f"DST_{exam_code}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    
    return send_file(
        output, 
        as_attachment=True, 
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)