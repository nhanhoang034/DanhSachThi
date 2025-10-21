from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from io import BytesIO
import datetime
import os

app = Flask(__name__)

# Đọc data
DATA_FILE = os.path.join('static', 'data.csv')
TEMPLATE_FILE = os.path.join('static', 'template_form.xlsx')

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
    
    df_sel = pd.DataFrame(df_sel_list)
    
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
            # Giữ nguyên nếu là Đẳng hoặc GV
            return quyen_str
    
    try:
        # Load template Excel
        wb = load_workbook(TEMPLATE_FILE)
        ws = wb.active
        
        # Tìm dòng bắt đầu ghi dữ liệu (sau header)
        start_row = 4  # Dòng 4 trong Excel (index 0 = dòng 1)
        
        # Ghi dữ liệu vào template
        for idx, (_, row) in enumerate(df_sel.iterrows(), start=1):
            current_row = start_row + idx - 1
            
            ws.cell(row=current_row, column=1, value=idx)  # STT
            ws.cell(row=current_row, column=2, value=exam_code)  # Mã kỳ thi
            ws.cell(row=current_row, column=3, value='TNIN')  # Mã đơn vị
            ws.cell(row=current_row, column=4, value='CLB_01102')  # Mã CLB
            ws.cell(row=current_row, column=5, value=row['Mã hội viên'])  # Mã hội viên
            ws.cell(row=current_row, column=6, value=calculate_cap_dang_ky(row['Quyền']))  # Cấp đăng ký dự thi
        
        # Lưu vào BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
    except Exception as e:
        app.logger.error(f"Lỗi tạo Excel: {e}")
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