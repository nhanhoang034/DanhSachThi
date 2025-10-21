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
    try:
        # Đọc CSV KHÔNG CÓ HEADER
        df = pd.read_csv(DATA_FILE, header=None, names=['Tên', 'Mã hội viên', 'Quyền'], encoding='utf-8')
        members = df.to_dict(orient='records')
        return render_template('index.html', members=members)
    except Exception as e:
        app.logger.error(f"Lỗi đọc CSV: {e}")
        return render_template('index.html', members=[])

@app.route('/export', methods=['POST'])
def export():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Không có dữ liệu'}), 400
            
        selected = data.get('selected', [])
        exam_code = data.get('exam_code', '').strip()
        
        if not selected:
            return jsonify({'error': 'Chưa chọn học viên nào'}), 400
        if not exam_code:
            return jsonify({'error': 'Chưa nhập mã kỳ thi'}), 400
        
        # Đọc CSV
        df = pd.read_csv(DATA_FILE, header=None, names=['Tên', 'Mã hội viên', 'Quyền'], encoding='utf-8')
        
        # Lọc theo thứ tự chọn
        df_sel_list = []
        for code in selected:
            row = df[df['Mã hội viên'] == str(code)]
            if not row.empty:
                df_sel_list.append(row.iloc[0])
        
        if not df_sel_list:
            return jsonify({'error': 'Không tìm thấy học viên'}), 400
        
        # Tính cấp đăng ký (chỉ trả về số)
        def calculate_cap(quyen):
            quyen_str = str(quyen).strip()
            if quyen_str.startswith('Cấp'):
                try:
                    cap = int(quyen_str.replace('Cấp', '').strip())
                    return cap - 1
                except:
                    return quyen_str
            return quyen_str
        
        # Tạo workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "DST"
        
        # Độ rộng cột
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 20
        ws.column_dimensions['F'].width = 28
        
        # Định nghĩa style border
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Tiêu đề (dòng 1-2)
        ws.merge_cells('A1:F2')
        title = ws['A1']
        title.value = 'DANH SACH DANG KY THAM DU THI THANG CAP DAI TAEKWONDO CLB_01102'
        title.font = Font(size=14, bold=True)
        title.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # Header (dòng 3)
        headers = ['STT', 'Ma ky thi', 'Ma Don vi', 'Ma CLB', 'Ma hoi vien', 'Cap dang dang ky du thi']
        
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = Font(size=11, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
        
        # Dữ liệu (từ dòng 4)
        for idx, row_data in enumerate(df_sel_list, start=1):
            row_num = 3 + idx
            
            # Dữ liệu cho từng cột
            data_row = [
                idx,                                    # STT
                exam_code,                              # Mã kỳ thi
                'TNIN',                                 # Mã đơn vị
                'CLB_01102',                            # Mã CLB
                str(row_data['Mã hội viên']),          # Mã hội viên
                calculate_cap(row_data['Quyền'])       # Cấp (số)
            ]
            
            for col_idx, value in enumerate(data_row, 1):
                cell = ws.cell(row=row_num, column=col_idx, value=value)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
        
        # Lưu vào BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"DST_{exam_code}.xlsx"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        app.logger.error(f"Lỗi: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)