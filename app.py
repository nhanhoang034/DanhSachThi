from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from io import BytesIO
import datetime
import os

app = Flask(__name__)

DATA_FILE = os.path.join('static', 'data.csv')

@app.route('/')
def index():
    try:
        df = pd.read_csv(DATA_FILE, header=None, names=['Tên', 'Mã hội viên', 'Quyền'], encoding='utf-8')
        members = df.to_dict(orient='records')
        return render_template('index.html', members=members)
    except Exception as e:
        app.logger.error(f"Lỗi: {e}")
        return render_template('index.html', members=[])

@app.route('/export', methods=['POST'])
def export():
    try:
        data = request.get_json()
        selected = data.get('selected', [])
        exam_code = data.get('exam_code', '').strip()
        
        if not selected or not exam_code:
            return jsonify({'error': 'Thiếu dữ liệu'}), 400
        
        # Đọc CSV
        df = pd.read_csv(DATA_FILE, header=None, names=['Tên', 'Mã hội viên', 'Quyền'], encoding='utf-8')
        
        # Chuẩn bị dữ liệu
        result_data = []
        for idx, code in enumerate(selected, start=1):
            row = df[df['Mã hội viên'] == code]
            if not row.empty:
                quyen = str(row.iloc[0]['Quyền']).strip()
                
                # Tính cấp (chỉ số, bỏ chữ "Cấp")
                if quyen.startswith('Cấp'):
                    try:
                        cap_num = int(quyen.replace('Cấp', '').strip()) - 1
                        cap_dang_ky = cap_num
                    except:
                        cap_dang_ky = quyen
                else:
                    cap_dang_ky = quyen
                
                result_data.append([
                    idx,
                    exam_code,
                    'TNIN',
                    'CLB_01102',
                    code,
                    cap_dang_ky
                ])
        
        if not result_data:
            return jsonify({'error': 'Không tìm thấy học viên'}), 400
        
        # Tạo workbook thủ công
        wb = Workbook()
        ws = wb.active
        ws.title = 'DST'
        
        # Độ rộng cột
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 20
        ws.column_dimensions['F'].width = 28
        
        # Border
        thin = Side(style='thin')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        
        # Tiêu đề (A1:F2)
        ws.merge_cells('A1:F2')
        ws['A1'] = 'DANH SACH DANG KY THAM DU THI THANG CAP DAI TAEKWONDO CLB_01102'
        ws['A1'].font = Font(size=14, bold=True)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # Header (row 3)
        headers = ['STT', 'Ma ky thi', 'Ma Don vi', 'Ma CLB', 'Ma hoi vien', 'Cap dang ky du thi']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col)
            cell.value = header
            cell.font = Font(size=11, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        
        # Data (row 4+)
        for row_idx, row_data in enumerate(result_data, 4):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = value
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
        
        # Lưu
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