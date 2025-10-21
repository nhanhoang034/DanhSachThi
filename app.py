from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
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
        
        # Lọc theo thứ tự
        result_data = []
        for idx, code in enumerate(selected, start=1):
            row = df[df['Mã hội viên'] == code]
            if not row.empty:
                quyen = str(row.iloc[0]['Quyền']).strip()
                
                # Tính cấp
                if quyen.startswith('Cấp'):
                    cap_num = int(quyen.replace('Cấp', '').strip()) - 1
                    cap_dang_ky = cap_num
                else:
                    cap_dang_ky = quyen
                
                result_data.append({
                    'STT': idx,
                    'Ma ky thi': exam_code,
                    'Ma Don vi': 'TNIN',
                    'Ma CLB': 'CLB_01102',
                    'Ma hoi vien': code,
                    'Cap dang ky': cap_dang_ky
                })
        
        # Tạo DataFrame
        df_export = pd.DataFrame(result_data)
        
        # Xuất Excel bằng pandas
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Ghi tiêu đề
            workbook = writer.book
            worksheet = workbook.active
            worksheet.title = 'DST'
            
            # Merge cells cho tiêu đề
            worksheet.merge_cells('A1:F2')
            worksheet['A1'] = 'DANH SACH DANG KY THAM DU THI THANG CAP DAI TAEKWONDO CLB_01102'
            worksheet['A1'].alignment = pd.io.formats.excel.Alignment(horizontal='center', vertical='center')
            
            # Ghi dữ liệu từ dòng 3
            df_export.to_excel(writer, sheet_name='DST', index=False, startrow=2, header=True)
        
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