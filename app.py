from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from io import BytesIO, StringIO
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
                
                result_data.append({
                    'STT': idx,
                    'Mã kỳ thi': exam_code,
                    'Mã Đơn vị': 'TNIN',
                    'Mã CLB': 'CLB_01102',
                    'Mã hội viên': code,
                    'Cấp đăng ký dự thi': cap_dang_ky
                })
        
        if not result_data:
            return jsonify({'error': 'Không tìm thấy học viên'}), 400
        
        # Tạo DataFrame
        df_export = pd.DataFrame(result_data)
        
        # Xuất CSV với UTF-8 BOM (để Excel mở đúng tiếng Việt)
        output = StringIO()
        df_export.to_csv(output, index=False, encoding='utf-8-sig')
        
        # Convert sang bytes
        csv_bytes = BytesIO(output.getvalue().encode('utf-8-sig'))
        csv_bytes.seek(0)
        
        filename = f"DST_{exam_code}.csv"
        
        return send_file(
            csv_bytes,
            mimetype='text/csv',
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