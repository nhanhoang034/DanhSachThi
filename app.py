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
        
        # DEBUG LOG
        app.logger.info(f"===== REQUEST DATA =====")
        app.logger.info(f"Selected: {selected}")
        app.logger.info(f"Exam code: {exam_code}")
        
        if not selected or not exam_code:
            return jsonify({'error': 'Thiếu dữ liệu'}), 400
        
        # Đọc CSV
        df = pd.read_csv(DATA_FILE, header=None, names=['Tên', 'Mã hội viên', 'Quyền'], encoding='utf-8')
        
        # DEBUG: In ra 5 dòng đầu
        app.logger.info(f"===== CSV DATA (5 dòng đầu) =====")
        for idx, row in df.head().iterrows():
            app.logger.info(f"Row {idx}: Tên='{row['Tên']}' | Mã='{row['Mã hội viên']}' | Quyền='{row['Quyền']}'")
        
        # Chuẩn bị dữ liệu - NORMALIZE mã hội viên
        df['Mã hội viên'] = df['Mã hội viên'].astype(str).str.strip()
        
        result_data = []
        for idx, code in enumerate(selected, start=1):
            # Normalize code từ frontend
            code_normalized = str(code).strip()
            
            # DEBUG
            app.logger.info(f"Tìm kiếm mã: '{code_normalized}'")
            
            # So sánh
            row = df[df['Mã hội viên'] == code_normalized]
            
            if not row.empty:
                app.logger.info(f"✓ Tìm thấy: '{code_normalized}'")
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
                    'Mã hội viên': code_normalized,
                    'Cấp đăng ký dự thi': cap_dang_ky
                })
            else:
                # LOG để debug
                app.logger.warning(f"Không tìm thấy mã: '{code_normalized}'")
        
        if not result_data:
            # LOG danh sách mã có trong CSV
            app.logger.error(f"Danh sách mã trong CSV: {df['Mã hội viên'].tolist()}")
            app.logger.error(f"Danh sách mã được chọn: {selected}")
            return jsonify({'error': 'Không tìm thấy học viên'}), 400
        
        # Tạo DataFrame
        df_export = pd.DataFrame(result_data)
        
        # Xuất CSV với UTF-8 BOM (để Excel mở đúng tiếng Việt)
        csv_string = df_export.to_csv(index=False, encoding='utf-8-sig')
        
        # Convert sang bytes
        output = BytesIO()
        output.write(csv_string.encode('utf-8-sig'))
        output.seek(0)
        
        filename = f"DST_{exam_code}.csv"
        
        # LOG để debug
        app.logger.info(f"===== XUẤT FILE =====")
        app.logger.info(f"Filename: {filename}")
        app.logger.info(f"Mimetype: text/csv")
        app.logger.info(f"Số học viên: {len(result_data)}")
        
        # Force CSV download
        response = send_file(
            output,
            mimetype='text/csv',
            as_attachment=True,
            download_name=filename
        )
        response.headers['Content-Type'] = 'text/csv; charset=utf-8'
        response.headers['Content-Disposition'] = f'attachment; filename="{filename}"'
        
        return response
        
    except Exception as e:
        app.logger.error(f"Lỗi: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)