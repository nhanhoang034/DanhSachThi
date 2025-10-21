from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import openpyxl
from io import BytesIO
import datetime
import os

app = Flask(__name__)

# Đọc data
DATA_FILE = os.path.join('static', 'data.csv')
FORM_FILE = os.path.join('static', 'template_form.xlsx')

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
    
    # Lọc theo mã được chọn
    df_sel = df[df['Mã hội viên'].isin(selected)].copy()
    
    if df_sel.empty:
        return jsonify({'error': 'Không tìm thấy học viên'}), 400
    
    # Xếp theo thứ tự người dùng đã chọn
    df_sel['order'] = df_sel['Mã hội viên'].apply(lambda x: selected.index(x) if x in selected else 999)
    df_sel = df_sel.sort_values('order')
    
    # Tạo cột "Cấp đăng ký dự thi"
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
    
    # Tạo DataFrame xuất Excel
    out_df = pd.DataFrame({
        'STT': range(1, len(df_sel) + 1),
        'Mã kỳ thi': [exam_code] * len(df_sel),
        'Mã đơn vị': ['TNIN'] * len(df_sel),
        'Mã CLB': ['CLB_01102'] * len(df_sel),
        'Mã hội viên': df_sel['Mã hội viên'].values,
        'Họ và tên': df_sel['Tên'].values,
        'Cấp hiện tại': df_sel['Quyền'].values,
        'Cấp đẳng đăng ký dự thi': df_sel['Quyền'].apply(calculate_cap_dang_ky).values
    })
    
    # Tạo file Excel
    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            out_df.to_excel(writer, index=False, sheet_name='DST')
        output.seek(0)
    except Exception as e:
        print(f"Lỗi tạo Excel: {e}")
        return jsonify({'error': f'Lỗi tạo file Excel: {str(e)}'}), 500
    
    filename = f"DST_{exam_code}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    
    return send_file(
        output, 
        as_attachment=True, 
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)