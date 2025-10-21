from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import openpyxl
from io import BytesIO
import datetime

app = Flask(__name__)

# Đọc data
DATA_FILE = "data.csv"
FORM_FILE = "template_form.xlsx"

@app.route('/')
def index():
    df = pd.read_csv(DATA_FILE)
    # Giả sử file data có cột: "Tên", "Mã hội viên", "Quyền"
    members = df.to_dict(orient='records')
    return render_template('index.html', members=members)

@app.route('/export', methods=['POST'])
def export():
    selected = request.json.get('selected', [])
    exam_code = request.json.get('exam_code', '').strip().upper()

    if not selected or not exam_code:
        return jsonify({'error': 'Thiếu dữ liệu'}), 400

    # Đọc data gốc
    df = pd.read_csv(DATA_FILE)
    df_sel = df[df['Mã hội viên'].isin(selected)].copy()

    # Xếp theo thứ tự người dùng đã chọn
    df_sel['order'] = df_sel['Mã hội viên'].apply(lambda x: selected.index(x))
    df_sel = df_sel.sort_values('order')

    # Tạo DataFrame xuất
    out_df = pd.DataFrame({
        'STT': range(1, len(df_sel) + 1),
        'Mã kỳ thi': [exam_code] * len(df_sel),
        'Mã đơn vị': ['TNIN'] * len(df_sel),
        'Mã CLB': ['CLB_01102'] * len(df_sel),
        'Mã hội viên': df_sel['Mã hội viên'],
        'Cấp đẳng đăng ký dự thi': df_sel['Quyền'].apply(lambda q: int(q.split()[-1]) - 1 if str(q).startswith('Cấp') else ''),
        'Cấp hiện tại': df_sel['Quyền']
    })

    # Ghi vào Excel trong bộ nhớ
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        out_df.to_excel(writer, index=False, sheet_name='DST')

    output.seek(0)
    filename = f"DST_{exam_code}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"

    return send_file(output, as_attachment=True, download_name=filename)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
