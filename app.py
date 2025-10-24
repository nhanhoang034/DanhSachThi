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
        df = pd.read_csv(DATA_FILE, header=None, encoding='utf-8')
        
        # Gán tên cột
        if len(df.columns) >= 3:
            df.columns = ['Tên', 'Mã hội viên', 'Quyền'] + [f'Extra_{i}' for i in range(len(df.columns) - 3)]
        else:
            return jsonify({'error': 'File CSV không đúng định dạng'}), 400
        
        # Chuẩn bị dữ liệu
        df['Mã hội viên'] = df['Mã hội viên'].astype(str).str.strip()
        
        result_data = []
        for idx, code in enumerate(selected, start=1):
            code_normalized = str(code).strip()
            row = df[df['Mã hội viên'] == code_normalized]
            
            if not row.empty:
                quyen = str(row.iloc[0]['Quyền']).strip()
                
                # Tính cấp đăng ký
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
        
        if not result_data:
            return jsonify({'error': 'Không tìm thấy học viên'}), 400
        
        # Tạo DataFrame
        df_export = pd.DataFrame(result_data)
        
        # Xuất Excel với XlsxWriter
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Ghi DataFrame từ dòng 4 (để dành chỗ cho title và header)
            df_export.to_excel(writer, sheet_name='Sheet1', index=False, startrow=3, header=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Format cho tiêu đề (merge A1:F2)
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 14,
                'align': 'center',
                'valign': 'vcenter',
                'font_name': 'Times New Roman'
            })
            
            # Format cho header
            header_format = workbook.add_format({
                'bold': True,
                'font_size': 11,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'font_name': 'Times New Roman',
                'text_wrap': True
            })
            
            # Format cho data
            data_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'font_name': 'Times New Roman'
            })
            
            # Merge cells cho title (A1:F2)
            worksheet.merge_range('A1:F2', 
                'DANH SÁCH ĐĂNG KÝ THAM DỰ THI THĂNG CẤP ĐAI TAEKWONDO CLB_01102', 
                title_format)
            
            # Set độ rộng cột
            worksheet.set_column('A:A', 8)   # STT
            worksheet.set_column('B:B', 15)  # Mã kỳ thi
            worksheet.set_column('C:C', 15)  # Mã Đơn vị
            worksheet.set_column('D:D', 15)  # Mã CLB
            worksheet.set_column('E:E', 25)  # Mã hội viên
            worksheet.set_column('F:F', 28)  # Cấp đăng ký dự thi
            
            # Ghi header (row 3, index 2)
            headers = ['STT', 'Mã kỳ thi', 'Mã Đơn vị', 'Mã CLB', 'Mã hội viên', 'Cấp đăng ký dự thi']
            for col_num, value in enumerate(headers):
                worksheet.write(2, col_num, value, header_format)
            
            # Apply format cho data cells (từ row 4, index 3)
            for row_num in range(len(df_export)):
                for col_num in range(len(df_export.columns)):
                    worksheet.write(row_num + 3, col_num, 
                                  df_export.iloc[row_num, col_num], 
                                  data_format)
        
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