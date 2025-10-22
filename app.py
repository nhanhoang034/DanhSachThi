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
        
        # DEBUG LOG
        app.logger.info(f"===== REQUEST DATA =====")
        app.logger.info(f"Selected: {selected}")
        app.logger.info(f"Exam code: {exam_code}")
        
        if not selected or not exam_code:
            return jsonify({'error': 'Thiếu dữ liệu'}), 400
        
        # Đọc CSV
        df = pd.read_csv(DATA_FILE, header=None, encoding='utf-8')
        
        # DEBUG: In ra thông tin CSV
        app.logger.info(f"===== CSV INFO =====")
        app.logger.info(f"Số cột: {len(df.columns)}")
        app.logger.info(f"Tên cột: {df.columns.tolist()}")
        app.logger.info(f"===== CSV DATA (3 dòng đầu RAW) =====")
        for idx in range(min(3, len(df))):
            app.logger.info(f"Row {idx}: {df.iloc[idx].tolist()}")
        
        # Gán tên cột (giả định: cột 0=Tên, cột 1=Mã, cột 2=Quyền)
        if len(df.columns) >= 3:
            df.columns = ['Tên', 'Mã hội viên', 'Quyền'] + [f'Extra_{i}' for i in range(len(df.columns) - 3)]
        else:
            app.logger.error(f"CSV không đủ 3 cột! Chỉ có {len(df.columns)} cột")
            return jsonify({'error': 'File CSV không đúng định dạng'}), 400
        
        app.logger.info(f"===== CSV DATA (sau khi gán tên cột) =====")
        for idx, row in df.head(3).iterrows():
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
        
        # LOG để debug
        app.logger.info(f"===== XUẤT FILE =====")
        app.logger.info(f"Số học viên: {len(result_data)}")
        
        # Xuất Excel với XlsxWriter engine
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Ghi DataFrame từ dòng 3 (để dành chỗ cho title)
            df_export.to_excel(writer, sheet_name='DST', index=False, startrow=2, header=False)
            
            # Lấy workbook và worksheet
            workbook = writer.book
            worksheet = writer.sheets['DST']
            
            # Format cho title
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
                'font_name': 'Times New Roman'
            })
            
            # Format cho data
            data_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'font_name': 'Times New Roman'
            })
            
            # Merge cells cho title (A1:E2) - 5 cột không có STT
            worksheet.merge_range('A1:E2', 
                'DANH SÁCH ĐĂNG KÝ THAM DỰ THI THĂNG CẤP ĐAI TAEKWONDO CLB_01102', 
                title_format)
            
            # Set độ rộng cột
            worksheet.set_column('A:A', 18)  # Mã kỳ thi
            worksheet.set_column('B:B', 15)  # Mã Đơn vị
            worksheet.set_column('C:C', 15)  # Mã CLB
            worksheet.set_column('D:D', 25)  # Mã hội viên
            worksheet.set_column('E:E', 28)  # Cấp đăng ký dự thi
            
            # Ghi header (row 3)
            headers = ['Mã kỳ thi', 'Mã Đơn vị', 'Mã CLB', 'Mã hội viên', 'Cấp đăng ký dự thi']
            for col_num, value in enumerate(headers):
                worksheet.write(2, col_num, value, header_format)
            
            # Apply format cho data cells (từ row 4)
            for row_num in range(len(df_export)):
                for col_num in range(len(df_export.columns)):
                    worksheet.write(row_num + 3, col_num, 
                                  df_export.iloc[row_num, col_num], 
                                  data_format)
        
        output.seek(0)
        
        filename = f"DST_{exam_code}.xlsx"
        
        app.logger.info(f"Filename: {filename}")
        app.logger.info(f"File size: {len(output.getvalue())} bytes")
        
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