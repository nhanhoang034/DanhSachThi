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

def get_cap_sort_key(quyen):
    """Trả về key để sắp xếp theo cấp (9->1, sau đó các Đẳng, cuối cùng là khác)"""
    quyen_str = str(quyen).strip()
    
    # Xử lý "Cấp X"
    if quyen_str.startswith('Cấp'):
        try:
            cap_num = int(quyen_str.replace('Cấp', '').strip())
            return (0, -cap_num)  # Group 0, sắp xếp giảm dần (9->1)
        except:
            pass
    
    # Xử lý "X Đẳng"
    if 'Đẳng' in quyen_str:
        try:
            dang_num = int(quyen_str.replace('Đẳng', '').strip())
            return (1, dang_num)  # Group 1, sắp xếp tăng dần (1->3)
        except:
            pass
    
    # Các trường hợp khác
    return (2, 0)

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
        
        # Tạo dict để lưu thông tin học viên và thứ tự chọn
        result_data = []
        for selection_order, code in enumerate(selected):
            code_normalized = str(code).strip()
            row = df[df['Mã hội viên'] == code_normalized]
            
            if not row.empty:
                quyen = str(row.iloc[0]['Quyền']).strip()
                
                # Tính cấp đăng ký (Cấp X -> X-1)
                if quyen.startswith('Cấp'):
                    try:
                        cap_num = int(quyen.replace('Cấp', '').strip()) - 1
                        cap_dang_ky = cap_num
                    except:
                        cap_dang_ky = quyen
                else:
                    cap_dang_ky = quyen
                
                result_data.append({
                    'Quyền': quyen,  # Để sort
                    'selection_order': selection_order,  # Thứ tự chọn
                    'Mã kỳ thi': exam_code,
                    'Mã Đơn vị': 'TNIN',
                    'Mã CLB': 'CLB_01102',
                    'Mã hội viên': code_normalized,
                    'Cấp đăng ký dự thi': cap_dang_ky
                })
        
        if not result_data:
            return jsonify({'error': 'Không tìm thấy học viên'}), 400
        
        # Sắp xếp: theo cấp (9->1), trong cùng cấp thì theo thứ tự chọn
        result_data.sort(key=lambda x: (get_cap_sort_key(x['Quyền']), x['selection_order']))
        
        # Thêm STT sau khi đã sắp xếp
        for idx, item in enumerate(result_data, start=1):
            item['STT'] = idx
        
        # Tạo DataFrame với thứ tự cột đúng
        df_export = pd.DataFrame(result_data)
        df_export = df_export[['STT', 'Mã kỳ thi', 'Mã Đơn vị', 'Mã CLB', 'Mã hội viên', 'Cấp đăng ký dự thi']]
        
        # Xuất Excel với XlsxWriter
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Ghi DataFrame từ dòng 3 (dòng 1: title, dòng 2: trống, dòng 3: header)
            df_export.to_excel(writer, sheet_name='Sheet1', index=False, startrow=2, header=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Format cho tiêu đề (CHỈ merge A1:F1, không merge dòng 2)
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
            
            # Merge cells cho title CHỈ Ở DÒNG 1 (A1:F1)
            worksheet.merge_range('A1:F1', 
                'DANH SÁCH ĐĂNG KÝ THAM DỰ THI THĂNG CẤP ĐAI TAEKWONDO CLB_01102', 
                title_format)
            
            # Set độ rộng cột (điều chỉnh cho vừa A4)
            worksheet.set_column('A:A', 5)   # STT - nhỏ
            worksheet.set_column('B:B', 10)  # Mã kỳ thi - nhỏ
            worksheet.set_column('C:C', 10)  # Mã Đơn vị - nhỏ
            worksheet.set_column('D:D', 12)  # Mã CLB - nhỏ
            worksheet.set_column('E:E', 22)  # Mã hội viên - quan trọng
            worksheet.set_column('F:F', 20)  # Cấp đăng ký - quan trọng
            
            # Set chiều cao dòng
            worksheet.set_row(0, 30)  # Dòng title cao hơn
            
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
            
            # Cài đặt in - VỪA 1 TRANG A4
            worksheet.set_paper(9)  # A4 paper
            worksheet.fit_to_pages(1, 0)  # Vừa 1 trang ngang, tự động chiều dọc
            worksheet.set_portrait()  # In dọc
            worksheet.set_margins(left=0.5, right=0.5, top=0.75, bottom=0.75)
        
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