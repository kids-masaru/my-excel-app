from flask import Flask, request, send_file
import openpyxl
import io
import os
import json
import datetime

app = Flask(__name__)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.xlsx')

@app.route('/api/process', methods=['POST'])
def process_excel():
    try:
        if 'file' not in request.files:
            return {"error": "No file uploaded"}, 400
        
        uploaded_file = request.files['file']
        table_data = json.loads(request.form.get('tableData'))

        # テンプレートを読み込む
        with open(TEMPLATE_PATH, 'rb') as f:
            template_buffer = io.BytesIO(f.read())
        
        # templateを開く（数式等を壊さないよう data_only=False）
        wb_template = openpyxl.load_workbook(template_buffer)

        # 1. アップロードされたExcelを「貼り付け用」シートへ
        ws_paste = wb_template['貼り付け用']
        
        # アップロードファイルは「値」だけを読み取る
        wb_uploaded = openpyxl.load_workbook(uploaded_file, data_only=True)
        ws_uploaded = wb_uploaded.worksheets[0]

        for i, row in enumerate(ws_uploaded.iter_rows(values_only=True), start=1):
            for j, value in enumerate(row, start=1):
                ws_paste.cell(row=i, column=j, value=value)

        # 2. Web上のデータを「子どもマスタ」シートへ
        ws_master = wb_template['子どもマスタ']
        
        for row_idx, row_data in enumerate(table_data):
            for col_idx, value in enumerate(row_data):
                ws_master.cell(row=row_idx + 2, column=col_idx + 1, value=value)
                
        wb_template.properties.calcId = None
        
        # 3. Excelとして保存
        output_stream = io.BytesIO()
        wb_template.save(output_stream)
        output_stream.seek(0)

        # ファイル名生成 (complete_20251119.xlsx)
        today_str = datetime.datetime.now().strftime('%Y%m%d')
        download_name = f"complete_{today_str}.xlsx"

        return send_file(
            output_stream,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        return {"error": str(e)}, 500
