from flask import Flask, request, send_file
import openpyxl
import io
import os
import json
import csv
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

        # テンプレート読み込み
        with open(TEMPLATE_PATH, 'rb') as f:
            template_buffer = io.BytesIO(f.read())
        wb_template = openpyxl.load_workbook(template_buffer)

        # 1. アップロードExcelを「貼り付け用」へ転記
        wb_uploaded = openpyxl.load_workbook(uploaded_file)
        ws_uploaded = wb_uploaded.worksheets[0]
        ws_paste = wb_template['貼り付け用'] # シート名指定に変更

        for i, row in enumerate(ws_uploaded.iter_rows(values_only=True), start=1):
            for j, value in enumerate(row, start=1):
                ws_paste.cell(row=i, column=j, value=value)

        # 2. Web上の表データを「子どもマスタ」へ転記
        ws_master = wb_template['子どもマスタ'] # シート名指定に変更
        for row_idx, row_data in enumerate(table_data):
            for col_idx, value in enumerate(row_data):
                ws_master.cell(row=row_idx + 2, column=col_idx + 1, value=value)

        # -------------------------------------------------------
        # 3. ここが変わりました: 「CSVフォーム」シートをCSVとして出力
        # -------------------------------------------------------
        
        # 「CSVフォーム」シートを取得 (名前はExcelと完全に一致させてください)
        if 'CSVフォーム' not in wb_template.sheetnames:
             return {"error": "template.xlsxに「CSVフォーム」シートが見つかりません"}, 500
             
        ws_csv_source = wb_template['CSVフォーム']

        # メモリ上でCSV書き込み
        csv_output = io.StringIO()
        writer = csv.writer(csv_output)

        # データがある範囲だけ書き出す
        for row in ws_csv_source.iter_rows(values_only=True):
            # 行の中身が全部Noneならスキップする処理を入れるとより丁寧ですが、
            # Excelの数式が入っている場合はそのまま書き出します
            writer.writerow(row)
        
        # 文字列ポインタを先頭に戻す
        csv_output.seek(0)

        # UTF-8 (BOM付き) のバイナリに変換
        # ※これがないとExcelで開いたとき日本語が文字化けします
        byte_output = io.BytesIO(csv_output.getvalue().encode('utf-8-sig'))

        # 今日の日付を取得 (例: 20251119)
        today_str = datetime.datetime.now().strftime('%Y%m%d')
        download_name = f"complete_{today_str}.csv"

        return send_file(
            byte_output,
            as_attachment=True,
            download_name=download_name,
            mimetype='text/csv'
        )

    except Exception as e:
        return {"error": str(e)}, 500
