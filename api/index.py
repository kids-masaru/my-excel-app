from flask import Flask, request, send_file
import openpyxl
import io
import os
import json

app = Flask(__name__)

# テンプレートファイルのパス
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.xlsx')

@app.route('/api/process', methods=['POST'])
def process_excel():
    try:
        # 1. アップロードされたファイルを取得
        if 'file' not in request.files:
            return {"error": "No file uploaded"}, 400
        
        uploaded_file = request.files['file']
        original_filename = uploaded_file.filename
        
        # 2. 手入力された表データを取得 (JSON文字列として送られてくる)
        table_data_json = request.form.get('tableData')
        if not table_data_json:
            return {"error": "No table data"}, 400
        
        table_data = json.loads(table_data_json) # リストのリストに変換

        # 3. テンプレートを読み込む
        # Vercel等の環境ではバイナリモードで開くのが安全
        with open(TEMPLATE_PATH, 'rb') as f:
            template_buffer = io.BytesIO(f.read())
        
        wb_template = openpyxl.load_workbook(template_buffer)

        # --- 処理A: アップロードされたExcelを「貼り付け用」シート(Sheet1)へ ---
        # アップロードされたファイルをメモリ上で開く
        wb_uploaded = openpyxl.load_workbook(uploaded_file)
        ws_uploaded = wb_uploaded.worksheets[0] # 1枚目のシート
        
        # テンプレートの「貼り付け用」シート（1枚目と仮定、名前指定も可）
        # 名前で指定する場合: ws_paste = wb_template['貼り付け用']
        ws_paste = wb_template.worksheets[0] 

        # アップロードされたデータの値をそのまま転記 (A1から)
        # iter_rowsを使って値をコピー
        for i, row in enumerate(ws_uploaded.iter_rows(values_only=True), start=1):
            for j, value in enumerate(row, start=1):
                ws_paste.cell(row=i, column=j, value=value)

        # --- 処理B: 手入力データを「子どもマスタ」シート(Sheet2)へ ---
        # テンプレートの「子どもマスタ」シート（2枚目と仮定）
        # 名前で指定する場合: ws_master = wb_template['子どもマスタ']
        ws_master = wb_template.worksheets[1]

        # A2から貼り付け (行: data_row_index + 2, 列: col_index + 1)
        # table_data は [ ["名前", "カナ", ...], ["名前", ...], ... ] の形式
        for row_idx, row_data in enumerate(table_data):
            for col_idx, value in enumerate(row_data):
                # 空白処理などがもし必要ならここで
                ws_master.cell(row=row_idx + 2, column=col_idx + 1, value=value)

        # 4. 編集したファイルをメモリに保存して返却
        output_stream = io.BytesIO()
        wb_template.save(output_stream)
        output_stream.seek(0)
        
        # ファイル名を生成
        base_name = os.path.splitext(original_filename)[0]
        download_name = f"{base_name}_complete.xlsx"

        return send_file(
            output_stream,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        return {"error": str(e)}, 500
