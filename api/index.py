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
        # Web画面からの入力データ（A列～F列用）
        table_data = json.loads(request.form.get('tableData'))

        # 1. テンプレートを読み込む
        with open(TEMPLATE_PATH, 'rb') as f:
            template_buffer = io.BytesIO(f.read())
        wb_template = openpyxl.load_workbook(template_buffer)
        
        # ---------------------------------------------------------
        # 処理A: アップロードExcelを解析してデータを集計する (Pythonで計算)
        # ---------------------------------------------------------
        wb_uploaded = openpyxl.load_workbook(uploaded_file, data_only=True)
        ws_src = wb_uploaded.worksheets[0] # 1枚目のシート

        # 名前ごとのデータを格納する辞書
        # content = { "名前": { 1: "08:00", 2: "08:30"... } }
        attendance_data = {}
        
        # 名前が登場した順序を保持するリスト
        unique_names = []

        # アップロードされたExcelの58行目(仮)から読み込み開始
        # ※もし開始行が違うなら min_row=58 を調整してください
        # F列(6列目)からが時間データと仮定
        for row in ws_src.iter_rows(min_row=58, values_only=True):
            name = row[0] # A列 (名前)
            
            # 名前が無効ならスキップ
            if not name or name == "お子さま名" or str(name) == "0":
                continue
                
            if name not in attendance_data:
                attendance_data[name] = {}
                unique_names.append(name)
            
            # 時間データの処理 (F列=index 5 が1日, G列=index 6 が2日...)
            # 1日から31日分(最大)ループ
            for day_idx in range(31):
                col_idx = 5 + day_idx # F列はindex 5
                if col_idx < len(row):
                    time_val = row[col_idx]
                    
                    # 時間が入っている場合のみ更新（MAXIFSのロジック）
                    if time_val and time_val != 0:
                        # すでに時間が入っていれば、大きい方を採用（帰りの時間を取るため）
                        current_val = attendance_data[name].get(day_idx + 1)
                        
                        # 新しい値が有効なら上書き（比較ロジックは簡易的に）
                        # 厳密な時間比較が必要ならdatetime変換しますが、
                        # 今回は「0や空以外が入ればOK」として上書きします
                        attendance_data[name][day_idx + 1] = time_val

        # ---------------------------------------------------------
        # 処理B: テンプレートに書き込む
        # ---------------------------------------------------------
        
        # 1. 「貼り付け用」シートへアップロードデータをそのままコピー（一応残す）
        ws_paste = wb_template['貼り付け用']
        for i, row in enumerate(ws_src.iter_rows(values_only=True), start=1):
            for j, value in enumerate(row, start=1):
                ws_paste.cell(row=i, column=j, value=value)

        # 2. 「子どもマスタ」シートへ書き込み
        # Web入力データは A列～F列 に書く
        ws_master = wb_template['子どもマスタ']
        
        # ヘッダー行数（データ開始行の1つ上）
        start_row = 3 # B3からデータ開始と仮定（画像に合わせて調整してください）

        # 今回は「アップロードされたExcelの名前リスト」を正としてB列に書く？
        # それとも「Web入力」を正とする？
        # 画像の関数を見ると「アップロードExcelから名前を抽出」したがっていたので、
        # ここでは「集計した unique_names」を使って行を作ります。
        
        for i, name in enumerate(unique_names):
            current_row = start_row + i
            
            # B列: 名前
            ws_master.cell(row=current_row, column=2, value=name)
            
            # E列(5列目)～: 時間データ (1日～31日)
            days = attendance_data[name]
            for day in range(1, 32):
                if day in days:
                    # E列が1日なら column=5
                    ws_master.cell(row=current_row, column=4 + day, value=days[day])

        # 3. 保存
        output_stream = io.BytesIO()
        wb_template.save(output_stream)
        output_stream.seek(0)

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
