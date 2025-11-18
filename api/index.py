from flask import Flask, request, send_file
import io
import os
import json
import csv
import datetime

app = Flask(__name__)

# テンプレートCSV（項目名が入ったファイル）のパス
FORMAT_CSV_PATH = os.path.join(os.path.dirname(__file__), 'format.csv')

@app.route('/api/process', methods=['POST'])
def process_excel():
    try:
        # 1. Web上の表データを取得
        # （Excelファイルも受け取りますが、CSV作成にはWebデータを使います）
        if 'file' not in request.files:
            return {"error": "No file uploaded"}, 400
        
        # フロントエンドから送られた表データ（JSON）をリストに変換
        table_data = json.loads(request.form.get('tableData'))

        # 2. 出力用のCSVをメモリ上に作成
        csv_output = io.StringIO()
        writer = csv.writer(csv_output)

        # 3. templateとなる format.csv があれば、まずそのヘッダー（1行目）を書き込む
        if os.path.exists(FORMAT_CSV_PATH):
            with open(FORMAT_CSV_PATH, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                for row in reader:
                    writer.writerow(row)
                    # 基本は1行目だけだと思いますが、もし複数行あれば全部コピーされます
        else:
            # format.csvが無い場合の予備（空のヘッダーなどを入れるか、何もしない）
            pass

        # 4. Webで入力されたデータ（値）をそのまま追記（これが「値貼り付け」になります）
        for row_data in table_data:
            writer.writerow(row_data)

        # 5. ファイル作成処理
        # 先頭に戻す
        csv_output.seek(0)
        
        # 文字化け防止のため BOM付きUTF-8 に変換
        byte_output = io.BytesIO(csv_output.getvalue().encode('utf-8-sig'))

        # ファイル名生成（サーバー側でも指定しますが、最終決定はJS側で行います）
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
