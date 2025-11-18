// 画面読み込み時に表(6列x30行)を生成
document.addEventListener('DOMContentLoaded', () => {
    const tbody = document.querySelector('#inputTable tbody');
    const ROWS = 30;
    const COLS = 6;

    for (let i = 0; i < ROWS; i++) {
        const tr = document.createElement('tr');
        for (let j = 0; j < COLS; j++) {
            const td = document.createElement('td');
            const input = document.createElement('input');
            input.type = 'text';
            
            // 貼り付けイベントのハンドリング (Excelからのコピペ対応)
            input.addEventListener('paste', handlePaste);
            // 矢印キー移動などの便宜上、ID等を振っても良いですが省略
            
            td.appendChild(input);
            tr.appendChild(td);
        }
        tbody.appendChild(tr);
    }
});

// Excelからの貼り付けを処理する関数
function handlePaste(e) {
    e.preventDefault();
    // クリップボードのデータを取得
    const pasteData = (e.clipboardData || window.clipboardData).getData('text');
    
    // 行と列に分解 (Excelはタブ区切り・改行区切り)
    const rows = pasteData.split(/\r\n|\r|\n/).filter(row => row.length > 0);
    
    // ペーストを開始したセルの位置を特定
    const targetInput = e.target;
    const targetTd = targetInput.parentElement;
    const targetTr = targetTd.parentElement;
    
    const startRowIndex = Array.from(targetTr.parentElement.children).indexOf(targetTr);
    const startColIndex = Array.from(targetTr.children).indexOf(targetTd);
    
    const tableRows = document.querySelectorAll('#inputTable tbody tr');

    rows.forEach((rowData, rIdx) => {
        const cols = rowData.split('\t');
        const currentRowIndex = startRowIndex + rIdx;

        if (currentRowIndex < tableRows.length) {
            const cells = tableRows[currentRowIndex].querySelectorAll('input');
            
            cols.forEach((cellData, cIdx) => {
                const currentColIndex = startColIndex + cIdx;
                if (currentColIndex < cells.length) {
                    cells[currentColIndex].value = cellData;
                }
            });
        }
    });
}

// 変換処理
async function processData() {
    const fileInput = document.getElementById('uploadFile');
    const convertBtn = document.getElementById('convertBtn');
    const statusMsg = document.getElementById('statusMessage');
    const downloadLink = document.getElementById('downloadLink');

    // 1. ファイルチェック
    if (!fileInput.files[0]) {
        alert('Excelファイルを選択してください。');
        return;
    }

    // 2. 表データの取得
    const tableData = [];
    const rows = document.querySelectorAll('#inputTable tbody tr');
    rows.forEach(tr => {
        const rowValues = [];
        tr.querySelectorAll('input').forEach(input => {
            rowValues.push(input.value);
        });
        tableData.push(rowValues);
    });

    // 3. 送信データの作成
    const formData = new FormData();
    formData.append('file', fileInput.files[0]);
    formData.append('tableData', JSON.stringify(tableData));

    // UI更新
    convertBtn.disabled = true;
    statusMsg.textContent = "処理中...";
    downloadLink.style.display = 'none';

    try {
        // 4. APIへ送信
        const response = await fetch('/api/process', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            const errText = await response.text();
            throw new Error('サーバーエラー: ' + errText);
        }

        // 5. ファイルダウンロード処理
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        
        // ヘッダーからファイル名を取得しようと試みる（できなければデフォルト）
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = fileInput.files[0].name.replace('.xlsx', '_complete.xlsx');
        
        downloadLink.href = url;
        downloadLink.download = filename;
        downloadLink.textContent = `ダウンロード (${filename})`;
        downloadLink.style.display = 'block';
        downloadLink.style.backgroundColor = '#28a745'; // 成功色
        
        statusMsg.textContent = "変換完了！下のボタンからダウンロードしてください。";

    } catch (error) {
        console.error(error);
        statusMsg.textContent = "エラーが発生しました: " + error.message;
        statusMsg.style.color = "red";
    } finally {
        convertBtn.disabled = false;
    }
}
