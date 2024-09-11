const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const bodyParser = require('body-parser');  // 用來解析 POST 請求的數據

const app = express();
const port = 3000;

// 設置靜態文件夾
app.use(express.static('public'));

// 解析 JSON 及 URL 編碼數據
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// 讀取 Excel 檔案並根據 ID 搜尋資料
app.post('/search', (req, res) => {
    // 使用絕對路徑指定 Excel 檔案的位置
    const filePath = path.join('C:', 'Users', '123', 'Desktop', 'web test', 'web_test.xlsx');
    
    // 讀取 Excel 檔案
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);

    // 輸出 Excel 讀取的資料來檢查
    console.log('Excel 讀取的資料:', data);

    // 從前端接收 ID 並去除空格
    const id = req.body.id.trim();
    
    // 根據 ID 搜尋相應的資料行，並去除 Excel 資料中的多餘空格
    const result = data.find(row => row.ID && row.ID.trim() === id);

    // 檢查尋找到的資料
    console.log('搜尋的 ID:', id);
    console.log('搜尋結果:', result);

    if (result) {
        // 準備返回的結果
        let response = [];

        // 檢查 B 欄和 C 欄是否有 "V"，並返回相應的欄位名稱
        if (result['證券'] === 'V') {
            response.push('證券');
        }
        if (result['期貨'] === 'V') {
            response.push('期貨');
        }

        // 返回結果
        res.json({ success: true, data: response });
    } else {
        // 如果沒有找到對應的 ID，返回錯誤信息
        res.json({ success: false, message: '無法找到對應資料' });
    }
});

// 啟動伺服器
app.listen(port, () => {
    console.log(`伺服器在 http://localhost:${port} 運行`);
});
