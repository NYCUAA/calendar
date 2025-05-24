// 設定常數
const SPREADSHEET_ID = '1tBqno9dbUYVkckSavCZePIDMTXAYb1ljUBSMgLLOXYE';
const SHEET_NAME = '行事曆總表';

// 用於API的驗證金鑰（應該設置為較長、隨機的字串）
const API_KEY = 'NYCUAA';

// 欄位常數定義（標題列名稱及其對應的索引）
const COLUMNS = {
  開始時間: '開始時間',
  結束時間: '結束時間',
  活動名稱: '活動名稱',
  活動地點: '活動地點',
  活動地點分類: '活動地點分類',
  分類: '分類',
  活動單位: '活動單位',
  活動相關連結: '活動相關連結',
  活動內容: '活動內容',
  活動性質: '活動性質',
  原始資料: '原始資料'
};

// GET請求處理 - 用於提供API功能
function doGet(e) {
  // 檢查API金鑰是否有效
  const providedKey = e.parameter.key;
  
  // 如果未提供金鑰，或金鑰不正確，則返回錯誤訊息
  if (!providedKey || providedKey !== API_KEY) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: '未授權的訪問'
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    // 獲取試算表所有資料
    const data = getAllSheetData();
    
    // 返回JSON格式的資料
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: data
    }))
    .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    // 發生錯誤時返回錯誤訊息
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: `獲取資料時發生錯誤：${error.message}`
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
}

// 獲取試算表中的所有資料
function getAllSheetData() {
  // 獲取試算表和工作表
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`找不到工作表：${SHEET_NAME}`);
  }
  
  // 獲取所有資料（包括標題行）
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // 如果試算表為空，則返回空數組
  if (values.length === 0) {
    return [];
  }
  
  // 獲取標題行
  const headers = values[0];
  
  // 轉換資料為JSON格式（從第2行開始，跳過標題行）
  const jsonData = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowData = {};
    
    // 將每一列映射到對應的標題
    for (let j = 0; j < headers.length; j++) {
      // 只包含非空值
      if (row[j] !== '') {
        rowData[headers[j]] = row[j];
      }
    }
    
    jsonData.push(rowData);
  }
  
  return jsonData;
}

// Web應用程式入口
function doPost(e) {
  // 解析LINE傳入的事件
  const event = JSON.parse(e.postData.contents).events[0];
  
  // 如果不是訊息事件，則不處理
  if (event.type !== 'message' || event.message.type !== 'text') {
    return ContentService.createTextOutput(JSON.stringify({'status': 'not a text message'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    // 獲取使用者傳送的訊息文字
    const userMessage = event.message.text;
    const userId = event.source.userId;
    
    // 檢查資料是否重複
    const isDuplicate = checkDuplicateData(userMessage);
    
    // 如果資料重複，直接回覆用戶
    if (isDuplicate) {
      replyToUser(event.replyToken, "資料庫已有此筆資料，無需重複新增。");
      
      return ContentService.createTextOutput(JSON.stringify({'status': 'duplicate data'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 使用GPT API處理文字並提取結構化數據
    const extractedData = extractDataWithGPT(userMessage);
    
    // 保存原始訊息到提取的數據中
    extractedData.原始資料 = userMessage;
    
    // 將數據寫入Google Sheets
    const addResult = addToCalendarSheet(extractedData);
    
    // 準備回覆訊息
    const replyMessage = `新增成功！\n\n活動名稱：${extractedData.活動名稱}\n開始時間：${extractedData.開始時間}\n結束時間：${extractedData.結束時間}\n活動地點：${extractedData.活動地點}\n活動性質：${extractedData.活動性質 || '其他'}`;
    
    // 回覆用戶
    replyToUser(event.replyToken, replyMessage);
    
    return ContentService.createTextOutput(JSON.stringify({'status': 'success'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    // 發生錯誤時回覆錯誤訊息
    replyToUser(event.replyToken, `處理訊息時發生錯誤：${error.message}`);
    
    return ContentService.createTextOutput(JSON.stringify({'status': 'error', 'message': error.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 使用GPT API提取結構化數據
function extractDataWithGPT(text) {
  // 動態取得當前年份
  const currentYear = new Date().getFullYear();
  
  const requestBody = {
    'model': 'gpt-4o',
    'messages': [
      {
        'role': 'system',
        'content': `你是一位專門從非結構化文字中提取活動資訊的專家。請從使用者提供的文字中提取活動相關資訊並依照指定格式輸出。其中，如果使用者沒有給明確年份，年份時間不確定的話一律用${currentYear}年份`
      },
      {
        'role': 'user',
        'content': text
      }
    ],
    'functions': [
      {
        'name': 'extract_calendar_event',
        'description': '從文字中提取行事曆活動資訊，注意，其他無法歸類的活動的詳細內容描述放「活動內容」中，完全用本文文字，勿自行詮釋說法',
        'parameters': {
          'type': 'object',
          'properties': {
            '開始時間': { 'type': 'string', 'description': `活動開始時間，盡量使用YYYY/MM/DD HH:MM格式，年份時間不確定的話一律用${currentYear}年份` },
            '結束時間': { 'type': 'string', 'description': '活動結束時間，盡量使用YYYY/MM/DD HH:MM格式' },
            '活動名稱': { 'type': 'string', 'description': '活動的名稱' },
            '活動地點': { 'type': 'string', 'description': '活動實際舉辦的確切地點' },
            '活動地點分類': { 'type': 'string', 'description': '活動地點的分類，呈現台灣的縣市名稱即可' },
            '分類': { 'type': 'string', 'description': '活動的分類，僅分為"校友會", "其他活動"。' },
            '活動單位': { 'type': 'string', 'description': '舉辦活動的單位' },
            '活動相關連結': { 'type': 'string', 'description': '活動的相關網頁連結' },
            '活動性質': { 'type': 'string', 'description': '活動的性質，僅能是 "遊玩" "演講" "課程" "餐會" "其他"，根據活動內容選擇最適合的一項' },
            '活動內容': { 'type': 'string', 'description': '其他無法歸類的活動的詳細內容描述放這邊，請完全用本文文字，勿自行詮釋' }
          },
          'required': ['開始時間', '活動名稱']
        }
      }
    ],
    'function_call': { 'name': 'extract_calendar_event' }
  };

  const options = {
    'method': 'post',
    'headers': {
      'Authorization': `Bearer ${OPENAI_API_KEY}`,
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify(requestBody),
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
  const responseData = JSON.parse(response.getContentText());
  
  if (responseData.error) {
    throw new Error(`GPT API錯誤: ${responseData.error.message}`);
  }
  
  // 從回應中取出函數調用參數
  const functionCall = responseData.choices[0].message.function_call;
  if (!functionCall || functionCall.name !== 'extract_calendar_event') {
    throw new Error('GPT未能正確提取資訊');
  }
  
  return JSON.parse(functionCall.arguments);
}

// 將數據添加到Google Sheets
function addToCalendarSheet(data) {
  // 獲取試算表和工作表
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`找不到工作表：${SHEET_NAME}`);
  }
  
  // 獲取標題行，用於確定每個欄位的位置
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnIndices = getColumnIndices(headerRow);
  
  // 如果缺少必要的欄位，建立它們
  if (Object.keys(columnIndices).length === 0) {
    // 創建標題行
    const headers = [
      COLUMNS.開始時間,
      COLUMNS.結束時間,
      COLUMNS.活動名稱,
      COLUMNS.活動地點,
      COLUMNS.活動地點分類,
      COLUMNS.分類,
      COLUMNS.活動單位,
      COLUMNS.活動相關連結,
      COLUMNS.活動內容,
      COLUMNS.活動性質,
      COLUMNS.原始資料
    ];
    sheet.appendRow(headers);
    
    // 重新獲取標題行
    const newHeaderRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Object.keys(COLUMNS).forEach((key, index) => {
      columnIndices[key] = index + 1;
    });
  }
  
  // 保存活動內容，確保保留原始格式
  let activityContent = data.活動內容 || '';
  
  // 保存原始資料
  let originalData = data.原始資料 || '';
  
  // 準備要寫入的數據
  const rowData = Array(sheet.getLastColumn()).fill(''); // 初始化一個全為空字符串的數組
  
  // 根據欄位索引填入數據
  if (columnIndices.開始時間) rowData[columnIndices.開始時間 - 1] = data.開始時間 || '';
  if (columnIndices.結束時間) rowData[columnIndices.結束時間 - 1] = data.結束時間 || '';
  if (columnIndices.活動名稱) rowData[columnIndices.活動名稱 - 1] = data.活動名稱 || '';
  if (columnIndices.活動地點) rowData[columnIndices.活動地點 - 1] = data.活動地點 || '';
  if (columnIndices.活動地點分類) rowData[columnIndices.活動地點分類 - 1] = data.活動地點分類 || '';
  if (columnIndices.分類) rowData[columnIndices.分類 - 1] = data.分類 || '';
  if (columnIndices.活動單位) rowData[columnIndices.活動單位 - 1] = data.活動單位 || '';
  if (columnIndices.活動相關連結) rowData[columnIndices.活動相關連結 - 1] = data.活動相關連結 || '';
  if (columnIndices.活動內容) rowData[columnIndices.活動內容 - 1] = activityContent;
  if (columnIndices.活動性質) rowData[columnIndices.活動性質 - 1] = data.活動性質 || '其他';
  if (columnIndices.原始資料) rowData[columnIndices.原始資料 - 1] = originalData;
  
  // 在最後一行添加數據
  const newRow = sheet.appendRow(rowData);
  
  // 獲取新增行的索引
  const lastRow = sheet.getLastRow();
  
  // 設定活動內容的格式為自動換行
  if (activityContent && columnIndices.活動內容) {
    // 獲取活動內容儲存格
    const contentCell = sheet.getRange(lastRow, columnIndices.活動內容);
    
    // 設定儲存格格式為自動換行
    contentCell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
  
  // 設定原始資料的格式為自動換行
  if (originalData && columnIndices.原始資料) {
    // 獲取原始資料儲存格
    const originalDataCell = sheet.getRange(lastRow, columnIndices.原始資料);
    
    // 設定儲存格格式為自動換行
    originalDataCell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
  
  return {
    success: true,
    message: '已成功添加數據到試算表'
  };
}

// 根據標題行獲取每個欄位的索引
function getColumnIndices(headerRow) {
  const indices = {};
  
  Object.values(COLUMNS).forEach(columnName => {
    const index = headerRow.indexOf(columnName);
    if (index !== -1) {
      const key = Object.keys(COLUMNS).find(key => COLUMNS[key] === columnName);
      indices[key] = index + 1; // 轉為1-based索引
    }
  });
  
  return indices;
}

// 回覆LINE用戶訊息
function replyToUser(replyToken, message) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  
  const payload = {
    'replyToken': replyToken,
    'messages': [
      {
        'type': 'text',
        'text': message
      }
    ]
  };
  
  const options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  return response;
}

// 用於測試GPT提取功能
function testGPTExtraction() {
  const testText = `📢 陽明交大校友快來集合！
🌊🚍 高雄愛河半日遊，6/28（六） 校友限定！ 🚢✨ 

😱剩下15個名額😱

這次我們將搭乘 雙層觀光巴士，欣賞高雄港灣美景，傍晚散步愛河，最後搭乘 愛之船 夜遊，感受高雄的迷人夜色！還有校友們的歡樂交流時光，千萬別錯過！

📅 集合時間：6/28（六）16:10
📍 集合地點：愛之船國賓站
💰 活動費用：
校友 300 元 / 非校友 500元（巴士+餐費）+ 100-150 元（愛之船）
🎟️ 報名方式：點進連結報名，名額有限！

👉 立即報名 🔗 https://forms.gle/Rz6ftHX2LVmPWYWj6
📢 歡迎校友們攜家帶眷，一起來場高雄校友情之旅！ 😍💙

📝 行程安排
16:30 - 17:30｜搭乘雙層巴士西子灣線，聆聽專業導覽，欣賞高雄港灣風光
17:30 - 18:30｜愛河散步 & 校友交流，沿途拍照、欣賞夕陽
18:30 - 20:00｜預計至【東京酒場】享用日式料理，體驗宮崎駿風格湯屋氛圍
20:00 - 20:30｜搭乘 愛之船（自費）夜遊愛河，感受高雄迷人夜景
20:30 - 21:00｜自由解散，結束美好校友之夜`;
  
  const result = extractDataWithGPT(testText);
  Logger.log(JSON.stringify(result, null, 2));
}

// 用於部署Web應用程式的設定
function getScriptURL() {
  return ScriptApp.getService().getUrl();
}

// 測試GPT提取資料後填入試算表
function testExtractAndAddToSheet() {
  // 測試文字
  const testText = `📢 陽明交大校友快來集合！
🌊🚍 高雄愛河半日遊，6/28（六） 校友限定！ 🚢✨ 

😱剩下15個名額😱

這次我們將搭乘 雙層觀光巴士，欣賞高雄港灣美景，傍晚散步愛河，最後搭乘 愛之船 夜遊，感受高雄的迷人夜色！還有校友們的歡樂交流時光，千萬別錯過！

📅 集合時間：6/28（六）16:10
📍 集合地點：愛之船國賓站
💰 活動費用：
校友 300 元 / 非校友 500元（巴士+餐費）+ 100-150 元（愛之船）
🎟️ 報名方式：點進連結報名，名額有限！

👉 立即報名 🔗 https://forms.gle/Rz6ftHX2LVmPWYWj6
📢 歡迎校友們攜家帶眷，一起來場高雄校友情之旅！ 😍💙

📝 行程安排
16:30 - 17:30｜搭乘雙層巴士西子灣線，聆聽專業導覽，欣賞高雄港灣風光
17:30 - 18:30｜愛河散步 & 校友交流，沿途拍照、欣賞夕陽
18:30 - 20:00｜預計至【東京酒場】享用日式料理，體驗宮崎駿風格湯屋氛圍
20:00 - 20:30｜搭乘 愛之船（自費）夜遊愛河，感受高雄迷人夜景
20:30 - 21:00｜自由解散，結束美好校友之夜`;
  
  try {
    // 步驟0：先檢查資料是否重複
    Logger.log('檢查資料是否重複...');
    const isDuplicate = checkDuplicateData(testText);
    
    if (isDuplicate) {
      Logger.log('發現重複資料，測試終止');
      return {
        success: false,
        message: '測試終止：資料庫已有此筆資料，無需重複新增',
        isDuplicate: true
      };
    }
    
    Logger.log('資料未重複，繼續進行測試...');
    
    // 步驟1：使用GPT提取數據
    Logger.log('開始提取數據...');
    const extractedData = extractDataWithGPT(testText);
    Logger.log('數據提取成功：');
    Logger.log(JSON.stringify(extractedData, null, 2));
    
    // 添加原始資料
    extractedData.原始資料 = testText;
    
    // 步驟2：將數據添加到試算表
    Logger.log('正在將數據添加到試算表...');
    const addResult = addToCalendarSheet(extractedData);
    
    // 步驟3：檢查結果並顯示成功信息
    if (addResult.success) {
      Logger.log('測試完成：數據已成功添加到試算表！');
      Logger.log(`活動名稱：${extractedData.活動名稱}`);
      Logger.log(`開始時間：${extractedData.開始時間}`);
      Logger.log(`結束時間：${extractedData.結束時間}`);
      Logger.log(`活動地點：${extractedData.活動地點}`);
      Logger.log(`活動性質：${extractedData.活動性質 || '其他'}`);
      return {
        success: true,
        message: '測試成功：GPT提取數據並成功添加到試算表',
        data: extractedData
      };
    } else {
      Logger.log('測試失敗：無法添加數據到試算表');
      return {
        success: false,
        message: '測試失敗：無法添加數據到試算表'
      };
    }
  } catch (error) {
    Logger.log(`測試過程中發生錯誤：${error.message}`);
    return {
      success: false,
      message: `測試失敗：${error.message}`
    };
  }
}

// 檢查是否有重複資料
function checkDuplicateData(userMessage) {
  // 獲取試算表和工作表
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`找不到工作表：${SHEET_NAME}`);
  }
  
  // 獲取標題行，用於確定原始資料欄位的位置
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columnIndices = getColumnIndices(headerRow);
  
  // 如果找不到原始資料欄位，無法檢查重複
  if (!columnIndices.原始資料) {
    return false;
  }
  
  // 獲取原始資料欄的索引
  const originalDataColumnIndex = columnIndices.原始資料;
  
  // 獲取原始資料欄的所有數據
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    // 表格中只有標題行或空表格，無需檢查
    return false;
  }
  
  // 獲取原始資料列全部數據（從第2行開始，跳過標題行）
  const originalDataRange = sheet.getRange(2, originalDataColumnIndex, lastRow - 1, 1);
  const originalDataValues = originalDataRange.getValues();
  
  // 檢查是否有相同的原始資料
  for (let i = 0; i < originalDataValues.length; i++) {
    // 比較原始資料是否相同（去除可能的首尾空格）
    if (originalDataValues[i][0].toString().trim() === userMessage.trim()) {
      return true; // 找到重複資料
    }
  }
  
  // 未找到重複資料
  return false;
}

// 部署Web應用程式並獲取URL
function getDeploymentUrls() {
  const scriptUrl = ScriptApp.getService().getUrl();
  Logger.log('目前的部署URL: ' + scriptUrl);
  Logger.log('---');
  Logger.log('部署說明:');
  Logger.log('1. 點選 [部署] > [新增部署]');
  Logger.log('2. 選擇類型: "網頁應用程式"');
  Logger.log('3. 設定說明:');
  Logger.log('   - 執行身分: "我的身分"');
  Logger.log('   - 存取權限: "任何人"');
  Logger.log('4. 點擊 [部署]');
  Logger.log('5. 複製生成的URL');
  Logger.log('---');
  Logger.log('LINE機器人Webhook URL: 使用上方相同的URL');
  Logger.log('API資料存取URL: 上方URL + "?key=your_secure_api_key_here"');
  
  return {
    currentUrl: scriptUrl,
    lineWebhookUrl: scriptUrl,
    apiUrl: scriptUrl + '?key=' + API_KEY
  };
}

// 測試doGet函數
function testDoGet() {
  Logger.log('===== 開始測試 doGet 函數 =====');
  
  // 測試1：使用正確的API金鑰
  Logger.log('測試1: 使用正確的API金鑰');
  const correctKeyParam = {
    parameter: {
      key: API_KEY
    }
  };
  
  try {
    const correctKeyResponse = doGet(correctKeyParam);
    const correctKeyContent = JSON.parse(correctKeyResponse.getContent());
    
    Logger.log('回應狀態: ' + (correctKeyContent.success ? '成功' : '失敗'));
    
    if (correctKeyContent.success) {
      Logger.log('獲取到的資料筆數: ' + correctKeyContent.data.length);
      
      // 顯示前5筆資料的基本資訊（如果有）
      if (correctKeyContent.data.length > 0) {
        Logger.log('資料預覽（前5筆）:');
        const previewCount = Math.min(5, correctKeyContent.data.length);
        
        for (let i = 0; i < previewCount; i++) {
          const item = correctKeyContent.data[i];
          Logger.log(`[${i+1}] 活動名稱: ${item.活動名稱 || '無'}, 開始時間: ${item.開始時間 || '無'}`);
        }
      }
    } else {
      Logger.log('錯誤訊息: ' + correctKeyContent.message);
    }
  } catch (error) {
    Logger.log('測試1執行錯誤: ' + error.message);
  }
  
  // 測試2：使用錯誤的API金鑰
  Logger.log('\n測試2: 使用錯誤的API金鑰');
  const wrongKeyParam = {
    parameter: {
      key: 'wrong_key'
    }
  };
  
  try {
    const wrongKeyResponse = doGet(wrongKeyParam);
    const wrongKeyContent = JSON.parse(wrongKeyResponse.getContent());
    
    Logger.log('回應狀態: ' + (wrongKeyContent.success ? '成功' : '失敗'));
    if (!wrongKeyContent.success) {
      Logger.log('錯誤訊息: ' + wrongKeyContent.message);
    }
  } catch (error) {
    Logger.log('測試2執行錯誤: ' + error.message);
  }
  
  // 測試3：不提供API金鑰
  Logger.log('\n測試3: 不提供API金鑰');
  const noKeyParam = {
    parameter: {}
  };
  
  try {
    const noKeyResponse = doGet(noKeyParam);
    const noKeyContent = JSON.parse(noKeyResponse.getContent());
    
    Logger.log('回應狀態: ' + (noKeyContent.success ? '成功' : '失敗'));
    if (!noKeyContent.success) {
      Logger.log('錯誤訊息: ' + noKeyContent.message);
    }
  } catch (error) {
    Logger.log('測試3執行錯誤: ' + error.message);
  }
  
  Logger.log('\n===== doGet 函數測試完成 =====');
  
  return '測試已完成，請查看執行記錄(Logs)以獲取詳細結果';
}

// 測試API回應格式與資料結構
function testAPIFormat() {
  Logger.log('===== 開始測試 API 回應格式 =====');
  
  // 使用正確的API金鑰獲取資料
  const params = {
    parameter: {
      key: API_KEY
    }
  };
  
  try {
    const response = doGet(params);
    const content = JSON.parse(response.getContent());
    
    if (!content.success) {
      Logger.log('API回應錯誤: ' + content.message);
      return;
    }
    
    // 檢查資料結構
    Logger.log('資料類型: ' + typeof content.data);
    
    if (!Array.isArray(content.data)) {
      Logger.log('錯誤: 資料不是陣列格式');
      return;
    }
    
    Logger.log('資料筆數: ' + content.data.length);
    
    // 如果有資料，檢查第一筆資料的欄位
    if (content.data.length > 0) {
      const firstItem = content.data[0];
      Logger.log('第一筆資料欄位:');
      
      Object.keys(firstItem).forEach(key => {
        Logger.log(`- ${key}: ${typeof firstItem[key]}`);
      });
      
      // 檢查欄位是否符合期望
      const expectedColumns = [
        '開始時間', '結束時間', '活動名稱', '活動地點', 
        '活動地點分類', '分類', '活動單位', '活動相關連結', 
        '活動內容', '活動性質'
      ];
      
      Logger.log('\n檢查必要欄位:');
      expectedColumns.forEach(column => {
        if (column in firstItem) {
          Logger.log(`✓ ${column} - 存在`);
        } else {
          Logger.log(`✗ ${column} - 缺失`);
        }
      });
    }
    
    // 檢查MIME類型
    Logger.log('\nMIME類型: ' + response.getMimeType());
    
  } catch (error) {
    Logger.log('測試過程中發生錯誤: ' + error.message);
  }
  
  Logger.log('\n===== API 格式測試完成 =====');
  
  return '測試已完成，請查看執行記錄(Logs)以獲取詳細結果';
} 