# LINE Bot 與 GPT 行事曆助手

這是一個結合 LINE Bot、OpenAI GPT 與 Google Sheets 的應用程式，能夠自動從使用者發送的文字中提取活動資訊，並將其添加到 Google 試算表中。

## 功能

- 從使用者傳送的文字中自動提取活動相關資訊
- 使用 OpenAI GPT-4o 處理自然語言並結構化資料
- 將提取的資訊自動填入 Google 試算表
- 回覆使用者添加成功的訊息與活動摘要

## 設置步驟

### 1. 建立 Google 試算表

1. 前往 [Google Sheets](https://sheets.google.com/) 並建立新的試算表
2. 將第一個工作表命名為「行事曆總表」
3. 在第一行設置以下欄位標題：
   - A: 開始時間
   - B: 結束時間
   - C: 活動名稱
   - D: 活動地點
   - E: 活動地點分類
   - F: 分類
   - G: 活動單位
   - H: 活動相關連結
   - I: 活動內容
4. 複製試算表的 ID（從 URL 中獲取，格式為：`https://docs.google.com/spreadsheets/d/{試算表ID}/edit`）

### 2. 設置 Google Apps Script

1. 在試算表中點選「擴充功能」→「Apps Script」
2. 刪除編輯器中的預設程式碼，並貼上此專案中的 `Code.gs` 內容
3. 更新以下常數：
   - `OPENAI_API_KEY`：從 [OpenAI API Keys](https://platform.openai.com/account/api-keys) 獲取
   - `SPREADSHEET_ID`：您剛才建立的試算表 ID
   - `SHEET_NAME`：確認為「行事曆總表」

### 3. 設置 LINE Bot

1. 前往 [LINE Developers Console](https://developers.line.biz/console/)
2. 建立新的 Provider（如果尚未有）
3. 創建一個新的 Messaging API Channel
4. 獲取 Channel Secret 和 Channel Access Token
5. 更新 `Code.gs` 中的 `LINE_CHANNEL_SECRET` 和 `LINE_CHANNEL_ACCESS_TOKEN`

### 4. 部署 Web 應用程式

1. 在 Google Apps Script 編輯器中，點選「部署」→「新增部署」
2. 選擇部署類型為「網頁應用程式」
3. 設置：
   - 執行身分：「我自己」
   - 誰可以存取：「所有人」
4. 點選「部署」，並複製生成的 Web 應用程式 URL

### 5. 設定 LINE Bot Webhook

1. 返回 LINE Developers Console
2. 在您的 Bot 設定頁面中，找到 Webhook URL 設定
3. 將剛才獲得的 Web 應用程式 URL 貼上
4. 啟用 Webhook
5. 關閉「自動回覆訊息」功能

## 使用方法

1. 在 LINE 中加入您的 Bot 為好友
2. 向 Bot 發送包含活動資訊的文字，例如：
   ```
   我們將在5月20日早上9點到下午3點在國際會議廳舉辦「2023科技創新論壇」，
   由資訊工程學系主辦，邀請了多位業界專家分享最新的AI技術發展。
   活動內容包括專題演講、Panel討論以及實際應用展示。
   詳情可參考活動網站：https://example.com/tech-forum
   ```
3. Bot 將自動處理訊息，提取活動資訊，將其添加到試算表，並回覆成功訊息

## 測試

您可以在 Google Apps Script 編輯器中執行 `testGPTExtraction` 函數來測試 GPT 提取功能，無需通過 LINE 介面。

## 注意事項

- OpenAI API 是付費服務，請確保您的帳戶有足夠的額度
- 確保 Google Apps Script 和試算表有足夠的權限
- 此應用程式每天有配額限制，詳情請參考 [Google Apps Script 配額](https://developers.google.com/apps-script/guides/services/quotas)

## 故障排除

如果遇到問題，請檢查：

1. Google Apps Script 執行記錄中的錯誤訊息
2. LINE Developers Console 中的 Webhook 回應狀態
3. 確認所有 API 金鑰和令牌是否正確
4. 檢查試算表權限是否正確設置 