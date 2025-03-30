# 中文文字校正工具

這是一個基於Python的中文文字校正工具，可以幫助用戶自動校正文字中的錯字。特別適用於離線環境，無需網路連接即可使用。

## 功能特點

- 用於文字校正的OpenCC整合
- 支援Word文檔拖放功能
- 可處理帶密碼保護的Word文檔
- 自訂詞彙保護表，防止特定詞彙被自動校正
- 900x600的用戶界面，含700x500的文字處理區和100x500的圖片顯示區

## 安裝說明

### 前提條件

- Windows作業系統
- Python 3.6+

### 安裝步驟

1. 安裝依賴庫：

   ```bash
   pip install -r requirements.txt
   ```

2. 確保tkinter可用（大多數Python安裝已包含）

### 離線安裝說明

如果您在完全離線的環境中工作，請按照以下步驟操作：

1. 在有網路連接的電腦上下載所有需要的Python庫：

   ```bash
   pip download -r requirements.txt -d ./offline_packages
   ```

2. 將整個專案資料夾（包括offline_packages目錄）複製到離線電腦

3. 在離線電腦上安裝：

   ```bash
   pip install --no-index --find-links=./offline_packages -r requirements.txt
   ```

## 使用方法

1. 執行主程式：

   ```bash
   python main.py
   ```

2. 使用方式：

   - 將Word文檔拖放到應用程式視窗中
   - 或使用選單列中的"開啟"選項
   - 如遇到密碼保護的文檔，系統會提示輸入密碼

3. 詞彙保護功能：

   - 使用選單列中的"管理保護詞彙"選項
   - 添加需要保護的詞彙（這些詞彙不會被自動校正）

## 注意事項

- 此程式依賴於OpenCC進行字元轉換
- 保護詞彙儲存在protected_words.json檔案中
- 離線環境下請確保所有依賴包已正確安裝
