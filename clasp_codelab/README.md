# Google Apps Script Clasp 專案開發

本專案使用 Google 的 `clasp` 工具進行本地端開發。

## ⚠️ 重要待辦事項
- [ ] **更新 Script ID**：請記得將 `.clasp.json` 檔案中的 `"scriptId"` 欄位，替換為您在 Google Apps Script 雲端專案的真實 Script ID。

---

## 常用 clasp 指令說明

在終端機中，請確保路徑位於 `clasp_codelab` 資料夾下，並可執行以下指令：

### 1. 登入 Google 帳號
首次使用 clasp 或是切換帳號時，需要先登入：
```bash
clasp login
```

### 2. 將本地程式碼上傳到雲端 (Push)
將本地的 `.gs`、`.html` 與 `appsscript.json` 檔案上傳並覆蓋雲端專案：
```bash
clasp push
```

### 3. 將雲端程式碼下載到本地 (Pull)
將雲端專案最新的程式碼下載並覆蓋本地檔案：
```bash
clasp pull
```

### 4. 開啟雲端開發編輯器 (Open)
直接在瀏覽器中開啟對應的 Google Apps Script 線上編輯器：
```bash
clasp open
```
