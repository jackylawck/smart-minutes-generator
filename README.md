# 📝 智能會議記錄生成器 (Smart Minutes Generator)

基於 **Streamlit**、**GitHub Models (GPT-4o-mini)** 與 **動態範本映射技術 (Dynamic Template Mapping)** 開發的通用企業級會議記錄自動化工具。

🌐 **線上使用網址：** [https://smart-minutes-generator.streamlit.app](https://smart-minutes-generator.streamlit.app)

---

## ✨ 核心特色與功能

* **📄 直接匯出高質感 Word (.docx)：** 支援一鍵導出具備深藍商務標頭、精確欄寬與防跨頁斷裂格式的標準 Word 文件，下載即用，符合真實企業辦公習慣。
* **🏛️ 多元內建商務範本：** 內建 3 大常見企業會議格式（通用週會/例會、董事會/高層決議型、專案進度/檢討型），無需預先準備檔案即可即刻使用。
* **📋 動態結構映射 (Dynamic Structure Mapping)：** 自動閱讀 Word 範本並提煉「編號與議題標題」作為骨架，將新會議內容精準對號入座。
* **📊 雙模式內容輸入：**
  * **模式 A (PPT 簡報)：** 自動解析簡報標題、內文與備註欄 (Notes)。
  * **模式 B (文字草稿)：** 無 PPT 時，可直接於網頁輸入框按層級格式 (如 `1.1 ... 1.2 ...`) 貼上會議紀錄草稿。
* **📥 側邊欄基準範本下載：** 內建 3 大商務 Word 範本檔，方便使用者隨時下載空白範本進行線下編輯。
* **🔒 Privacy-First / 遵循 ISO 資安標準：** 採用 **Session-Only 記憶體即時運算**，使用者上傳的檔案與生成結果於關閉網頁後即刻銷毀，遵循 ISO 27001 / ISO 27701 數據最小化與 Zero Data Retention (ZDR) 原則。

---

## 📖 使用指南 (User Guide)

1. **選擇結構範本：** 
   * **方式 A：** 選擇「內建標準範本」，並下拉挑選週會、董事會或專案檢討型格式。
   * **方式 B：** 選擇「自行上傳」，上傳貴單位過往使用的會議記錄 (.docx) 作為骨架。
2. **提供本次會議內容 (二選一)：**
   * **選擇 A：** 上傳本次會議簡報 (`.pptx`)。
   * **選擇 B：** 直接於文字框貼上會議草稿 (按 `1. 議題` -> `1.1 子項目` 層級輸入)。
3. **生成與下載：** 點擊「🚀 即刻依範本結構生成會議記錄」，預覽結果無誤後，即可點擊「📄 即刻下載標準 Word 格式會議記錄 (.docx)」。

---

## 🛡️ 資安與隱私聲明 (Security & Privacy)

* **零資料留存 (Zero Data Retention)：** 本系統採用純 Session 記憶體處理，使用者上傳之 `.docx` 及 `.pptx` 檔案不會儲存於任何伺服器硬碟、資料庫或 GitHub 儲存庫。
* **傳輸安全 (TLS/SSL)：** 網頁傳輸與 API 連線全程採用 HTTPS / TLS 加密。
* **Secrets 管理：** API 密鑰儲存於受保護的環境變數中，確保憑證不外洩。
* **架構合規：** 系統設計遵循 ISO/IEC 27001 (資訊安全管理) 及 ISO/IEC 27701 (隱私資訊管理) 之「數據最小化」與「機密性保護」原則。

---

## 🛠️ 技術架構 (Tech Stack)

* **前端與 UI：** Streamlit
* **大語言模型 API：** GitHub Models / Azure OpenAI (`gpt-4o-mini`)
* **文件解析與轉換：** `python-docx` (Word 解析與高質感 Word 生成) + `python-pptx` (PPT 簡報解析)
* **樣式優化：** Custom CSS + OpenXML Table Styling (控制 Word 表格欄寬、背景色與對齊)

---

## 📄 授權說明 (License)

本專案採用 **MIT License** 開源授權。系統僅供行政與會議記錄整理輔助使用，所有資料均於本地瀏覽器 Session 記憶體內處理，不存檔至任何外部雲端資料庫。
