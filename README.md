# 📝 智能會議記錄生成器 (Smart Minutes Generator)

基於 **Streamlit**、**GitHub Models (GPT-4o-mini)** 與 **動態範本映射技術 (Dynamic Template Mapping)** 開發的通用企業級 AI 會議記錄自動化工具。

🌐 **線上使用網址：** [https://smart-minutes-generator.streamlit.app](https://smart-minutes-generator.streamlit.app)

---

## ✨ 核心特色與商業價值

* **📄 直接匯出高質感 Word (.docx)：** 支援一鍵導出具備深藍商務標頭、精確欄寬與防跨頁斷裂格式的標準 Word 文件，下載即用，完美符合真實企業辦公習慣。
* **🏛️ 多元內建商務範本：** 內建 3 大常見企業會議格式（通用週會/例會、董事會/高層決議型、專案進度/檢討型），無須預先準備檔案即可即刻使用。
* **📋 動態結構映射 (Dynamic Structure Mapping)：** 自動閱讀 Word 範本並提煉「編號與議題標題」作為骨架，將新會議內容精準對號入座。
* **📊 雙模式內容輸入：**
  * **模式 A (PPT 簡報)：** 自動解析簡報標題、內文與備註欄 (Notes)。
  * **模式 B (文字草稿)：** 無 PPT 時，可直接於網頁輸入框按層級格式 (如 `1.1 ... 1.2 ...`) 貼上會議紀錄草稿。
* **📥 側邊欄基準範本下載：** 內建 3 大商務 Word 範本檔，方便使用者隨時下載空白範本進行線下編輯。
* **🛡️ 零數據留存資安架構 (Zero Data Retention)：** 採用純 Session 記憶體即時運算，關閉網頁後數據即刻徹底銷毀，完全符合企業商業機密保護要求。

---

## 🛡️ ISO 資訊安全與隱私合規聲明 (ISO Security & Privacy Compliance)

本系統專為重視商業機密與數據隱私的企業團隊設計，合規架構明確對齊國際標準：

* **🔒 ISO/IEC 27001 (資訊安全管理體系 - ISMS) 對齊：**
  * **零數據留存 (Zero Data Retention / ZDR)：** 系統採用 Session-Only 記憶體運算架構。使用者上傳之 `.docx` 範本、`.pptx` 簡報及生成之會議記錄**絕不儲存於任何伺服器硬碟、資料庫或第三方儲存庫**。
  * **即時銷毀 (Instant Destruction)：** 當使用者關閉或重新整理瀏覽器分頁，所有運算記憶體與對話 Session 即刻被徹底銷毀。
* **🔒 ISO/IEC 27701 (隱私資訊管理體系 - PIMS) 對齊：**
  * **數據最小化 (Data Minimization)：** 系統僅在記憶體中提煉必要的議題結構與會議文字進行對齊，不進行任何無關數據的分析或留存。
  * **無模型訓練 (No AI Model Training)：** 透過企業級 API 進行運算，傳輸數據嚴禁用於任何大語言模型 (LLM) 的再訓練。
* **🌐 傳輸安全 (TLS/SSL Encryption)：** 全程採用 HTTPS / TLS 1.3 國際標準加密傳輸，確保資料在傳送過程免受攔截。

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

## 🚀 企業級功能演進藍圖 (Enterprise Roadmap)

為了進一步深化 ISO 27001/27701 合規實踐與企業級營運體驗，本專案規劃以下三個階段的迭代計畫：

* **Phase 1: 資安與數據主權強化 (Security & Data Sovereignty)**
  * [ ] **PII 本地動態遮蔽 (Local PII Redaction):** 於送出 API 前進行本地端 HKID、電話、薪資等敏感個資正則遮蔽，實現「傳輸前假名化 (Pseudonymization)」。
  * [ ] **自備 API Key (BYOK - Bring Your Own Key):** 提供企業自選 OpenAI / Azure OpenAI 私有端點，讓資料流精確符合企業內部的 Azure VPC 安全規範。
* **Phase 2: 產品體驗與範本忠實度升級 (User Experience & Fidelity)**
  * [ ] **原檔範本忠實留存 (Preserve Original Template):** 導入「模板克隆 + 結構注入」模式，直接回填文字至用戶上傳的 Word 檔，100% 完整保留企業 Logo、頁首頁尾與專屬色系。
  * [ ] **區塊化對話微調 (Section-based Chat Refinement):** 允許使用者點選單一章節（如 3.1 條文）進行局部 AI 對話修訂，節省 Token 消耗並實現即時完稿。
* **Phase 3: 治理與稽核軌跡 (AI Governance & Audit Trail)**
  * [ ] **結構化稽核日誌 (Compliance Audit Trail):** 記憶體內產生 Session-Only 的操作與 PII 遮蔽摘要日誌，支援導出供企業資安稽核員 (CISO/Auditor) 備查。

---

## 🛠️ 技術架構 (Tech Stack)

* **前端與 UI：** Streamlit
* **大語言模型 API：** GitHub Models / Azure OpenAI (`gpt-4o-mini`)
* **文件解析與轉換：** `python-docx` (Word 解析與高質感 Word 生成) + `python-pptx` (PPT 簡報解析)
* **樣式優化：** Custom CSS + OpenXML Table Styling (控制 Word 表格欄寬、背景色與對齊)

---

## 💡 技術支援與聯絡方式 (Support)

本專案由 [Jacky Law](https://jackylawck.github.io/jackylawck/) 開發與維護。如有任何系統建置、商業部署或 AI 治理諮詢問題，歡迎透過連結聯繫。

---

## 📄 授權說明 (License)

本專案採用 **MIT License** 開源授權。系統僅供行政與會議記錄整理輔助使用，所有資料均於本地瀏覽器 Session 記憶體內處理，不存檔至任何外部雲端資料庫。
