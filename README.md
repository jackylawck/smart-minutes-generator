# 📝 智能會議記錄生成器 (Smart Minutes Generator)

基於 **Streamlit**、**GitHub Models / OpenAI (GPT-4o-mini)** 與 **動態範本映射技術 (Dynamic Template Mapping)** 開發的通用企業級 AI 會議記錄自動化工具。

🌐 **線上使用網址：** [https://smart-minutes-generator.streamlit.app](https://smart-minutes-generator.streamlit.app)

---

## ✨ 核心特色與商業價值 (Core Capabilities)

* **📄 高質感 Word (.docx) 原檔匯出：** 支援一鍵導出具備深藍商務標頭、精確欄寬與 OpenXML 防跨頁斷裂格式的標準 Word 文件，下載即用，符合董事會與高層呈閱標準。
* **🏛️ 原檔範本結構留存 (`fill_user_template_from_json`)：** 自動提取上傳範本的議題骨架與表格結構，並將內容直接回填至原本的 Word 檔中，100% 完整保留企業專屬 Logo、頁首頁尾與排版樣式。
* **🌐 多語言雙語支援 (Bilingual i18n)：** 支援「繁體中文 / English」一鍵切換，介面與 System Prompt 同步動態調適，支援生成全英文商務會議記錄。
* **📊 雙模式內容輸入：**
  * **模式 A (PPT 簡報)：** 自動解析簡報大標題、內文清單與備註欄 (Notes)。
  * **模式 B (文字草稿)：** 無 PPT 時，可直接於網頁輸入框按層級格式 (如 `1.1 ... 1.2 ...`) 貼上會議紀錄草稿。
* **🛡️ 金融防禦級安全架構 (Defensive Architecture)：**
  * **Magic Bytes 檔案校驗：** 檢查 OpenXML (`PK\x03\x04`) 壓縮標頭，嚴防惡意程式偽裝。
  * **強制 JSON 結構化約束：** API 底層強制 `response_format={"type": "json_object"}`，徹底消除輸出格式異常。
  * **Session 併發互斥鎖 (Mutex)：** 防止重複點擊造成 Token 浪費與 API Rate Limit 衝擊。
* **🔒 零數據留存資安架構 (Zero Data Retention)：** 採用純 Session 記憶體即時運算，關閉網頁後數據即刻徹底銷毀。

---

## 🛡️ ISO 資訊安全、隱私與 AI 治理聲明 (AIMS & Privacy Compliance)

本系統專為重視商業機密與數據隱私的企業團隊設計，合規架構明確對齊國際標準：

* **🔒 ISO/IEC 27001 (資訊安全管理體系 - ISMS) 對齊：**
  * **零數據留存 (Zero Data Retention / ZDR)：** 系統採用 Session-Only 記憶體運算架構。使用者上傳之 `.docx` 範本、`.pptx` 簡報及生成之會議記錄**絕不儲存於任何伺服器硬碟、資料庫或第三方儲存庫**。
  * **即時銷毀 (Instant Destruction)：** 當使用者關閉或重新整理瀏覽器分頁，所有運算記憶體與對話 Session 即刻徹底銷毀。
* **🔒 ISO/IEC 27701 (隱私資訊管理體系 - PIMS) 對齊：**
  * **數據最小化與本地 PII 假名化：** 資料送出 API 前，本地 Python 端 (`PIIMasker`) 自動遮蔽 HKID、台灣身分證、中國身分證、電話、Email 與薪資金額，收回回應後再於本地遞迴還原，第三方 LLM 0% 接觸敏感個資。
  * **無模型訓練 (No AI Model Training)：** 透過企業級 API 進行推論，傳輸數據嚴禁用於任何大語言模型的二次訓練。
* **🛡️ ISO/IEC 42001 (人工智慧管理體系 - AIMS) 對齊：**
  * **審計軌跡導出 (Audit Trail JSON)：** 自動於 Session 內記錄每次呼叫之時間戳、模型版本、去識別化 Token 數量與執行狀態，支援一鍵導出標準 JSON 審計報告供企業資安稽核員 (CISO/Auditor) 查驗。
* **🌐 傳輸安全 (TLS/SSL Encryption)：** 全程採用 HTTPS / TLS 1.3 國際標準加密傳輸，確保資料在傳送過程免受中間人攔截。

### 📋 完整 ISO 42001 & EU AI Act 審計文件庫 (Audit Documentation)
* 📇 **[MODEL_CARD.md](MODEL_CARD.md)** — 核心模型規格與雙語 Prompt 調校軌跡
* 🛡️ **[docs/RISK_ASSESSMENT.md](docs/RISK_ASSESSMENT.md)** — AI 幻覺、API 外洩與流量資安風險矩陣
* 🔒 **[docs/DATA_GOVERNANCE.md](docs/DATA_GOVERNANCE.md)** — ZDR 零留存與 PII 本地端動態假名化政策
* 👤 **[docs/HUMAN_OVERSIGHT.md](docs/HUMAN_OVERSIGHT.md)** — 人工審核 (HITL)、平滑回退與流量控管機制
* 🔄 **[docs/LIFECYCLE_MANAGEMENT.md](docs/LIFECYCLE_MANAGEMENT.md)** — 模型升級、Prompt 變更控制與版本履歷

---

## 📖 使用指南 (User Guide)

1. **選擇結構範本：** 
   * **方式 A：** 選擇「內建標準範本」，並下拉挑選週會、董事會或專案檢討型格式。
   * **方式 B：** 選擇「自行上傳」，上傳貴單位過往使用的會議記錄 (.docx) 作為骨架（系統自動檢驗 Magic Bytes 與 3 欄表格）。
2. **提供本次會議內容 (二選一)：**
   * **選擇 A：** 上傳本次會議簡報 (`.pptx`)。
   * **選擇 B：** 直接於文字框貼上會議草稿 (按 `1. 議題` -> `1.1 子項目` 層級輸入)。
3. **生成與下載：** 
   * 點擊「🚀 即刻依範本結構生成會議記錄」。
   * 預覽無誤後，可一鍵下載：
     * 📄 **標準 Word 格式會議記錄 (.docx)**
     * 📥 **Markdown 原始檔 (.md)**
     * 🛡️ **ISO 42001 審計日誌 (.json)**

---

## 🚀 企業級功能演進藍圖 (Enterprise Roadmap)

* **Phase 1: 資安與數據主權強化 (Security & Data Sovereignty)**
  * [x] **PII 本地動態遮蔽 (Local PII Redaction):** 於送出 API 前進行本地端 HKID、電話、Email、薪資等敏感個資正則遮蔽，實現「傳輸前假名化 (Pseudonymization)」。
  * [x] **自備 API Key (BYOK - Bring Your Own Key):** 提供企業自選 OpenAI / Azure OpenAI / GitHub Models 私有端點。
  * [x] **API 防洩漏與流量控管:** 攔截詳細 Exception 報錯防 Header 外洩，並加入免費通道 Session 10 次配額限制。
  * [x] **檔案 Magic Bytes 防偽校驗:** 檢查檔案頭簽名 `PK\x03\x04`，防範惡意偽裝檔案上傳。
* **Phase 2: 產品體驗與範本忠實度升級 (User Experience & Fidelity)**
  * [x] **原檔範本忠實留存 (`fill_user_template_from_json`):** 導入模板克隆與 OpenXML 結構注入，直接回填文字至用戶上傳的 Word 檔，100% 完整保留企業 Logo、頁首頁尾與專屬樣式。
  * [x] **多語言雙語支援 (Bilingual i18n):** 支援「繁體中文 / English」介面與 LLM System Prompt 動態雙語切換。
  * [x] **API JSON Mode 強制約束:** 鎖定 API 回傳純 JSON 物件，消除 Text Parsing 異常。
  * [ ] **區塊化對話微調 (Section-based Chat Refinement):** 允許使用者點選單一章節進行局部 AI 對話修訂。
* **Phase 3: 治理與稽核軌跡 (AI Governance & Audit Trail)**
  * [x] **ISO 42001 Model Card 歸檔:** 完成 [MODEL_CARD.md](MODEL_CARD.md) 審計文件規範。
  * [x] **結構化稽核日誌 (Compliance Audit Trail):** 記憶體內產生 Session-Only 的操作與 PII 遮蔽摘要日誌，支援一鍵導出 JSON 供企業資安稽核員 (CISO/Auditor) 備查。

---

## 🛠️ 技術架構 (Tech Stack)

* **前端與 UI：** Streamlit
* **大語言模型 API：** GitHub Models / OpenAI / Azure OpenAI (`gpt-4o-mini`)
* **資料解析與處理：** `python-docx` (OpenXML 表格生成與樣式繼承) + `python-pptx` (簡報解析)
* **資安與合規：** Magic Bytes 驗證、正則防二次污染 PII Masker、Session-Only 記憶體隔離

---

## 💡 技術支援與聯絡方式 (Support)

本專案由 [Jacky Law](https://jackylawck.github.io/jackylawck/) 開發與維護。如有任何系統建置、商業部署或 AI 治理諮詢問題，歡迎透過連結聯繫。

---

## 📄 授權說明 (License)

本專案採用 **MIT License** 開源授權。系統僅供行政與會議記錄整理輔助使用，所有資料均於本地瀏覽器 Session 記憶體內處理，不存檔至任何外部雲端資料庫。

```
