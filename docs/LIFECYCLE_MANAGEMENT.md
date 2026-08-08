# 🔄 AI Lifecycle & Change Management (生命週期與變更管理)

> **Compliance Framework:** ISO/IEC 42001:2023 Control A.6 (AI System Lifecycle)  
> **Document Control ID:** AIMS-LM-2026-001  
> **System Name:** Smart Minutes Generator

---

## 1. 📌 模型與 Prompt 變更控制 (Change Control)

所有影響 AI 輸出品質與合規性之變更，均須經過以下測試與驗證程序：

1. **基礎模型評估 (Model Evaluation):** 從 `gpt-4o` 調整至 `gpt-4o-mini` 之變更，須通過基準測試，確保摘要精確度無明顯下降，且推論速度與 Token 成本達到最佳平衡。
2. **雙語 System Prompt 版本控管:** 
   * **`zh-HK` 語系:** 強制規範企業級繁體中文語彙與「通過/記錄/跟進」三分類。
   * **`en-US` 語系:** 強制規範 Corporate English 與 "(Approved)/(Noted)/(Action Required)" 標籤。

---

## 2. 📜 系統版本變更履歷 (Change Audit Trail)

* **v1.0.0 (2026-04):** PoC 發布，支援動態範本映射與基礎 Word 匯出。
* **v1.1.0 (2026-05):** 導入 ISO 27001 ZDR 聲明與 OpenXML 商業級表格防斷裂排版 (`<w:cantSplit/>`)。
* **v1.2.0 (2026-07):** **ISO 42001 治理升級** — 實作本地端 PII 假名化 (`PIIMasker`)、BYOK 自備通道、API 報錯防洩漏與 Session Rate Limiting。
* **v1.3.0 (2026-08):** **多語言與完整文檔閉環** — 導入繁/英雙語切換 (i18n)，並完成 `docs/` 合規文件庫建立。
