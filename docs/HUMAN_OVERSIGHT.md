# 👤 Human Oversight & Fallback Protocol (人工監督與平滑回退機制)

> **Compliance Framework:** EU AI Act Article 14 (Human Oversight) & ISO/IEC 42001 Control A.8.4  
> **Document Control ID:** AIMS-HO-2026-001  
> **System Name:** Smart Minutes Generator

---

## 1. 👥 人機協同原則 (Human-in-the-Loop / HITL)

本系統定位為 **行政輔助 Copilot 工具**，無自主決策權限。為確保合規，系統設計以下監督機制：

1. **草稿性質 (Draft Status):** AI 生成之會議記錄於 UI 界面明確標示為「預覽草稿 (Preview Draft)」。
2. **人工覆核與簽核 (Human Review & Approval):** 最終 Word (.docx) 文件必須由會議主席、秘書或授權行政人員親自覆核內容正確性後，方可作正式提交或存檔。

---

## 2. 🚨 緊急平滑回退機制 (Fallback & Kill Switch)

當系統偵測到非預期例外或格式異常時，會觸發自動防護架構：

* **自訂範本解析失敗時:** 系統自動捕獲 Exception，顯示提示警告，並切換至內建標準 OpenXML 表格渲染機制（`convert_md_to_docx`），確保使用者永遠能取得完整文件。
* ** API 異常與金鑰安全:** 若 API 連線中斷或授權失敗，系統僅印出 `Error Type`，拒絕暴露底層堆疊細節或 Header 資訊。
* **用量超額控制 (Rate Limiting):** 單一 Session 使用免費通道超過 10 次時強制鎖定，要求切換至 BYOK 模式，防範資源濫用。
