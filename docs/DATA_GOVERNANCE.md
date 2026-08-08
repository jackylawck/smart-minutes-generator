# 🔒 Data Governance & Privacy Policy (數據與隱私治理政策)

> **Compliance Framework:** ISO/IEC 27001 (ISMS), ISO/IEC 27701 (PIMS) & GDPR Article 25  
> **Document Control ID:** AIMS-DG-2026-001  
> **System Name:** Smart Minutes Generator

---

## 1. 📌 數據生命週期與零留存架構 (Zero Data Retention / ZDR)

本系統採 **Privacy-by-Design** 設計，所有資料處理由起點至終點完全遵守「數據最小化 (Data Minimization)」原則：

```text
[使用者上傳 .pptx/.docx] ➔ [RAM 記憶體運算] ➔ [本地 PII 假名化] ➔ [API 摘要] ➔ [本地 PII 還原] ➔ [匯出 Word] ➔ [Session 關閉即刻徹底銷毀]
```

無硬碟殘留 (Zero Disk Persistence): 上傳檔與生成之會議記錄僅存放於伺服器虛擬記憶體 (RAM)。

即時銷毀 (Instant Destruction): 網頁關閉、重新整理或 Session 過期時，記憶體自動觸發垃圾回收 (Garbage Collection) 銷毀所有數據。

2. 🛡️ 本地端 PII 假名化規範 (Pre-flight Pseudonymization)
在任何數據離開伺服器送往公有雲端 API 之前，本地 Python 端 PIIMasker 會強制進行敏感個資遮蔽：

香港身份證 (HKID): 遮蔽為 [REDACTED_HKID_N]

電話號碼 (Phone): 遮蔽為 [REDACTED_PHONE_N]

電子郵件 (Email): 遮蔽為 [REDACTED_EMAIL_N]

薪資金額 (Salary): 遮蔽為 [REDACTED_SALARY_N]

3. 🚫 無模型再訓練條款 (No AI Model Training Clause)
本系統所選用之企業級 API 通道（GitHub Models / Azure OpenAI Service）明確包含 No Model Training 協議。

所有過境數據嚴禁被供應商用於改善或再訓練底層大語言模型。


---

#### 3. 建立 `docs/HUMAN_OVERSIGHT.md`
```markdown
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
* **用量超額控制 (Rate Limiting):** 單一 Session 使用免費通道超過 10 次時強制鎖定，要求切換至 BYOK 模式，防範資
