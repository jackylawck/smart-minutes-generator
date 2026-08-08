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

## 2. 🛡️ 本地端 PII 假名化規範 (Pre-flight Pseudonymization)
在任何數據離開伺服器送往公有雲端 API 之前，本地 Python 端 PIIMasker 會強制進行敏感個資遮蔽：

香港身份證 (HKID): 遮蔽為 [REDACTED_HKID_N]

電話號碼 (Phone): 遮蔽為 [REDACTED_PHONE_N]

電子郵件 (Email): 遮蔽為 [REDACTED_EMAIL_N]

薪資金額 (Salary): 遮蔽為 [REDACTED_SALARY_N]

## 3. 🚫 無模型再訓練條款 (No AI Model Training Clause)
本系統所選用之企業級 API 通道（GitHub Models / Azure OpenAI Service）明確包含 No Model Training 協議。

所有過境數據嚴禁被供應商用於改善或再訓練底層大語言模型。

