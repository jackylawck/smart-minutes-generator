# 🛡️ AI Risk Assessment Report (風險評估報告)

> **Compliance Framework:** ISO/IEC 42001:2023 Annex A.5 (AI Risk Assessment) & EU AI Act Article 9  
> **Document Control ID:** AIMS-RA-2026-001  
> **System Name:** Smart Minutes Generator

---

## 1. 📌 風險評估範圍 (Scope)
本報告針對 Smart Minutes Generator 進行 AI 專屬風險識別、機率與衝擊分析，並記錄對應之技術緩釋措施 (Mitigation Controls)。

---

## 2. 📊 AI 風險矩陣與緩釋措施 (Risk Matrix & Mitigations)

| 風險項目 (Risk Event) | 原始風險 (Impact/Likelihood) | 緩釋控制措施 (Mitigation Control) | 殘餘風險 (Residual Risk) |
| :--- | :--- | :--- | :--- |
| **1. AI 幻覺與內容捏造 (Hallucination)** | 🔴 **High** (中機率/高衝擊) | • 超參數設定 `temperature=0.2` 極低隨機性。<br>• System Prompt 強制規範：「若某議題未提及，必須填寫『本次會議暫無相關事項』，不可捏造」。<br>• 實施 Human-in-the-Loop (HITL)，最終紀錄須經人工簽核。 | 🟢 **Low** |
| **2. API 金鑰與敏感 Header 洩漏** | 🔴 **High** (低機率/極高衝擊) | • 程式碼 (`app.py`) 攔截詳細 Exception，報錯僅顯示 `type(e).__name__`，嚴禁印出 HTTP Headers。<br>• 免費系統 Key 實施 Session 10 次流量限制。<br>• 支援 BYOK 模式，金鑰存於記憶體不留檔。 | 🟢 **Low** |
| **3. PII 個資外洩至第三方 LLM** | 🟡 **Medium** (中機率/中衝擊) | • 傳輸前本地端 `PIIMasker` 正則表達式自動遮蔽 HKID、電話、Email 與薪資。<br>• 本地端收回結果後再行還原，確保第三方 LLM 0% 接觸真實個資。 | 🟢 **Low** |
| **4. 自訂範本格式解析崩潰** | 🟡 **Medium** (中機率/低衝擊) | • 實施 `try-except` 平滑回退機制 (Fallback Mechanism)，若自訂 Word 解析失敗，自動切換至內建標準深藍排版，確保服務不中斷。 | 🟢 **Low** |

---

## 3. ⚖️ 歐盟 AI 法案風險分類 (EU AI Act Classification)
* **分類結果:** **Limited Risk / Minimal Risk** (非 High-Risk AI 系統)。
* **透明度義務 (Article 50):** 系統於 UI 顯眼處標明 AI 生成提示，並提供完整 `MODEL_CARD.md` 供稽核查驗。
