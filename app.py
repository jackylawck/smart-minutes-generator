import os
import io
import re
import json
import datetime
import streamlit as st
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from pptx import Presentation
import httpx

# ==========================================
# 0. Session State 初始化
# ==========================================
if "generated_minutes_json" not in st.session_state:
    st.session_state["generated_minutes_json"] = None

if "generated_minutes_md" not in st.session_state:
    st.session_state["generated_minutes_md"] = ""

if "generation_count" not in st.session_state:
    st.session_state["generation_count"] = 0

if "is_processing" not in st.session_state:
    st.session_state["is_processing"] = False

if "audit_log" not in st.session_state:
    st.session_state["audit_log"] = []

# ==========================================
# 1. 頁面配置與多語言 (i18n) 字典
# ==========================================
st.set_page_config(
    page_title="Smart Minutes Generator | 智能會議記錄生成器",
    page_icon="📝",
    layout="wide"
)

with st.sidebar:
    lang = st.radio("🌐 Language / 語言", ["繁體中文", "English"], horizontal=True)

I18N = {
    "繁體中文": {
        "title": "📝 智能會議記錄生成器 (Smart Minutes Generator)",
        "caption": "金融防禦級架構：原生推論引擎、Magic Bytes 校驗、PII 深度假名化與 ISO 42001 審計日誌",
        "sidebar_template_download": "📥 下載基準範本 (.docx)",
        "sidebar_guidelines": "📖 使用方式與格式指引",
        "sidebar_guide_a": "<b>📊 模式 A：上傳 PPT 簡報</b><br>• 排版建議：頁面頂部為大議題，中間為子題目。<br>• 顯示邏輯：自動填入範本結構。",
        "sidebar_guide_b": "<b>📝 模式 B：文字草稿輸入</b><br>• 使用情境：無 PPT 時，按議題層級貼上草稿。<br>• 處理邏輯：AI 自動歸納至範本結構。",
        "sidebar_iso": "🔒 **ISO 數據安全聲明：** 本系統採 Session-Only 記憶體運算，關閉即銷毀。支援合規審計日誌導出。",
        "byok_title": "🔑 企業資安選項：自備 API Key (BYOK)",
        "byok_toggle": "啟用自備 API Key (自帶企業 API 通道)",
        "byok_provider": "API 供應商類型",
        "byok_key": "輸入 API Key / Token",
        "byok_model": "模型名稱",
        "byok_url": "Base URL",
        "byok_url_hint": "💡 GitHub Models 官方端點：`https://models.github.ai/inference`",
        "byok_model_hint": "💡 GitHub Models 請使用 `openai/gpt-4o-mini`",
        "sec_template": "📁 1. 選擇或上傳會議記錄格式範本",
        "template_src": "請選擇會議記錄結構來源：",
        "template_builtin": "內建標準範本 (免上傳)",
        "template_custom": "自行上傳自訂 Word 範本 (.docx)",
        "template_type": "選擇內建商務範本類型：",
        "upload_custom_tpl": "請上傳作為結構基準的 Word 檔案 (.docx)",
        "success_custom_tpl": "成功載入並校驗自訂範本！匯出時將保留專屬 Logo 與排版。",
        "warn_no_table": "⚠️ 偵測到自訂範本內未包含 3 欄式表格。匯出時系統將自動平滑切換至標準排版，保留完整內容。",
        "sec_content": "📊 2. 輸入本次會議內容來源 (選擇 A 或 選擇 B)",
        "mode_a": "選擇 A：上傳本次會議簡報 (.pptx)",
        "mode_b": "選擇 B：貼上會議草稿 (無 PPT 時使用)",
        "placeholder_b": "若無 PPT，請依照議題與編號輸入草稿內容，例如：\n1. 會議摘要\n1.1 通過上次會議紀錄。\n1.2 確認季度營運目標達成率為 95%。",
        "btn_generate": "🚀 即刻依範本結構生成會議記錄",
        "btn_processing": "⏳ 正在進行深度去識別化與結構化生成...",
        "preview_title": "📋 會議記錄預覽 (Preview)",
        "btn_download_docx_custom": "📄 下載標準 Word 記錄 (保留自訂範本 Logo 與格式)",
        "btn_download_docx_builtin": "📄 下載標準 Word 記錄 (內建商務格式)",
        "btn_download_docx_fallback": "📄 下載標準 Word 記錄 (安全回退格式)",
        "btn_download_md": "📥 下載 Markdown 原始檔 (.md)",
        "btn_download_audit": "🛡️ 導出 ISO 42001 審計日誌 (.json)",
        "err_no_template": "🛑 請務必選擇有效的內建範本或上傳自訂 Word 範本！",
        "err_no_content": "🛑 請提供本次會議內容，上傳 PPT 簡報（選擇 A）或輸入會議草稿（選擇 B）。",
        "err_no_key": "🛑 未檢測到有效的 API Key。請確認 Secrets 已設定 GITHUB_TOKEN，或開啟 BYOK 輸入密鑰。",
        "err_rate_limit": "🛑 免費額度已達上限（每位訪客限 10 次）。請開啟上方 BYOK 輸入專屬 API Key。",
        "err_429_ratelimit": "🛑 觸發 API 流量限制 (HTTP 429)。請稍候 30 秒後重試，或開啟上方 BYOK 切換至企業專屬通道。",
        "err_connection": "🛑 無法連線至 API 端點。請確認端點為 `https://models.github.ai/inference`。",
        "err_file_security": "🛑 檔案安全校驗失敗：檔案非合法 OpenXML 格式或已損毀。",
        "msg_success": "✨ 本次會議記錄已成功生成 (完成本地 PII 去識別化與結構化對齊)！",
        "msg_fallback": "⚠️ 自訂範本特殊格式套用失敗，已自動平滑回退至標準排版。",
        "fallback_item_topic": "系統已完成會議內容比對，未檢測到與範本對應之明確議題，建議檢查範本結構或輸入草稿。",
        "fallback_item_res": "（記錄）",
        "system_prompt": """你是一名精通企業行政與結構化合規管理的高級秘書。你的任務是進行「動態結構映射與內容提煉」。
請務必輸出嚴格合法的 JSON 物件格式：
{
  "title": "會議記錄標題",
  "meta_info": {
    "date": "",
    "location": "",
    "chairperson": "",
    "secretary": ""
  },
  "agenda_items": [
    {
      "id": "1.1",
      "topic": "議題內容摘要",
      "resolution": "（通過）/（記錄）/（跟進）"
    }
  ]
}
注意：
1. 嚴禁輸出多餘的 Markdown 贅字說明。
2. 文中包含的 [REDACTED_...] 標籤必須完全原樣留存。語言：繁體中文。"""
    },
    "English": {
        "title": "📝 Smart Minutes Generator",
        "caption": "Enterprise Defense Architecture: Native Inference Engine, Magic Bytes Validation, Deep PII Masking & ISO 42001 Audit Logging",
        "sidebar_template_download": "📥 Download Baseline Templates (.docx)",
        "sidebar_guidelines": "📖 User Guide & Format Rules",
        "sidebar_guide_a": "<b>📊 Mode A: Upload PPT Presentation</b><br>• Layout: Agenda topics at top, detailed bullet points below.<br>• Logic: Automatically maps content into template structure.",
        "sidebar_guide_b": "<b>📝 Mode B: Paste Draft Text</b><br>• Use Case: Paste raw draft using structured bullet levels.<br>• Logic: Summarizes text into matched agenda items.",
        "sidebar_iso": "🔒 **ISO Security Statement:** Session-Only memory execution. Compliant with ISO 42001 audit logging.",
        "byok_title": "🔑 Enterprise Option: Bring Your Own Key (BYOK)",
        "byok_toggle": "Enable BYOK Mode (Use Custom API Channel)",
        "byok_provider": "API Provider Type",
        "byok_key": "Enter API Key / Token",
        "byok_model": "Model Name",
        "byok_url": "Base URL",
        "byok_url_hint": "💡 GitHub Models official endpoint: `https://models.github.ai/inference`",
        "byok_model_hint": "💡 GitHub Models requires `openai/gpt-4o-mini`",
        "sec_template": "📁 1. Select or Upload Meeting Minutes Template",
        "template_src": "Select template source:",
        "template_builtin": "Built-in Standard Template",
        "template_custom": "Upload Custom Word Template (.docx)",
        "template_type": "Select Built-in Template Style:",
        "upload_custom_tpl": "Upload a Word file (.docx) as structural baseline",
        "success_custom_tpl": "Custom template validated! Layout and logo will be preserved upon export.",
        "warn_no_table": "⚠️ No 3-column table detected in custom template. System will fall back to standard layout on export.",
        "sec_content": "📊 2. Input Meeting Content (Option A or Option B)",
        "mode_a": "Option A: Upload Meeting Presentation (.pptx)",
        "mode_b": "Option B: Paste Meeting Draft Text",
        "placeholder_b": "Paste draft text using agenda hierarchy, e.g.:\n1. Executive Summary\n1.1 Approved previous meeting minutes.\n1.2 Confirmed Q2 target achievement rate at 95%.",
        "btn_generate": "🚀 Generate Meeting Minutes Now",
        "btn_processing": "⏳ Processing PII redaction and structured synthesis...",
        "preview_title": "📋 Meeting Minutes Preview",
        "btn_download_docx_custom": "📄 Download Word (.docx) - Preserving Custom Template Logo/Style",
        "btn_download_docx_builtin": "📄 Download Word (.docx) - Standard Corporate Style",
        "btn_download_docx_fallback": "📄 Download Word (.docx) - Safe Fallback Style",
        "btn_download_md": "📥 Download Markdown Raw File (.md)",
        "btn_download_audit": "🛡️ Export ISO 42001 Audit Log (.json)",
        "err_no_template": "🛑 Please select a built-in template or upload a custom Word template!",
        "err_no_content": "🛑 Please provide meeting content via PPTX upload (Option A) or text draft (Option B).",
        "err_no_key": "🛑 No valid API Key detected. Please configure secrets or enable BYOK mode.",
        "err_rate_limit": "🛑 Free trial usage limit reached (10 generations per session). Please enable BYOK mode.",
        "err_429_ratelimit": "🛑 Rate limit reached (HTTP 429). Please wait 30 seconds or enable BYOK.",
        "err_connection": "🛑 Cannot connect to API endpoint. For GitHub Models, ensure `https://models.github.ai/inference`.",
        "err_file_security": "🛑 File validation failed: The uploaded file is not a valid OpenXML document.",
        "msg_success": "✨ Meeting minutes successfully generated (with local PII redaction and structure mapping)!",
        "msg_fallback": "⚠️ Custom template filling failed. Automatically falling back to standard format.",
        "fallback_item_topic": "Meeting content summarized. No direct matching agenda topics detected. Please verify template hierarchy.",
        "fallback_item_res": "(Noted)",
        "system_prompt": """You are an executive assistant specializing in corporate governance. Your task is "Dynamic Structure Mapping and Summarization".
Please respond with a valid JSON format:
{
  "title": "Meeting Minutes Title",
  "meta_info": {
    "date": "",
    "location": "",
    "chairperson": "",
    "secretary": ""
  },
  "agenda_items": [
    {
      "id": "1.1",
      "topic": "Topic summary",
      "resolution": "(Approved) / (Noted) / (Action Required)"
    }
  ]
}
Preserve [REDACTED_...] tokens exactly. Language: Professional Corporate English."""
    }
}

t = I18N[lang]

# ==========================================
# 2. 🛡️ 檔案安全檢驗
# ==========================================
def validate_openxml_magic(file) -> bool:
    if file is None:
        return True
    try:
        file.seek(0)
        header = file.read(4)
        file.seek(0)
        return header == b'PK\x03\x04'
    except Exception:
        return False

def has_valid_table(file) -> bool:
    try:
        file.seek(0)
        doc = Document(io.BytesIO(file.read()))
        file.seek(0)
        for table in doc.tables:
            if len(table.columns) >= 3:
                return True
        return False
    except Exception:
        return False

# ==========================================
# 3. 🛡️ PII 遮蔽器
# ==========================================
class PIIMasker:
    PATTERNS = [
        ('HKID',   r'\b[A-Z]{1,2}\d{6}[\(]?\d?[\)]?[A-Z]?\b'),
        ('TW_ID',  r'\b[A-Z][12]\d{8}\b'),
        ('CN_ID',  r'\b\d{17}[\dXx]\b'),
        ('EMAIL',  r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b'),
        ('PHONE',  r'(?:\+?852[\s\-]?)?(?:\b\d{4}[\s\-]?\d{4}\b|\b\d{8}\b)'),
        ('SALARY', r'(?:薪資|薪|月薪|年薪|salary)[\s：:]*(?:HK\$|NT\$|¥|\$)?\s*[\d,]+(?:K|k|萬)?'),
    ]

    def __init__(self):
        self.vault = {}
        self.counters = {}

    def mask(self, text: str) -> str:
        if not text:
            return text
        self.vault.clear()
        self.counters.clear()
        out = text
        for typ, pat in self.PATTERNS:
            def _repl(m):
                val = m.group(0)
                if "[REDACTED_" in val:
                    return val
                self.counters[typ] = self.counters.get(typ, 0) + 1
                tok = f"[REDACTED_{typ}_{self.counters[typ]}]"
                self.vault[tok] = val
                return tok
            out = re.sub(pat, _repl, out, flags=re.IGNORECASE)
        return out

    def unmask(self, text: str) -> str:
        if not text or not self.vault:
            return text
        out = text
        for tok in sorted(self.vault.keys(), key=len, reverse=True):
            out = out.replace(tok, self.vault[tok])
        return out

    def unmask_json(self, data):
        if isinstance(data, dict):
            return {k: self.unmask_json(v) for k, v in data.items()}
        elif isinstance(data, list):
            return [self.unmask_json(v) for v in data]
        elif isinstance(data, str):
            return self.unmask(data)
        return data

# ==========================================
# 4. 側邊欄
# ==========================================
with st.sidebar:
    st.markdown(f"**{t['sidebar_template_download']}**")
    templates_to_download = [
        {"file": "template_weekly.docx", "label": "📄 Weekly / 例會範本", "out": "Weekly_Meeting_Template.docx"},
        {"file": "template_board.docx", "label": "🏛️ Board / 董事會範本", "out": "Board_Meeting_Template.docx"},
        {"file": "template_project.docx", "label": "📊 Project / 專案檢討範本", "out": "Project_Review_Template.docx"}
    ]

    for item in templates_to_download:
        possible_paths = [item["file"], f"templates/{item['file']}"]
        found_path = next((p for p in possible_paths if os.path.exists(p)), None)
        if found_path:
            with open(found_path, "rb") as f:
                st.download_button(
                    label=item["label"],
                    data=f,
                    file_name=item["out"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        else:
            st.caption(f"⚠️ `{item['file']}` missing")

    st.markdown("---")
    st.markdown(f"**{t['sidebar_guidelines']}**")
    st.markdown(t['sidebar_guide_a'], unsafe_allow_html=True)
    st.markdown(t['sidebar_guide_b'], unsafe_allow_html=True)
    st.markdown("---")
    st.caption(t["sidebar_iso"])
    st.markdown("---")
    st.markdown("<div style='font-size: 0.8em; color: #a0aec0;'>💡 Lead Architect: <a href='https://jackylawck.github.io/jackylawck/' target='_blank' style='color: #63b3ed;'>Jacky Law</a></div>", unsafe_allow_html=True)

# ==========================================
# 5. 主畫面與 BYOK
# ==========================================
st.title(t["title"])
st.caption(t["caption"])

with st.expander(t["byok_title"], expanded=False):
    use_byok = st.toggle(t["byok_toggle"], value=st.session_state.get("use_byok", False), key="byok_toggle")
    st.session_state["use_byok"] = use_byok

    if use_byok:
        api_provider = st.selectbox(t["byok_provider"], ["GitHub Models", "OpenAI", "Azure OpenAI"], key="byok_provider")
        byok_key = st.text_input(t["byok_key"], type="password", placeholder="sk-... / ghp_... / github_pat_...", key="byok_key")
        
        default_model = "openai/gpt-4o-mini" if api_provider == "GitHub Models" else "gpt-4o-mini"
        byok_model = st.text_input(t["byok_model"], value=default_model, key="byok_model")
        if api_provider == "GitHub Models":
            st.caption(t["byok_model_hint"])

        default_url = "https://models.github.ai/inference" if api_provider == "GitHub Models" else ""
        byok_url = st.text_input(t["byok_url"], value=default_url, key="byok_url")
        if api_provider == "GitHub Models":
            st.caption(t["byok_url_hint"])
    else:
        raw_secret = st.secrets.get("GITHUB_TOKEN", "")
        byok_key = str(raw_secret).strip().replace("\n", "").replace("\r", "").strip("\"'")
        byok_model = "openai/gpt-4o-mini"
        byok_url = "https://models.github.ai/inference"

# ==========================================
# 6. 🚀 原生安全推論請求 (絕不走淘汰 Azure 轉址)
# ==========================================
def call_chat_completion(api_key: str, base_url: str, model_name: str, messages: list) -> str:
    key = api_key.strip()
    url = base_url.strip().rstrip("/") if base_url else "https://models.github.ai/inference"
    
    # 強制替換任何殘留的 Azure 域名
    if "models.inference.ai.azure.com" in url:
        url = "https://models.github.ai/inference"

    if not url.endswith("/chat/completions"):
        full_endpoint = f"{url}/chat/completions"
    else:
        full_endpoint = url

    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "User-Agent": "SmartMinutesGenerator/1.0",
        "Accept": "application/json"
    }

    payload = {
        "model": model_name.strip(),
        "messages": messages,
        "temperature": 0.2
    }

    with httpx.Client(timeout=90.0, follow_redirects=False) as client:
        response = client.post(full_endpoint, headers=headers, json=payload)
        
        if response.status_code != 200:
            raise Exception(f"HTTP {response.status_code} - {response.text}")
            
        res_json = response.json()
        return res_json["choices"][0]["message"]["content"]

# ==========================================
# 7. 強健自癒解析引擎
# ==========================================
def parse_llm_output_to_data(raw_text: str):
    clean_text = raw_text.strip()
    
    match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*
