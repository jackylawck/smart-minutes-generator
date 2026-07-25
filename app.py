import os
import io
import re
import streamlit as st
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from pptx import Presentation
from openai import OpenAI

# ==========================================
# 0. Session State 初始化
# ==========================================
if "generated_minutes" not in st.session_state:
    st.session_state["generated_minutes"] = ""

# ==========================================
# 1. 頁面配置與美化 CSS (支援深色/淺色模式)
# ==========================================
st.set_page_config(
    page_title="智能會議記錄生成器",
    page_icon="📝",
    layout="wide"
)

st.markdown("""
    <style>
    .sidebar-title {
        font-weight: bold;
        font-size: 1.2em;
        margin-bottom: 10px;
    }
    .guideline-card {
        background-color: rgba(255, 255, 255, 0.08);
        color: inherit;
        padding: 12px;
        border-radius: 6px;
        border: 1px solid rgba(128, 128, 128, 0.3);
        border-left: 4px solid #0056b3;
        margin-bottom: 12px;
        font-size: 0.88em;
        line-height: 1.5;
    }
    .guideline-card b {
        color: #3182ce;
    }
    table {
        width: 100% !important;
        border-collapse: collapse;
    }
    th:nth-child(1), td:nth-child(1) { width: 8% !important; text-align: center !important; }
    th:nth-child(2), td:nth-child(2) { width: 77% !important; }
    th:nth-child(3), td:nth-child(3) { width: 15% !important; text-align: center !important; white-space: nowrap !important; }
    
    .footer-support {
        font-size: 0.8em;
        color: #a0aec0;
        margin-top: 15px;
    }
    .footer-support a {
        color: #63b3ed;
        text-decoration: underline;
    }
    .footer-support a:hover {
        color: #90cdf4;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 🛡️ 本地端 PII 敏感數據遮蔽器 (ISO 27701 假名化)
# ==========================================
class PIIMasker:
    PATTERNS = {
        'HKID':    r'\b[A-Z]{1,2}\d{6}[\(]?\d?[\)]?[A-Z]?\b',
        'TW_ID':   r'\b[A-Z][12]\d{8}\b',
        'CN_ID':   r'\b\d{17}[\dXx]\b',
        'EMAIL':   r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}',
        'PHONE':   r'(?:\+?852[\s\-]?)?(?:\d{4}[\s\-]?\d{4}|\d{8})',
        'SALARY':  r'(?:薪資|薪|月薪|年薪|salary)[\s：:]*(?:HK\$|NT\$|¥|\$)?\s*[\d,]+(?:K|k|萬)?',
    }

    def __init__(self):
        self.vault = {}
        self.counters = {}

    def mask(self, text: str) -> str:
        if not text:
            return text
        self.vault.clear()
        self.counters.clear()
        out = text
        for typ, pat in self.PATTERNS.items():
            def _repl(m):
                self.counters[typ] = self.counters.get(typ, 0) + 1
                tok = f"[REDACTED_{typ}_{self.counters[typ]}]"
                self.vault[tok] = m.group(0)
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

# ==========================================
# 3. 🛡️ 左側側邊欄：商務範本下載 & 指引
# ==========================================
with st.sidebar:
    st.markdown("<div class='sidebar-title'>📥 下載基準範本 (.docx)</div>", unsafe_allow_html=True)
    
    templates_to_download = [
        {"file": "template_weekly.docx", "label": "📄 下載：週會例會範本", "out": "週會例會會議記錄範本.docx"},
        {"file": "template_board.docx", "label": "🏛️ 下載：董事會/高層決議範本", "out": "董事會高層會議記錄範本.docx"},
        {"file": "template_project.docx", "label": "📊 下載：專案跟進檢討範本", "out": "專案檢討會議記錄範本.docx"}
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
            st.caption(f"⚠️ 未偵測到 `{item['file']}`")

    st.markdown("---")
    st.markdown("<div class='sidebar-title'>📖 使用方式與格式指引</div>", unsafe_allow_html=True)
    st.markdown("本系統支援兩種會議內容輸入模式（二選一）：")
    
    st.markdown("""
    <div class='guideline-card'>
    <b>📊 模式 A：上傳 PPT 簡報</b><br>
    • <b>排版建議：</b>頁面頂部為大議題，中間為子題目，下方為詳細文字。<br>
    • <b>顯示邏輯：</b>系統會自動將內容填入範本結構；若 PPT 頁面僅有子題目而無詳細文字，紀錄中將直接顯示該子題目。
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class='guideline-card'>
    <b>📝 模式 B：文字草稿輸入 (無 PPT 時使用)</b><br>
    • <b>使用情境：</b>若無 PPT 簡報，可直接在草稿框按層級格式（如：大議題 -> 1.1 ... 1.2 ...）寫下文字草稿。<br>
    • <b>處理邏輯：</b>AI 會將草稿內容精準歸納並整理至對應的議題架構中。
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.caption("🔒 **ISO 數據安全聲明：** 本系統採 Session-Only 記憶體即時運算，關閉網頁數據即刻徹底銷毀。所有 PII 資料在送出 API 前均已本地假名化。")
    
    st.markdown("---")
    st.markdown(
        "<div class='footer-support'>💡 如有系統使用問題，歡迎聯絡 <a href='https://jackylawck.github.io/jackylawck/' target='_blank'>Jacky Law</a>。</div>", 
        unsafe_allow_html=True
    )

# ==========================================
# 4. 主畫面介面與 BYOK (自備 API Key) 設定
# ==========================================
st.title("📝 智能會議記錄生成器 (Smart Minutes Generator)")
st.caption("基於動態模板映射、PII 本地假名化與 BYOK 企業資安設計")

with st.expander("🔑 企業資安選項：自備 API Key (BYOK)", expanded=False):
    use_byok = st.toggle("啟用自備 API Key (自帶企業 API 通道)", value=st.session_state.get("use_byok", False), key="byok_toggle")
    st.session_state["use_byok"] = use_byok

    if use_byok:
        api_provider = st.selectbox("API 供應商類型", ["GitHub Models", "OpenAI", "Azure OpenAI"], key="byok_provider")
        byok_key = st.text_input("輸入 API Key / Token", type="password", placeholder="sk-... 或 ghp_...", key="byok_key")
        byok_model = st.text_input("模型名稱", value=st.session_state.get("byok_model", "gpt-4o-mini"), key="byok_model")
        byok_url = st.text_input("Base URL (Azure OpenAI 或自架端點必填)", value=st.session_state.get("byok_url", ""), placeholder="https://your-resource.openai.azure.com/", key="byok_url")
    else:
        byok_key = st.secrets.get("GITHUB_TOKEN", "")
        byok_model = "gpt-4o-mini"
        byok_url = "https://models.inference.ai.azure.com"

def get_openai_client(api_key: str, base_url: str):
    kwargs = {"api_key": api_key}
    if base_url and base_url.strip():
        kwargs["base_url"] = base_url.strip()
    return OpenAI(**kwargs)

# ==========================================
# 5. Word 處理核心模組 (含範本留存邏輯)
# ==========================================
def extract_text_from_docx(file):
    doc = docx.Document(file)
    content = []
    for p in doc.paragraphs:
        if p.text.strip():
            content.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            row_data = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            if any(row_data):
                content.append(" | ".join(row_data))
    return "\n".join(content)

def extract_text_from_pptx(file):
    prs = Presentation(file)
    content = []
    for idx, slide in enumerate(prs.slides, start=1):
        content.append(f"\n--- [ Slide {idx} ] ---")
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    if p.text.strip():
                        content.append(p.text.strip())
            if shape.has_table:
                for row in shape.table.rows:
                    row_data = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                    if any(row_data):
                        content.append(" | ".join(row_data))
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes_text = slide.notes_slide.notes_text_frame.text.strip()
            if notes_text:
                content.append(f"[備註/補充說明]: {notes_text}")
    return "\n".join(content)

# [新增] 複製儲存格格式 (Phase 2)
def copy_cell_formatting(src_cell, dst_cell):
    """精細複製儲存格的對齊與字體樣式"""
    if not src_cell.paragraphs:
        return
    src_para = src_cell.paragraphs[0]
    for dst_para in dst_cell.paragraphs:
        if src_para.alignment is not None:
            dst_para.alignment = src_para.alignment
        if src_para.runs and dst_para.runs:
            src_run = src_para.runs[0]
            dst_run = dst_para.runs[0]
            dst_run.font.name = src_run.font.name
            dst_run.font.size = src_run.font.size
            dst_run.font.bold = src_run.font.bold
            dst_run.font.italic = src_run.font.italic
            if src_run.font.color and src_run.font.color.rgb:
                dst_run.font.color.rgb = src_run.font.color.rgb

# [新增] 將 AI Markdown 填入用戶原檔 (Phase 2)
def fill_user_template(template_file, md_text: str) -> io.BytesIO:
    """保留用戶原檔 Logo、頁首尾，僅將資料注入表格"""
    template_file.seek(0)
    doc = Document(io.BytesIO(template_file.read()))

    # 1. 尋找原檔中第一個至少有 3 欄的表格
    target_table = None
    for table in doc.tables:
        if len(table.columns) >= 3:
            target_table = table
            break

    if target_table is None or len(target_table.rows) == 0:
        raise ValueError("無法在自訂範本中找到合適的會議記錄表格 (需至少 3 欄)。")

    header_row = target_table.rows[0]
    num_cols = len(target_table.columns)
    header_cells = [header_row.cells[i] for i in range(num_cols)]

    # 2. 刪除所有舊資料列 (保留標頭)
    while len(target_table.rows) > 1:
        tbl = target_table._tbl
        tr = target_table.rows[-1]._tr
        tbl.remove(tr)

    # 3. 解析 Markdown 表格
    lines = [ln.strip() for ln in md_text.splitlines()]
    clean_rows = []
    in_table = False
    for line in lines:
        if line.startswith("|") and line.endswith("|"):
            in_table = True
            if re.match(r'^\|[\s\:\-\|]+\|$', line): # 略過分隔線
                continue
            clean_rows.append([c.strip() for c in line.strip("|").split("|")])
        elif in_table and line == "":
            break

    if not clean_rows:
        raise ValueError("無法從 AI 回應中解析出有效的表格結構。")

    # 4. 逐行注入資料並繼承標頭樣式
    for md_row in clean_rows[1:]: # 跳過 Markdown 的 Header 行
        new_row = target_table.add_row()
        for col_idx in range(num_cols):
            cell = new_row.cells[col_idx]
            text = md_row[col_idx].replace('**', '').replace('<br>', '\n') if col_idx < len(md_row) else ""
            cell.text = text
            copy_cell_formatting(header_cells[col_idx], cell)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# [保留] 系統內建高質感渲染 (Fallback / 內建範本使用)
def convert_md_to_docx(md_text):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    lines = md_text.strip().split('\n')
    in_table = False
    table_data = []

    for line in lines:
        line_str = line.strip()
        if line_str.startswith('|') and line_str.endswith('|'):
            if re.match(r'^\|[\s\:\-\|]+\|$', line_str):
                continue
            cells = [c.strip() for c in line_str.strip('|').split('|')]
            table_data.append(cells)
            in_table = True
            continue
        else:
            if in_table and table_data:
                table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                table.style = 'Table Grid'
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                col_widths = [Inches(0.8), Inches(4.8), Inches(1.0)]

                for r_idx, row_cells in enumerate(table_data):
                    row = table.rows[r_idx]
                    trPr = row._tr.get_or_add_trPr()
                    trPr.append(docx.oxml.OxmlElement('w:cantSplit'))

                    for c_idx, cell_value in enumerate(row_cells):
                        cell = row.cells[c_idx]
                        if c_idx < len(col_widths):
                            cell.width = col_widths[c_idx]
                        clean_text = cell_value.replace('**', '').replace('<br>', '\n')
                        cell.text = clean_text
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                        p = cell.paragraphs[0]
                        p.paragraph_format.space_before = Pt(3)
                        p.paragraph_format.space_after = Pt(3)
                        p.paragraph_format.line_spacing = 1.15

                        if r_idx == 0:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for run in p.runs:
                                run.font.bold = True
                                run.font.size = Pt(10.5)
                                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                            shading = docx.oxml.parse_xml(r'<w:shd {} w:fill="1A365D"/>'.format(docx.oxml.ns.nsdecls('w')))
                            cell._tc.get_or_add_tcPr().append(shading)
                            trPr.append(docx.oxml.OxmlElement('w:tblHeader'))
                        else:
                            if c_idx == 0 or c_idx == len(row_cells) - 1:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for run in p.runs:
                                run.font.size = Pt(9.5)
                table_data = []
                in_table = False
                doc.add_paragraph()

        if not line_str:
            continue

        if line_str.startswith('# '):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(12)
            run = p.add_run(line_str.replace('# ', '').strip())
            run.font.size = Pt(18)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            parts = re.split(r'(\*\*.*?\*\*)', line_str)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    run = p.add_run(part[2:-2])
                    run.font.bold = True
                    run.font.size = Pt(10)
                else:
                    run = p.add_run(part)
                    run.font.size = Pt(10)

    if in_table and table_data:
        table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        col_widths = [Inches(0.8), Inches(4.8), Inches(1.0)]

        for r_idx, row_cells in enumerate(table_data):
            row = table.rows[r_idx]
            for c_idx, cell_value in enumerate(row_cells):
                cell = row.cells[c_idx]
                if c_idx < len(col_widths):
                    cell.width = col_widths[c_idx]
                clean_text = cell_value.replace('**', '').replace('<br>', '\n')
                cell.text = clean_text
                p = cell.paragraphs[0]
                if r_idx == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in p.runs:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    shading = docx.oxml.parse_xml(r'<w:shd {} w:fill="1A365D"/>'.format(docx.oxml.ns.nsdecls('w')))
                    cell._tc.get_or_add_tcPr().append(shading)
                else:
                    if c_idx == 0 or c_idx == len(row_cells) - 1:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ==========================================
# 6. 使用者輸入區域 (多範本選擇)
# ==========================================
st.subheader("📁 1. 選擇或上傳會議記錄格式範本")

template_option = st.radio(
    "請選擇會議記錄結構來源：",
    ["內建標準範本 (免上傳)", "自行上傳自訂 Word 範本 (.docx)"],
    horizontal=True
)

format_file_source = None

if template_option == "內建標準範本 (免上傳)":
    builtin_template = st.selectbox(
        "選擇內建商務範本類型：",
        [
            "通用團隊例會/週會範本 (Weekly / Team Meeting)",
            "高層/董事會決議型範本 (Board / Governance)",
            "專案跟進與檢討型範本 (Project / Deliverables)"
        ]
    )
    
    template_map = {
        "通用團隊例會/週會範本 (Weekly / Team Meeting)": ["template_weekly.docx", "templates/template_weekly.docx", "meeting_template.docx"],
        "高層/董事會決議型範本 (Board / Governance)": ["template_board.docx", "templates/template_board.docx"],
        "專案跟進與檢討型範本 (Project / Deliverables)": ["template_project.docx", "templates/template_project.docx"]
    }
    
    candidate_paths = template_map.get(builtin_template, [])
    for path in candidate_paths:
        if os.path.exists(path):
            format_file_source = path
            break
            
    if format_file_source:
        st.success(f"已成功載入內建範本：`{builtin_template}`")
    else:
        st.warning(f"⚠️ 系統尚未偵測到 `{builtin_template}` 的 `.docx` 檔案，請確認已上傳至 GitHub。")
else:
    format_file = st.file_uploader("請上傳作為結構基準的 Word 檔案 (.docx)", type=["docx"])
    if format_file:
        format_file_source = format_file
        st.success("成功載入自訂範本！系統將在匯出時保留您的專屬 Logo 與格式。")

st.markdown("---")

st.subheader("📊 2. 輸入本次會議內容來源 (選擇 A 或 選擇 B)")
col1, col2 = st.columns([1, 1])

with col1:
    current_ppt_file = st.file_uploader("選擇 A：上傳本次會議簡報 (.pptx)", type=["pptx"])

with col2:
    current_draft_text = st.text_area(
        "選擇 B：貼上會議草稿 (無 PPT 時使用)",
        height=180,
        placeholder="若無 PPT，請依照議題與編號輸入草稿內容，例如：\n1. 會議摘要\n1.1 通過上次會議紀錄。\n1.2 確認季度營運目標達成率為 95%。\n\n2. 資訊系統升級\n2.1 伺服器遷移計劃於下月第一週進行。"
    )

# ==========================================
# 7. AI 生成邏輯
# ==========================================
if st.button("🚀 即刻依範本結構生成會議記錄", type="primary", use_container_width=True):
    if not format_file_source:
        st.error("🛑 請務必選擇有效的內建範本或上傳自訂 Word 範本以建立會議骨架！")
    elif not current_ppt_file and not current_draft_text.strip():
        st.error("🛑 請提供本次會議內容，上傳 PPT 簡報（選擇 A）或輸入會議草稿（選擇 B）。")
    elif not byok_key:
        st.error("🛑 未檢測到有效的 API Key。請確認 Secrets 已設定 GITHUB_TOKEN，或開啟「自備 API Key」輸入密鑰。")
    else:
        with st.spinner("⏳ 正在進行 PII 本地遮蔽並分析範本歸納會議記錄..."):
            try:
                # 抽取文字
                if hasattr(format_file_source, 'read'):
                    format_file_source.seek(0)
                    format_structure_text = extract_text_from_docx(io.BytesIO(format_file_source.read()))
                else:
                    format_structure_text = extract_text_from_docx(format_file_source)
                    
                ppt_content_text = extract_text_from_pptx(current_ppt_file) if current_ppt_file else ""

                # PII 遮蔽
                masker = PIIMasker()
                masked_format = masker.mask(format_structure_text)
                masked_ppt = masker.mask(ppt_content_text)
                masked_draft = masker.mask(current_draft_text)

                client = get_openai_client(api_key=byok_key, base_url=byok_url)

                system_prompt = """
                你是一名精通企業行政與結構化合規管理的高級秘書。你的任務是進行「動態結構映射與內容提煉」。
                
                【核心運作邏輯】：
                1. 深度分析【格式範本/上次紀錄】，完全提取其內部使用的「編號」、「議題標題」與「表格結構」。這將作為本次會議紀錄的骨架。
                2. 對比【本次會議內容 (PPT 簡報或文字草稿)】，將內容歸納並填入範本對應的編號與議題下。
                3. 若上傳的是 PPT 簡報且某個頁面只有子題目而缺乏詳細內文，請直接將該子題目列入記錄即可。
                4. 如果某個議題在本次內容中完全沒有提及，請於該議題下寫「本次會議暫無相關事項。」，不可遺漏範本原有的任何一個大項。
                5. 文中若包含 [REDACTED_...] 格式的脫敏標籤，請完全原樣留存於對應位置，切勿刪除或修改。

                【排版與格式嚴格要求】：
                1. 頁首基本資料：使用純文字 + 粗體排版，嚴禁包含管道符號（`|`）。
                2. 表格規範 (Markdown Table)：
                   - 表格標頭必須嚴格包含 3 欄：`| 編號 | 議題 | 決議 |` (若範本為動議/專案表頭，請調整為對應 3 欄標題)
                   - 分隔線必須為：`| :--- | :--- | :---: |`
                   - 「決議」欄位請精簡輸出，只允許出現：（通過）、（記錄）或（跟進）這三個標籤，嚴禁把議題內文混在決議欄。
                   - 同議題下的子項目請分行呈現，確保每行的「決議」精確對齊第 3 欄。
                3. 語言：繁體中文（專業企業語彙）。
                """

                user_prompt = f"""
                請完全依照【格式範本】的編號與議題大項，將【本次會議內容】進行歸納填寫，產出全新的會議記錄：

                ===【1. 格式範本 / 上次紀錄 (定義本次會議的編號與議題骨架)】===
                {masked_format}

                ===【2. 本次會議內容來源】===
                【簡報檔提煉文字】：{masked_ppt}
                【文字草稿】：{masked_draft}
                """

                response = client.chat.completions.create(
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    model=byok_model,
                    temperature=0.2,
                )

                raw_md = response.choices[0].message.content
                final_minutes_md = masker.unmask(raw_md)

                st.session_state["generated_minutes"] = final_minutes_md
                st.success("✨ 本次會議記錄已成功生成 (完成本地 PII 去識別化與架構對齊)！")

            except Exception as e:
                st.error(f"❌ 生成失敗，請確認 API Key、Base URL 或資料內容是否正確: {str(e)}")

# ==========================================
# 8. 預覽與下載 (平滑回退邏輯)
# ==========================================
if st.session_state.get("generated_minutes"):
    st.markdown("---")
    st.subheader("📋 會議記錄預覽 (Preview)")
    st.markdown(st.session_state["generated_minutes"], unsafe_allow_html=True)
    st.markdown("---")
    
    # 決定匯出 Word 檔的處理方式
    try:
        # 如果用戶有上傳自訂範本，啟動 Phase 2 範本留存功能
        if hasattr(format_file_source, 'read'):
            word_buf = fill_user_template(format_file_source, st.session_state["generated_minutes"])
            download_label = "📄 下載標準 Word 記錄 (保留自訂範本 Logo 與格式)"
        else:
            word_buf = convert_md_to_docx(st.session_state["generated_minutes"])
            download_label = "📄 下載標準 Word 記錄 (內建商務格式)"
    except Exception as e:
        # 解析失敗平滑回退
        st.warning(f"⚠️ 自訂範本特殊格式套用失敗 ({e})，已自動切換至系統標準排版。")
        word_buf = convert_md_to_docx(st.session_state["generated_minutes"])
        download_label = "📄 下載標準 Word 記錄 (系統安全回退格式)"

    col_d1, col_d2 = st.columns([2, 1])
    
    with col_d1:
        st.download_button(
            label=download_label,
            data=word_buf,
            file_name="本次會議記錄.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
        
    with col_d2:
        md_buf = io.BytesIO(st.session_state["generated_minutes"].encode("utf-8"))
        st.download_button(
            label="📥 下載 Markdown 原始檔 (.md)",
            data=md_buf,
            file_name="本次會議記錄.md",
            mime="text/markdown",
            use_container_width=True
        )

