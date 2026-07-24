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
# 1. 頁面配置與美化 CSS (完美支援深色/淺色模式)
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
    /* 指引卡片：採用半透明背景與動態文字，深淺色模式皆清晰 */
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
    
    /* 低調 Footer 聯絡資訊 */
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
# 2. 🛡️ 左側側邊欄：支援 3 大商務範本下載 & 低調個人 Hub 連結
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
    st.caption("🔒 **ISO 數據安全聲明：** 本系統採 Session-Only 記憶體即時運算，關閉網頁數據即刻徹底銷毀。")
    
    # 低調的個人 Hub 聯絡連結
    st.markdown("---")
    st.markdown(
        "<div class='footer-support'>💡 如有系統使用問題，歡迎聯絡 <a href='https://jackylawck.github.io/jackylawck/' target='_blank'>Jacky Law</a>。</div>", 
        unsafe_allow_html=True
    )

# ==========================================
# 3. 主畫面介面
# ==========================================
st.title("📝 智能會議記錄生成器 (Smart Minutes Generator)")
st.caption("基於動態模板映射與純本地端下載設計")

github_token = st.secrets.get("GITHUB_TOKEN", "")

# ==========================================
# 4. 檔案解析與 Word 轉換函式
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

def convert_md_to_docx(md_text):
    """將 Markdown 內容轉換為高質感、商業級排版的 Word (.docx) 文件"""
    doc = Document()
    
    # 設定頁面邊距 (2.54 cm 標準邊距)
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
        
        # 識別 Markdown 表格行
        if line_str.startswith('|') and line_str.endswith('|'):
            if re.match(r'^\|[\s\:\-\|]+\|$', line_str):
                continue
            cells = [c.strip() for c in line_str.strip('|').split('|')]
            table_data.append(cells)
            in_table = True
            continue
        else:
            # 渲染表格
            if in_table and table_data:
                table = doc.add_table(rows=len(table_data), cols=len(table_data[0]))
                table.style = 'Table Grid'
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                # 設定精確欄寬比例：編號 0.8 吋 | 議題 4.8 吋 | 決議 1.0 吋
                col_widths = [Inches(0.8), Inches(4.8), Inches(1.0)]

                for r_idx, row_cells in enumerate(table_data):
                    row = table.rows[r_idx]
                    
                    # 防止表格跨頁斷裂
                    trPr = row._tr.get_or_add_trPr()
                    trPr.append(docx.oxml.OxmlElement('w:cantSplit'))

                    for c_idx, cell_value in enumerate(row_cells):
                        cell = row.cells[c_idx]
                        if c_idx < len(col_widths):
                            cell.width = col_widths[c_idx]
                            
                        clean_text = cell_value.replace('**', '').replace('<br>', '\n')
                        cell.text = clean_text
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                        # 段落樣式與行距控制
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_before = Pt(3)
                        p.paragraph_format.space_after = Pt(3)
                        p.paragraph_format.line_spacing = 1.15

                        # 標頭列 (Header) 深藍底白字美化
                        if r_idx == 0:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for run in p.runs:
                                run.font.bold = True
                                run.font.size = Pt(10.5)
                                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF) # 白字
                            
                            # 背景色填滿：深藍色 (#1A365D)
                            shading = docx.oxml.parse_xml(r'<w:shd {} w:fill="1A365D"/>'.format(docx.oxml.ns.nsdecls('w')))
                            cell._tc.get_or_add_tcPr().append(shading)
                            
                            # 跨頁自動重複標頭
                            trPr.append(docx.oxml.OxmlElement('w:tblHeader'))
                        else:
                            # 編號與決議欄位居中對齊
                            if c_idx == 0 or c_idx == len(row_cells) - 1:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for run in p.runs:
                                run.font.size = Pt(9.5)

                table_data = []
                in_table = False
                doc.add_paragraph() # 表格後空行

        if not line_str:
            continue

        # 頁頭大標題
        if line_str.startswith('# '):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(12)
            run = p.add_run(line_str.replace('# ', '').strip())
            run.font.size = Pt(18)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
        # 基本資料段落 (日期、時間、出席人等)
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

    # 處理尾端表格數據
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
# 5. 使用者輸入區域 (多範本選擇)
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

st.markdown("---")

st.subheader("📊 2. 輸入本次會議內容來源 (選擇 A 或 選擇 B)")
col1, col2 = st.columns([1, 1])

with col1:
    current_ppt_file = st.file_uploader("選擇 A：上傳本次會議簡報 (.pptx)", type=["pptx"])

with col2:
    current_draft_text = st.text_area(
        "選擇 B：貼上會議草稿 (無 PPT 時使用)",
        height=180,
        placeholder="若無 PPT，請依照議題與編號輸入草稿內容，例如：\n1. 會議摘要\n1.1 通過上次會議紀錄。\n1.2 確認季度營運目標達成率為 95%。\n\n2. 資訊系統升級\n2.1 伺服器遷移計劃於下月第一週進行。\n2.2 預算審批授權予 IT 部門經理跟進。\n\n3. 辦公室行政\n3.1 採購新批次辦公設備。"
    )

# ==========================================
# 6. AI 動態對齊生成邏輯
# ==========================================
if st.button("🚀 即刻依範本結構生成會議記錄", type="primary", use_container_width=True):
    if not format_file_source:
        st.error("🛑 請務必選擇有效的內建範本或上傳自訂 Word 範本以建立會議骨架！")
    elif not current_ppt_file and not current_draft_text.strip():
        st.error("🛑 請提供本次會議內容，上傳 PPT 簡報（選擇 A）或輸入會議草稿（選擇 B）。")
    elif not github_token:
        st.error("🛑 系統未偵測到 GITHUB_TOKEN Secrets，請檢查 Streamlit Cloud 後台設定。")
    else:
        with st.spinner("⏳ 正在分析範本骨架並歸納會議記錄..."):
            try:
                format_structure_text = extract_text_from_docx(format_file_source)
                ppt_content_text = extract_text_from_pptx(current_ppt_file) if current_ppt_file else ""

                client = OpenAI(
                    base_url="https://models.inference.ai.azure.com",
                    api_key=github_token,
                )

                system_prompt = """
                你是一名精通企業行政與結構化合規管理的高級秘書。你的任務是進行「動態結構映射與內容提煉」。
                
                【核心運作邏輯】：
                1. 深度分析【格式範本/上次紀錄】，完全提取其內部使用的「編號」、「議題標題」與「表格結構」。這將作為本次會議紀錄的骨架。
                2. 對比【本次會議內容 (PPT 簡報或文字草稿)】，將內容歸納並填入範本對應的編號與議題下。
                3. 若上傳的是 PPT 簡報且某個頁面只有子題目而缺乏詳細內文，請直接將該子題目列入記錄即可。
                4. 如果某個議題在本次內容中完全沒有提及，請於該議題下寫「本次會議暫無相關事項。」，不可遺漏範本原有的任何一個大項。

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
                {format_structure_text}

                ===【2. 本次會議內容來源】===
                【簡報檔提煉文字】：{ppt_content_text}
                【文字草稿】：{current_draft_text}
                """

                response = client.chat.completions.create(
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    model="gpt-4o-mini",
                    temperature=0.2,
                )

                minutes_md = response.choices[0].message.content
                st.session_state["generated_minutes"] = minutes_md
                st.success("✨ 本次會議記錄已根據範本結構成功生成！")

            except Exception as e:
                st.error(f"❌ 生成失敗，請確認資料內容或 Token 設定: {str(e)}")

# ==========================================
# 7. 本地端純下載預覽 (預設提供 Word .docx 下載)
# ==========================================
if "generated_minutes" in st.session_state:
    st.markdown("---")
    st.subheader("📋 會議記錄預覽 (Preview)")
    st.markdown(st.session_state["generated_minutes"], unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 即時生成高質感 Word (.docx) 檔案
    docx_file = convert_md_to_docx(st.session_state["generated_minutes"])
    
    col_d1, col_d2 = st.columns([2, 1])
    
    with col_d1:
        st.download_button(
            label="📄 即刻下載標準 Word 格式會議記錄 (.docx)",
            data=docx_file,
            file_name="本次會議記錄.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
        
    with col_d2:
        st.download_button(
            label="📥 下載 Markdown 原始檔 (.md)",
            data=st.session_state["generated_minutes"],
            file_name="本次會議記錄.md",
            mime="text/markdown",
            use_container_width=True
        )
