import os
import streamlit as st
import docx
from pptx import Presentation
from openai import OpenAI

# ==========================================
# 1. 頁面配置與美化 CSS
# ==========================================
st.set_page_config(
    page_title="通用智能會議記錄生成器",
    page_icon="📝",
    layout="wide"
)

st.markdown("""
    <style>
    .sidebar-title {
        color: #1a365d;
        font-weight: bold;
        font-size: 1.2em;
        margin-bottom: 10px;
    }
    .guideline-card {
        background-color: #ffffff;
        padding: 12px;
        border-radius: 6px;
        border: 1px solid #e0e0e0;
        border-left: 4px solid #0056b3;
        margin-bottom: 12px;
        font-size: 0.88em;
        line-height: 1.5;
    }
    table {
        width: 100% !important;
        border-collapse: collapse;
    }
    th:nth-child(1), td:nth-child(1) { width: 8% !important; text-align: center !important; }
    th:nth-child(2), td:nth-child(2) { width: 77% !important; }
    th:nth-child(3), td:nth-child(3) { width: 15% !important; text-align: center !important; white-space: nowrap !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 🛡️ 左側側邊欄：標準 PPT 與草稿模式指引
# ==========================================
with st.sidebar:
    st.markdown("<div class='sidebar-title'>📖 使用方式與格式指引</div>", unsafe_allow_html=True)
    
    st.markdown("""
    本系統支援兩種會議內容輸入模式（二選一）：
    """)
    
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

# ==========================================
# 3. 主畫面介面
# ==========================================
st.title("📝 通用智能會議記錄生成器 (Smart Minutes Generator)")
st.caption("基於動態模板映射與純本地端下載設計")

github_token = st.secrets.get("GITHUB_TOKEN", "")

# ==========================================
# 4. 檔案解析函式 (Docx & Pptx)
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

# ==========================================
# 5. 使用者輸入區域 (通用商業/行政範例)
# ==========================================
st.subheader("📁 1. 上傳會議記錄格式範本 / 上次紀錄")
format_file = st.file_uploader("請上傳作為結構基準的 Word 檔案 (.docx)", type=["docx"])

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
    if not format_file:
        st.error("🛑 請務必上傳「1. 格式範本 / 上次紀錄 (.docx)」以建立會議骨架！")
    elif not current_ppt_file and not current_draft_text.strip():
        st.error("🛑 請提供本次會議內容，上傳 PPT 簡報（選擇 A）或輸入會議草稿（選擇 B）。")
    elif not github_token:
        st.error("🛑 系統未偵測到 GITHUB_TOKEN Secrets，請檢查 Streamlit Cloud 後台設定。")
    else:
        with st.spinner("⏳ 正在分析範本骨架並歸納會議記錄..."):
            try:
                format_structure_text = extract_text_from_docx(format_file)
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
                   - 表格標頭必須嚴格包含 3 欄：`| 編號 | 議題 | 決議 |`
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
# 7. 本地端純下載預覽
# ==========================================
if "generated_minutes" in st.session_state:
    st.markdown("---")
    st.subheader("📋 會議記錄預覽 (Preview)")
    st.markdown(st.session_state["generated_minutes"], unsafe_allow_html=True)
    
    st.markdown("---")
    st.download_button(
        label="📥 即刻下載通用格式會議記錄 (.md)",
        data=st.session_state["generated_minutes"],
        file_name="本次會議記錄.md",
        mime="text/markdown",
        type="primary",
        use_container_width=True
    )
