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
    .warning-card {
        background-color: #fff3cd;
        padding: 10px;
        border-radius: 6px;
        border-left: 4px solid #ffc107;
        margin-bottom: 12px;
        font-size: 0.85em;
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
# 2. 🛡️ 左側側邊欄：標準 PPT 格式與輸入指引
# ==========================================
with st.sidebar:
    st.markdown("<div class='sidebar-title'>📖 PPT 格式與輸入指引</div>", unsafe_allow_html=True)
    
    st.markdown("""
    為了讓 AI 能 100% 精準擷取會議內容並生成完整紀錄，請確保上傳的 **PPT 頁面符合以下排版格式**：
    """)
    
    st.markdown("""
    <div class='guideline-card'>
    <b>📌 PPT 標準結構排版：</b><br>
    • <b>頂部：</b>主議題 (對應範本大標題，如「三、培訓」)<br>
    • <b>中間：</b>子題目 (對應項目編號，如「4. 培訓資助」)<br>
    • <b>最下面：</b>詳細文字內容 (具體決議、名單或數據)
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class='warning-card'>
    ⚠️ <b>重要注意事項：</b><br>
    1. <b>必須使用純文字框：</b> 最下方的詳細文字<b>嚴禁使用圖片或電郵截圖</b>！若是圖片，系統無法讀取文字，會議紀錄將<b>只會顯示子題目</b>。<br>
    2. <b>無文字時補救：</b> 若 PPT 最下方未寫文字，請於右側「草稿 / 補充說明框」手動補上文字內容。
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
# 4. 檔案解析函式 (Docx & Pptx + Notes)
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
# 5. 使用者輸入區域
# ==========================================
st.subheader("📁 1. 上傳會議記錄格式範本 / 上次紀錄")
format_file = st.file_uploader("請上傳作為結構基準的 Word 檔案 (.docx)", type=["docx"])

st.markdown("---")

st.subheader("📊 2. 輸入本次會議內容來源")
col1, col2 = st.columns([1, 1])

with col1:
    current_ppt_file = st.file_uploader("選擇 A：上傳本次會議簡報 (.pptx)", type=["pptx"])

with col2:
    current_draft_text = st.text_area(
        "選擇 B：補充說明 / 若 PPT 最下方無文字請在此輸入",
        height=160,
        placeholder="若 PPT 頁面包含圖片/截圖，或最下方缺少詳細文字，請在此補上重點。例如：\n3.5 培訓資助：建議全數資助朱宗亮及李孝慈（Jason）報讀第二科《面對轉變》。"
    )

# ==========================================
# 6. AI 動態對齊生成邏輯
# ==========================================
if st.button("🚀 即刻依範本結構生成會議記錄", type="primary", use_container_width=True):
    if not format_file:
        st.error("🛑 請務必上傳「1. 格式範本 / 上次紀錄 (.docx)」以建立會議骨架！")
    elif not current_ppt_file and not current_draft_text.strip():
        st.error("🛑 請提供本次會議內容，上傳 PPT 簡報或在右側文字框輸入紀錄。")
    elif not github_token:
        st.error("🛑 系統未偵測到 GITHUB_TOKEN Secrets，請檢查 Streamlit Cloud 後台設定。")
    else:
        with st.spinner("⏳ 正在分析範本骨架並動態映射內容..."):
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
                2. 對比【本次會議內容 (包含 PPT 文字、PPT 備註欄及文字草稿)】，尋找與範本議題名稱或編號對應的最新動態、決議與數據。
                3. 將新內容精準填入對應的編號與議題下面。如果某個議題在本次簡報、備註或草稿中完全沒有提及，請於該議題下寫「本次會議暫無相關事項。」，不可遺漏範本原有的任何一個議題大項。

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
                【簡報檔及備註欄提煉文字】：{ppt_content_text}
                【文字草稿/補充說明】：{current_draft_text}
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
