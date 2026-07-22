import os
import Streamlit as st
import docx
from pptx import Presentation
from openai import OpenAI

# ==========================================
# 1. 頁面配置與美化 CSS
# ==========================================
st.set_page_config(
    Page_title="智能會議記錄生成器",
    Page_icon="📝",
    Layout="wide"
)

st.markdown("""
    <style>
    .sidebar-title {
        Color: #1a365d;
        Font-weight: bold;
        Font-size: 1.2em;
        Margin-bottom: 10px;
    }
    .guideline-card {
        Background-color: #ffffff;
        Padding: 12px;
        Border-radius: 6px;
        Border: 1px solid #e0e0e0;
        Border-left: 4px solid #0056b3;
        Margin-bottom: 12px;
        Font-size: 0.88em;
        Line-height: 1.5;
    }
    Table {
        Width: 100% !important;
        Border-collapse: collapse;
    }
    Th:nth-child(1), td:nth-child(1) { width: 8% !important; text-align: center !important; }
    Th:nth-child(2), td:nth-child(2) { width: 77% !important; }
    Th:nth-child(3), td:nth-child(3) { width: 15% !important; text-align: center !important; white-space: nowrap !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 🛡️ 左側側邊欄：標準下載、PPT 與草稿模式指引
# ==========================================
With st.sidebar:
    St.markdown("<div class='sidebar-title'>📥 下載基準範本</div>", unsafe_allow_html=True)
    
    # 嘗試讀取預設範本 meeting_template.docx 提供下載
    Default_template_path = "meeting_template.docx"
    If os.path.exists(default_template_path):
        With open(default_template_path, "rb") as f:
            St.download_button(
                Label="📄 下載空白週會例會範本 (.docx)",
                Data=f,
                File_name="週會例會會議記錄標準範本.docx",
                Mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                Use_container_width=True
            )
    Else:
        St.info("💡 若需於此處提供預設下載，請將 meeting_template.docx 上傳至 GitHub 根目錄。")

    St.markdown("---")
    St.markdown("<div class='sidebar-title'>📖 使用方式與格式指引</div>", unsafe_allow_html=True)
    St.markdown("本系統支援兩種會議內容輸入模式（二選一）：")
    
    St.markdown("""
    <div class='guideline-card'>
    <b>📊 模式 A：上傳 PPT 簡報</b><br>
    • <b>排版建議：</b>頁面頂部為大議題，中間為子題目，下方為詳細文字。<br>
    • <b>顯示邏輯：</b>系統會自動將內容填入範本結構；若 PPT 頁面僅有子題目而無詳細文字，紀錄中將直接顯示該子題目。
    </div>
    """, unsafe_allow_html=True)
    
    St.markdown("""
    <div class='guideline-card'>
    <b>📝 模式 B：文字草稿輸入 (無 PPT 時使用)</b><br>
    • <b>使用情境：</b>若無 PPT 簡報，可直接在草稿框按層級格式（如：大議題 -> 1.1 ... 1.2 ...）寫下文字草稿。<br>
    • <b>處理邏輯：</b>AI 會將草稿內容精準歸納並整理至對應的議題架構中。
    </div>
    """, unsafe_allow_html=True)

    St.markdown("---")
    St.caption("🔒 **ISO 數據安全聲明：** 本系統採 Session-Only 記憶體即時運算，關閉網頁數據即刻徹底銷毀。")

# ==========================================
# 3. 主畫面介面
# ==========================================
St.title("📝 智能會議記錄生成器 (Smart Minutes Generator)")
St.caption("基於動態模板映射與純本地端下載設計")

Github_token = st.secrets.get("GITHUB_TOKEN", "")

# ==========================================
# 4. 檔案解析函式 (Docx & Pptx)
# ==========================================
Def extract_text_from_docx(file):
    # 支援傳入檔案路徑或 UploadedFile 物件
    Doc = docx.Document(file)
    Content = []
    For p in doc.paragraphs:
        If p.text.strip():
            Content.append(p.text.strip())
    For table in doc.tables:
        For row in table.rows:
            Row_data = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            If any(row_data):
                Content.append(" | ".join(row_data))
    Return "\n".join(content)

Def extract_text_from_pptx(file):
    Prs = Presentation(file)
    Content = []
    For idx, slide in enumerate(prs.slides, start=1):
        Content.append(f"\n--- [ Slide {idx} ] ---")
        For shape in slide.shapes:
            If shape.has_text_frame:
                For p in shape.text_frame.paragraphs:
                    If p.text.strip():
                        Content.append(p.text.strip())
            If shape.has_table:
                For row in shape.table.rows:
                    Row_data = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                    If any(row_data):
                        Content.append(" | ".join(row_data))
                        
        If slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            Notes_text = slide.notes_slide.notes_text_frame.text.strip()
            If notes_text:
                Content.append(f"[備註/補充說明]: {notes_text}")
                
    Return "\n".join(content)

# ==========================================
# 5. 使用者輸入區域 (Phase 1 升級：多範本選擇)
# ==========================================
St.subheader("📁 1. 選擇或上傳會議記錄格式範本")

Template_option = st.radio(
    "請選擇會議記錄結構來源：",
    ["內建標準範本 (免上傳)", "自行上傳自訂 Word 範本 (.docx)"],
    Horizontal=True
)

Format_file_source = None

If template_option == "內建標準範本 (免上傳)":
    Builtin_template = st.selectbox(
        "選擇內建商務範本類型：",
        [
            "通用團隊例會/週會範本 (Weekly / Team Meeting)",
            "高層/董事會決議型範本 (Board / Governance)",
            "專案跟進與檢討型範本 (Project / Deliverables)"
        ]
    )
    
    # 映射可能存在的檔案路徑（相容子目錄 templates/ 或根目錄）
    Template_map = {
        "通用團隊例會/週會範本 (Weekly / Team Meeting)": ["templates/template_weekly.docx", "template_weekly.docx", "meeting_template.docx"],
        "高層/董事會決議型範本 (Board / Governance)": ["templates/template_board.docx", "template_board.docx"],
        "專案跟进與檢討型範本 (Project / Deliverables)": ["templates/template_project.docx", "template_project.docx"]
    }
    
    Candidate_paths = template_map.get(builtin_template, [])
    For path in candidate_paths:
        If os.path.exists(path):
            Format_file_source = path
            Break
            
    If format_file_source:
        St.success(f"已成功載入內建範本：`{builtin_template}`")
    Else:
        St.warning(f"⚠️ 系統尚未偵測到 `{builtin_template}` 的 `.docx` 檔案，請確認已上傳至 GitHub。")

Else:
    Format_file = st.file_uploader("請上傳作為結構基準的 Word 檔案 (.docx)", type=["docx"])
    If format_file:
        Format_file_source = format_file

St.markdown("---")

St.subheader("📊 2. 輸入本次會議內容來源 (選擇 A 或 選擇 B)")
Col1, col2 = st.columns([1, 1])

With col1:
    Current_ppt_file = st.file_uploader("選擇 A：上傳本次會議簡報 (.pptx)", type=["pptx"])

With col2:
    Current_draft_text = st.text_area(
        "選擇 B：貼上會議草稿 (無 PPT 時使用)",
        Height=180,
        Placeholder="若無 PPT，請依照議題與編號輸入草稿內容，例如：\n1. 會議摘要\n1.1 通過上次會議紀錄。\n1.2 確認季度營運目標達成率為 95%。\n\n2. 資訊系統升級\n2.1 伺服器遷移計劃於下月第一週進行。\n2.2 預算審批授權予 IT 部門經理跟進。\n\n3. 辦公室行政\n3.1 採購新批次辦公設備。"
    )

# ==========================================
# 6. AI 動態對齊生成邏輯
# ==========================================
If st.button("🚀 即刻依範本結構生成會議記錄", type="primary", use_container_width=True):
    If not format_file_source:
        St.error("🛑 請務必選擇有效的內建範本或上傳自訂 Word 範本以建立會議骨架！")
    Elif not current_ppt_file and not current_draft_text.strip():
        St.error("🛑 請提供本次會議內容，上傳 PPT 簡報（選擇 A）或輸入會議草稿（選擇 B）。")
    Elif not github_token:
        St.error("🛑 系統未偵測到 GITHUB_TOKEN Secrets，請檢查 Streamlit Cloud 後台設定。")
    Else:
        With st.spinner("⏳ 正在分析範本骨架並歸納會議記錄..."):
            Try:
                Format_structure_text = extract_text_from_docx(format_file_source)
                Ppt_content_text = extract_text_from_pptx(current_ppt_file) if current_ppt_file else ""

                Client = OpenAI(
                    Base_url="https://models.inference.ai.azure.com",
                    Api_key=github_token,
                )

                System_prompt = """
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

                User_prompt = f"""
                請完全依照【格式範本】的編號與議題大項，將【本次會議內容】進行歸納填寫，產出全新的會議記錄：

                ===【1. 格式範本 / 上次紀錄 (定義本次會議的編號與議題骨架)】===
                {format_structure_text}

                ===【2. 本次會議內容來源】===
                【簡報檔提煉文字】：{ppt_content_text}
                【文字草稿】：{current_draft_text}
                """

                Response = client.chat.completions.create(
                    Messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    Model="gpt-4o-mini",
                    Temperature=0.2,
                )

                Minutes_md = response.choices[0].message.content
                St.session_state["generated_minutes"] = minutes_md
                St.success("✨ 本次會議記錄已根據範本結構成功生成！")

            Except Exception as e:
                St.error(f"❌ 生成失敗，請確認資料內容或 Token 設定: {str(e)}")

# ==========================================
# 7. 本地端純下載預覽
# ==========================================
If "generated_minutes" in st.session_state:
    St.markdown("---")
    St.subheader("📋 會議記錄預覽 (Preview)")
    St.markdown(st.session_state["generated_minutes"], unsafe_allow_html=True)
    
    St.markdown("---")
    St.download_button(
        Label="📥 即刻下載通用格式會議記錄 (.md)",
        Data=st.session_state["generated_minutes"],
        File_name="本次會議記錄.md",
        Mime="text/markdown",
        Type="primary",
        Use_container_width=True
    )
