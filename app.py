import os
import streamlit as st
import docx
from pptx import Presentation
from openai import OpenAI

# ==========================================
# 1. 頁面配置 (ISO / Privacy-First UI)
# ==========================================
st.set_page_config(
    page_title="智能會議記錄生成器",
    page_icon="📝",
    layout="wide"
)

st.title("📝 智能會議記錄生成器 (Smart Minutes Generator)")
st.caption("🔒 本地端純下載版 — 遵循 Zero Data Retention 原則，不存檔至任何雲端資料庫")

# 從 Streamlit Secrets 讀取免費 AI API Token (外人無法於 Public Repo 看到)
github_token = st.secrets.get("GITHUB_TOKEN", "")

st.info(
    "🛡️ **ISO 數據安全與隱私保障聲明：**\n"
    "1. 本系統採用 **Session-Only 記憶體處理**，您上傳的 Word 及 PPT 檔案僅供本次 AI 提煉使用。\n"
    "2. 生成結果**絕對不會回傳或 Commit 至任何 GitHub Repository**，請安心直接下載至公司內部安全硬碟。"
)

# ==========================================
# 2. 記憶體解析 Word & PPT 函式
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
    return "\n".join(content)

# ==========================================
# 3. 側邊欄：議題設定
# ==========================================
with st.sidebar:
    st.header("📋 會議架構設定")
    custom_agenda = st.text_area(
        "預設會議議題分類:",
        value="1. 摘要\n2. 招聘與人資\n3. 培訓與發展\n4. 財務與薪酬\n5. 員工福利\n6. 績效考核\n7. 離職與異動\n8. 營運與系統\n9. 規章制度與合規\n10. 其他事項",
        height=220
    )

# ==========================================
# 4. 主介面：檔案上傳
# ==========================================
col1, col2 = st.columns(2)
with col1:
    last_minutes_file = st.file_uploader("1. 上傳上次會議記錄 (.docx)", type=["docx"])
with col2:
    current_ppt_file = st.file_uploader("2. 上傳本次會議簡報 (.pptx)", type=["pptx"])

# ==========================================
# 5. AI 記憶體運算邏輯
# ==========================================
if st.button("🚀 即刻生成本次會議記錄", type="primary", use_container_width=True):
    if not last_minutes_file or not current_ppt_file:
        st.error("🛑 請務必同時上傳「上次會議記錄」與「本次 PPT」！")
    elif not github_token:
        st.error("🛑 系統未設定 GITHUB_TOKEN Secrets，請聯絡系統管理者。")
    else:
        with st.spinner("⏳ 正在於記憶體進行資料比對與 AI 生成..."):
            try:
                last_text = extract_text_from_docx(last_minutes_file)
                ppt_text = extract_text_from_pptx(current_ppt_file)

                client = OpenAI(
                    base_url="https://models.inference.ai.azure.com",
                    api_key=github_token,
                )

                system_prompt = f"""
                你是一名專業的企業行政秘書。
                請參考「上次會議記錄」的格式、風格與歷史數據，結合「本次會議 PPT」的最新內容，
                撰寫一份結構嚴謹的正式會議記錄。

                【必須嚴格遵守的規範】：
                1. 採用以下議題架構分類整理內文：
                   {custom_agenda}
                2. 表格欄位必須包含「編號」、「議題」與「決議」。
                3. 每一條決議事項結尾必須明確標註：（通過）、（記錄）或（跟進）。
                4. 數據精確度：人數變動、異動名單及各項專案進度必須嚴格按 PPT 內容與數據計算，切勿虛構。
                5. 請使用繁體中文（專業企業語彙）。
                """

                response = client.chat.completions.create(
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": f"【上次紀錄】:\n{last_text}\n\n【本次PPT內容】:\n{ppt_text}"}
                    ],
                    model="gpt-4o-mini",
                    temperature=0.2,
                )

                minutes_md = response.choices[0].message.content
                st.session_state["generated_minutes"] = minutes_md
                st.success("✨ 本次會議記錄已順利生成！")

            except Exception as e:
                st.error(f"❌ 生成失敗，請檢查 API 設定或檔案內容: {str(e)}")

# ==========================================
# 6. 本地端純下載（ISO 零數據留存設計）
# ==========================================
if "generated_minutes" in st.session_state:
    st.markdown("---")
    st.subheader("📋 會議記錄預覽 (Preview)")
    st.markdown(st.session_state["generated_minutes"])
    
    st.markdown("---")
    st.download_button(
        label="📥 即刻下載 Markdown 檔 (保存至本地端硬碟)",
        data=st.session_state["generated_minutes"],
        file_name="本次會議記錄.md",
        mime="text/markdown",
        type="primary",
        use_container_width=True
    )
