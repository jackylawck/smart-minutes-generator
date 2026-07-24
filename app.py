import os
import streamlit as st
from langchain_community.document_loaders import TextLoader, WebBaseLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_openai import ChatOpenAI
from langchain_community.vectorstores import Chroma
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.runnables import RunnablePassthrough
from langchain_core.output_parsers import StrOutputParser

# --- 1. 管治與 UI 護欄 (Governance & UI Guardrails) ---
st.set_page_config(page_title="AI 調解員詢問站", page_icon="⚖️")
st.title("⚖️ 香港新手調解員 AI 詢問站 (PoC)")

st.warning("""
**⚠️ 嚴格保密警告 (Confidentiality Notice)：**
根據《調解條例》（第620章）第8條、PD 31/31.1 及政府調解規則，調解通訊與會議內容具嚴格保密特權。
請**絕對不要**輸入任何真實案件的當事人姓名、公司名稱、財務條款或具體爭議細節。

*註：本系統僅作程序指引及學術參考，並不構成正式法律意見。*
""")

# --- 2. 獲取 Token ---
github_token = st.secrets.get("GITHUB_TOKEN") or st.secrets.get("OPENAI_API_KEY")
if not github_token:
    st.error("❌ 找不到 Token：請在 Streamlit App Settings -> Secrets 中設定 `GITHUB_TOKEN` 或 `OPENAI_API_KEY`。")
    st.stop()

# --- 3. 知識庫初始化 (Cap 620 + Cap 631 + 2025政府調解規則 + PD31/31.1 + 官方簡介 + 外部守則) ---
@st.cache_resource(show_spinner="正在重新建立中英雙語知識庫 (包含 HKMC 調解員守則)...")
def initialize_knowledge_base():
    documents = []
    
    # 策略 A: 載入本地所有 Markdown 文件
    local_files = [
        "data/Cap620.md", 
        "data/PD31.md", 
        "data/PD31_1.md",
        "data/Mediation_Intro.md",
        "data/Legal_Framework.md",
        "data/Gov_Mediation_Rules_2025.md"
    ]
    for file_path in local_files:
        try:
            loader = TextLoader(file_path, encoding="utf-8")
            documents.extend(loader.load())
        except Exception as e:
            st.toast(f"⚠️ 無法載入本地檔案 {file_path}: {e}")

    # 策略 B: 動態抓取外部官方與機構網頁 (包含 HKMC 兩條 Rules Link)
    urls = [
        "https://mediation.judiciary.hk/tc/mediation_faq.html",                       # 司法機構 FAQ (中文)
        "https://mediation.judiciary.hk/en/mediation_faq.html",                       # 司法機構 FAQ (英文)
        "https://www.hkmaal.org.hk/tc/HongKongMediationCode.php",                      # HKMAAL 守則
        "https://www.mediationcentre.org.hk/tc/services/MediationRules.php",          # HKMC 調解服務規則
        "https://www.mediationcentre.org.hk/en/mediators/Rules.php",                 # 新增：HKMC 調解員紀律與專業規則
        "https://hkiac.org/zh-hant/other-services/mediation/rules/hkiac-mediation-rules/", # HKIAC 中文
        "https://hkiac.org/other-services/mediation/rules/hkiac-mediation-rules/"            # HKIAC 英文
    ]
    for url in urls:
        try:
            web_loader = WebBaseLoader(url)
            documents.extend(web_loader.load())
        except Exception as e:
            st.toast(f"⚠️ 無法即時讀取外部網頁 ({url})。")

    if not documents:
        st.error("知識庫內容空白，請檢查 data/ 目錄下的 Markdown 檔案。")
        st.stop()

    # 優化 Chunking 與多語言向量搜尋
    text_splitter = RecursiveCharacterTextSplitter(chunk_size=600, chunk_overlap=150)
    splits = text_splitter.split_documents(documents)
    
    # 使用多語言 Embedding 模型
    embeddings = HuggingFaceEmbeddings(model_name="sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2")
    vectorstore = Chroma.from_documents(documents=splits, embedding=embeddings)
    
    return vectorstore.as_retriever(search_kwargs={"k": 5})

retriever = initialize_knowledge_base()

# --- 4. 串接 API 與 System Prompt ---
llm = ChatOpenAI(
    model="gpt-4o-mini",
    api_key=github_token,
    base_url="https://models.inference.ai.azure.com",
    temperature=0
)

system_prompt = (
    "你是一個專為香港新手調解員提供程序指引的 AI 助手。"
    "知識庫包含：香港法例第620章《調解條例》、第631章《道歉條例》、2025年《香港特別行政區政府調解規則》、實務指示 PD 31 / PD 31.1、司法機構 FAQ 及 HKMAAL/HKMC/HKIAC 守則與調解員紀律規則。\n\n"
    "請遵守以下嚴格管治規則：\n"
    "1. 仔細閱讀 Context（包含英文與中文文本）。如果 Context 包含答案，請務必精準回答，並用中文解釋。\n"
    "2. 只有在 Context 完全沒有提及相關話題時，才回答『根據現有知識庫，無法提供確切答案』。\n"
    "3. 必須明確引用資料來源（例如：『根據香港調解中心 (HKMC) 調解員守則』或『根據《政府調解規則 (2025年版)》』）。\n"
    "4. 如用戶輸入真實案件細節，請拒絕回答並提示保密條文。\n\n"
    "Context:\n{context}"
)

prompt = ChatPromptTemplate.from_messages([
    ("system", system_prompt),
    ("human", "{input}"),
])

def format_docs(docs):
    return "\n\n".join(doc.page_content for doc in docs)

rag_chain = (
    {"context": retriever | format_docs, "input": RunnablePassthrough()}
    | prompt
    | llm
    | StrOutputParser()
)

# --- 5. 對話介面 ---
if "messages" not in st.session_state:
    st.session_state.messages = []

for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

if user_query := st.chat_input("請輸入關於香港調解程序、政府2025調解規則、HKMC守則或保密條款的問題..."):
    st.session_state.messages.append({"role": "user", "content": user_query})
    with st.chat_message("user"):
        st.markdown(user_query)

    with st.chat_message("assistant"):
        with st.spinner("檢索知識庫中..."):
            try:
                answer = rag_chain.invoke(user_query)
                st.markdown(answer)
                st.session_state.messages.append({"role": "assistant", "content": answer})
            except Exception as e:
                st.error(f"系統生成回答時發生錯誤：{e}")
