import streamlit as st
import re
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pptx import Presentation
from openai import OpenAI

st.set_page_config(page_title="Smart Minutes Generator", layout="wide")

# ─────────────────────────────────────────
# BYOK (Bring Your Own Key) — Session State
# ─────────────────────────────────────────
if "use_byok" not in st.session_state:
    st.session_state["use_byok"] = False
if "byok_provider" not in st.session_state:
    st.session_state["byok_provider"] = "GitHub Models"
if "byok_key" not in st.session_state:
    st.session_state["byok_key"] = ""
if "byok_model" not in st.session_state:
    st.session_state["byok_model"] = "gpt-4o"
if "byok_url" not in st.session_state:
    st.session_state["byok_url"] = "https://models.inference.ai.azure.com"

# ─────────────────────────────────────────
# PII Masker (ISO 27701 Pseudonymization)
# ─────────────────────────────────────────
class PIIMasker:
    PATTERNS = {
        'HKID':    r'\b[A-Z]{1,2}\d{6}[$]?\d?[$]?[A-Z]?\b',
        'TW_ID':   r'\b[A-Z][12]\d{8}\b',
        'CN_ID':   r'\b\d{17}[\dXx]\b',
        'EMAIL':   r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}',
        'PHONE':   r'(?:\+?852[\s\-]?)?(?:\d{4}[\s\-]?\d{4}|\d{8})',
        'SALARY':  r'(?:薪資|薪|月薪|年薪|salary)[\s：:]*(?:HK\$|NT\$|¥|\$)?\s*[\d,]+(?:K|k|萬)?',
    }

    def __init__(self):
        self.vault = {}
        self.counters = {k: 0 for k in self.PATTERNS}

    def mask(self, text: str) -> str:
        self.vault.clear()
        self.counters = {k: 0 for k in self.PATTERNS}
        masked = text
        for pii_type, pattern in self.PATTERNS.items():
            def replace(m):
                self.counters[pii_type] += 1
                token = f"[REDACTED_{pii_type}_{self.counters[pii_type]}]"
                self.vault[token] = m.group(0)
                return token
            masked = re.sub(pattern, replace, masked)
        return masked

    def unmask(self, text: str) -> str:
        for token, original in self.vault.items():
            text = text.replace(token, original)
        return text

# ─────────────────────────────────────────
# OpenAI Client Factory
# ─────────────────────────────────────────
def get_openai_client(api_key: str, base_url: str):
    kwargs = {"api_key": api_key}
    if base_url and base_url.strip():
        kwargs["base_url"] = base_url.strip()
    return OpenAI(**kwargs)

# ─────────────────────────────────────────
# Markdown Table Parser
# ─────────────────────────────────────────
def parse_md_table(md_text: str):
    """
    Parse a markdown table from text.
    Returns list of lists (rows) or None if no table found.
    """
    lines = [ln.strip() for ln in md_text.splitlines()]
    table_lines = []
    in_table = False
    for line in lines:
        if line.startswith("|") and line.endswith("|"):
            table_lines.append(line)
            in_table = True
        elif in_table and line == "":
            break
        elif in_table:
            break
    if len(table_lines) < 2:
        return None
    # Remove separator line (like |---|---|)
    clean_rows = []
    for tl in table_lines:
        if re.match(r'\|[\s\-:|]+\|', tl):
            continue
        cells = [c.strip() for c in tl.strip("|").split("|")]
        clean_rows.append(cells)
    return clean_rows if clean_rows else None

def parse_md_to_rows(md_text: str):
    """
    Try to extract rows from markdown.
    If a table is found, return its rows.
    Otherwise, treat each non-empty line as a single-column row.
    """
    rows = parse_md_table(md_text)
    if rows:
        return rows
    # Fallback: each line is one cell in one column
    return [[ln.strip()] for ln in md_text.splitlines() if ln.strip()]

# ─────────────────────────────────────────
# Helper: Copy cell formatting
# ─────────────────────────────────────────
def copy_cell_formatting(src_cell, dst_cell):
    """Copy paragraph alignment, font name/size/color/bold/italic from src to dst."""
    if not src_cell.paragraphs:
        return
    src_para = src_cell.paragraphs[0]
    for dst_para in dst_cell.paragraphs:
        # Alignment
        if src_para.alignment is not None:
            dst_para.alignment = src_para.alignment
        # Font
        if src_para.runs and dst_para.runs:
            src_run = src_para.runs[0]
            dst_run = dst_para.runs[0]
            dst_run.font.name = src_run.font.name
            dst_run.font.size = src_run.font.size
            dst_run.font.bold = src_run.font.bold
            dst_run.font.italic = src_run.font.italic
            if src_run.font.color and src_run.font.color.rgb:
                dst_run.font.color.rgb = src_run.font.color.rgb
        # Copy paragraph style
        if src_para.style:
            try:
                dst_para.style = src_para.style
            except Exception:
                pass

# ─────────────────────────────────────────
# Phase 2 Feature: Fill User Template
# ─────────────────────────────────────────
def fill_user_template(template_file, md_text: str) -> io.BytesIO:
    """
    1. Load the user-uploaded Word template (keeps logos, watermarks, headers/footers, fonts).
    2. Find the first table with >= 3 columns.
    3. Keep the header row (row 0) styling.
    4. Delete old data rows.
    5. Fill with parsed Markdown rows, inheriting header cell formatting.
    6. Return a BytesIO buffer.
    """
    doc = Document(template_file)

    target_table = None
    for table in doc.tables:
        if len(table.columns) >= 3:
            target_table = table
            break

    if target_table is None:
        raise ValueError("Template does not contain a table with at least 3 columns.")

    if len(target_table.rows) == 0:
        raise ValueError("Target table has no rows.")

    header_row = target_table.rows[0]
    num_cols = len(target_table.columns)

    # Store header cell prototypes for formatting
    header_cells = [header_row.cells[i] for i in range(num_cols)]

    # Delete all rows except header (from bottom to top to keep indices valid)
    while len(target_table.rows) > 1:
        tbl = target_table._tbl
        tr = target_table.rows[-1]._tr
        tbl.remove(tr)

    # Parse markdown content into rows
    md_rows = parse_md_to_rows(md_text)

    # Add new rows and fill cells
    for md_row in md_rows:
        new_row = target_table.add_row()
        for col_idx in range(num_cols):
            cell = new_row.cells[col_idx]
            # Text: use md cell if available, else empty
            text = md_row[col_idx] if col_idx < len(md_row) else ""
            cell.text = text
            # Apply header formatting
            copy_cell_formatting(header_cells[col_idx], cell)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────
# Legacy: Convert Markdown to DOCX (Fallback)
# ─────────────────────────────────────────
def convert_md_to_docx(md_text: str, template_choice: str) -> io.BytesIO:
    doc = Document()

    # Page margins
    sections = doc.sections[0]
    sections.top_margin = Inches(1)
    sections.bottom_margin = Inches(1)
    sections.left_margin = Inches(1)
    sections.right_margin = Inches(1)

    # Title
    title = doc.add_heading("會議紀錄", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.runs[0]
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    run.font.name = "Microsoft JhengHei"

    # Template-specific header info
    if template_choice == "Standard":
        info = doc.add_paragraph()
        info.alignment = WD_ALIGN_PARAGRAPH.LEFT
        info_run = info.add_run("Date: _______________    Time: _______________    Location: _______________")
        info_run.font.size = Pt(10)
        info_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        info_run.font.name = "Microsoft JhengHei"
        doc.add_paragraph()
    elif template_choice == "Executive":
        info = doc.add_paragraph()
        info.alignment = WD_ALIGN_PARAGRAPH.LEFT
        info_run = info.add_run("CONFIDENTIAL\nDate: _______________    Chair: _______________")
        info_run.font.size = Pt(10)
        info_run.font.color.rgb = RGBColor(0x99, 0x00, 0x00)
        info_run.font.name = "Microsoft JhengHei"
        doc.add_paragraph()

    # Parse markdown table
    rows = parse_md_table(md_text)
    if rows and len(rows) >= 1:
        num_cols = max(len(r) for r in rows)
        table = doc.add_table(rows=1, cols=num_cols)
        table.style = 'Table Grid'

        # Header
        hdr_cells = table.rows[0].cells
        for i in range(num_cols):
            if i < len(rows[0]):
                hdr_cells[i].text = rows[0][i]
            hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in hdr_cells[i].paragraphs[0].runs:
                run.font.bold = True
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.name = "Microsoft JhengHei"
            # Dark background via shading
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), '1A365D')
            hdr_cells[i]._tc.get_or_add_tcPr().append(shading_elm)

        # Data rows
        for row_data in rows[1:]:
            row_cells = table.add_row().cells
            for i in range(num_cols):
                if i < len(row_data):
                    row_cells[i].text = row_data[i]
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in row_cells[i].paragraphs[0].runs:
                    run.font.size = Pt(10)
                    run.font.name = "Microsoft JhengHei"

        # Column widths
        for row in table.rows:
            for idx, cell in enumerate(row.cells):
                if idx == 0:
                    cell.width = Inches(1.8)
                elif idx == 1:
                    cell.width = Inches(2.5)
                else:
                    cell.width = Inches(2.5)
    else:
        # No table: plain text
        for line in md_text.splitlines():
            if line.strip():
                p = doc.add_paragraph(line.strip())
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                for run in p.runs:
                    run.font.size = Pt(11)
                    run.font.name = "Microsoft JhengHei"

    # Footer
    doc.add_paragraph()
    footer = doc.add_paragraph("Generated by Smart Minutes Generator")
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer.runs[0]
    footer_run.font.size = Pt(9)
    footer_run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    footer_run.font.name = "Microsoft JhengHei"

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────
# Sidebar: Built-in Templates
# ─────────────────────────────────────────
with st.sidebar:
    st.header("📋 Templates")
    template_choice = st.selectbox("Select template", ["Standard", "Executive", "Minimal"])

    if template_choice == "Standard":
        st.info("Standard meeting minutes with Date/Time/Location header.")
    elif template_choice == "Executive":
        st.info("Executive summary format with CONFIDENTIAL watermark.")
    else:
        st.info("Minimal format with just the table.")

    st.markdown("---")
    st.markdown("### 📥 Download Templates")
    # Generate a dummy template for download
    dummy_doc = Document()
    dummy_doc.add_heading("Meeting Minutes Template", 0)
    dummy_doc.add_paragraph("Date: _______________")
    dummy_doc.add_paragraph("Time: _______________")
    dummy_doc.add_paragraph("Location: _______________")
    t = dummy_doc.add_table(rows=1, cols=3)
    t.style = 'Table Grid'
    hdr = t.rows[0].cells
    hdr[0].text = "Topic"
    hdr[1].text = "Discussion"
    hdr[2].text = "Action / Owner"
    dummy_buf = io.BytesIO()
    dummy_doc.save(dummy_buf)
    dummy_buf.seek(0)
    st.download_button("Download Word Template", dummy_buf, "template.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ─────────────────────────────────────────
# Main UI
# ─────────────────────────────────────────
st.title("📝 Smart Minutes Generator")
st.markdown("AI-powered meeting minutes generator with PII protection & BYOK support.")

# BYOK UI
with st.expander("🔑 企業資安選項：自備 API Key (BYOK)", expanded=False):
    use_byok = st.toggle(
        "啟用自備 API Key",
        value=st.session_state.get("use_byok", False),
        key="byok_toggle"
    )
    st.session_state["use_byok"] = use_byok

    if use_byok:
        provider = st.selectbox(
            "API Provider",
            ["GitHub Models", "OpenAI", "Azure OpenAI", "Custom"],
            index=["GitHub Models", "OpenAI", "Azure OpenAI", "Custom"].index(
                st.session_state.get("byok_provider", "GitHub Models")
            ),
            key="byok_provider_sel"
        )
        st.session_state["byok_provider"] = provider

        byok_key = st.text_input(
            "API Key",
            value=st.session_state.get("byok_key", ""),
            type="password",
            key="byok_key_input"
        )
        st.session_state["byok_key"] = byok_key

        if provider == "GitHub Models":
            default_model = "gpt-4o"
            default_url = "https://models.inference.ai.azure.com"
        elif provider == "OpenAI":
            default_model = "gpt-4o"
            default_url = "https://api.openai.com/v1"
        elif provider == "Azure OpenAI":
            default_model = "gpt-4o"
            default_url = ""
        else:
            default_model = ""
            default_url = ""

        byok_model = st.text_input(
            "Model Name",
            value=st.session_state.get("byok_model", default_model) or default_model,
            key="byok_model_input"
        )
        st.session_state["byok_model"] = byok_model

        byok_url = st.text_input(
            "Base URL (optional)",
            value=st.session_state.get("byok_url", default_url) or default_url,
            key="byok_url_input"
        )
        st.session_state["byok_url"] = byok_url

        if not byok_key.strip():
            st.warning("請輸入 API Key，否則無法使用自備 Key 模式。")
    else:
        byok_key = st.secrets.get("GITHUB_TOKEN", "")
        byok_model = "gpt-4o"
        byok_url = "https://models.inference.ai.azure.com"

st.markdown("---")

# ─────────────────────────────────────────
# Input Mode Selection
# ─────────────────────────────────────────
input_mode = st.radio("Choose input mode:", ["Upload PPT / PDF", "Paste draft text"])

uploaded_template = None
ppt_text = ""
draft_text = ""

if input_mode == "Upload PPT / PDF":
    uploaded_ppt = st.file_uploader("Upload PowerPoint", type=["pptx"])
    if uploaded_ppt:
        prs = Presentation(uploaded_ppt)
        slides_text = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slides_text.append(shape.text.strip())
        ppt_text = "\n".join(slides_text)
        st.success(f"Extracted text from {len(prs.slides)} slides.")
        with st.expander("Preview extracted text"):
            st.text(ppt_text[:2000])

    st.markdown("---")
    st.markdown("### 📄 Optional: Upload your own Word template")
    uploaded_template = st.file_uploader(
        "Upload custom Word template (.docx)",
        type=["docx"],
        key="custom_template"
    )
    if uploaded_template:
        st.success("Custom template uploaded. The system will preserve its logos, watermarks, headers/footers, and fonts.")
else:
    draft_text = st.text_area("Paste your draft meeting content here:", height=300)

    st.markdown("---")
    st.markdown("### 📄 Optional: Upload your own Word template")
    uploaded_template = st.file_uploader(
        "Upload custom Word template (.docx)",
        type=["docx"],
        key="custom_template"
    )
    if uploaded_template:
        st.success("Custom template uploaded. The system will preserve its logos, watermarks, headers/footers, and fonts.")

# ─────────────────────────────────────────
# Generate Minutes
# ─────────────────────────────────────────
if st.button("✨ Generate Minutes", type="primary"):
    source_text = ppt_text if input_mode == "Upload PPT / PDF" else draft_text
    if not source_text.strip():
        st.error("Please provide input content (upload PPT or paste draft text).")
        st.stop()

    # Resolve API key / model / base_url
    if st.session_state.get("use_byok", False):
        api_key = st.session_state.get("byok_key", "").strip()
        model = st.session_state.get("byok_model", "gpt-4o").strip()
        base_url = st.session_state.get("byok_url", "").strip()
        if not api_key:
            st.error("BYOK is enabled but no API Key provided.")
            st.stop()
    else:
        api_key = st.secrets.get("GITHUB_TOKEN", "")
        model = "gpt-4o"
        base_url = "https://models.inference.ai.azure.com"
        if not api_key:
            st.error("No API key found. Please enable BYOK and provide your key, or set GITHUB_TOKEN in secrets.")
            st.stop()

    # PII Masking
    masker = PIIMasker()
    masked_text = masker.mask(source_text)

    # System prompt with PII token preservation instruction
    system_prompt = (
        "You are a professional meeting minutes assistant. "
        "Generate structured meeting minutes in Markdown table format. "
        "Preserve any [REDACTED_...] tokens exactly as they appear in the input. "
        "Do not alter, translate, or remove these tokens."
    )

    user_prompt = (
        f"Generate meeting minutes from the following content. "
        f"Use a Markdown table with columns: Topic, Discussion, Action/Owner.\n\n{masked_text}"
    )

    with st.spinner("Generating minutes with AI..."):
        try:
            client = get_openai_client(api_key, base_url)
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.3,
            )
            raw_output = response.choices[0].message.content
        except Exception as e:
            st.error(f"API call failed: {e}")
            st.stop()

    # Unmask PII
    final_output = masker.unmask(raw_output)

    st.markdown("---")
    st.subheader("📄 Generated Minutes")
    st.markdown(final_output)

    # ─────────────────────────────────────────
    # Export to Word (Template-aware)
    # ─────────────────────────────────────────
    try:
        if uploaded_template is not None:
            word_buf = fill_user_template(uploaded_template, final_output)
            download_label = "Download Word (Custom Template)"
        else:
            word_buf = convert_md_to_docx(final_output, template_choice)
            download_label = "Download Word (Built-in Template)"
    except Exception as e:
        st.warning(f"Custom template filling failed ({e}). Falling back to built-in template.")
        word_buf = convert_md_to_docx(final_output, template_choice)
        download_label = "Download Word (Built-in Template Fallback)"

    st.download_button(
        label=download_label,
        data=word_buf,
        file_name="meeting_minutes.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    # Export to Markdown
    md_buf = io.BytesIO(final_output.encode("utf-8"))
    st.download_button(
        label="Download Markdown",
        data=md_buf,
        file_name="meeting_minutes.md",
        mime="text/markdown"
    )

