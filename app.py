import streamlit as st
import docx
import re
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def check_word_document(doc):
    checklist_data = {
        "Grading Criteria": [
            "Font: Times New Roman, 12pt",
            "Line spacing: Double",
            "Margins: 1 inch on all sides",
            "Title on title page: Centered & Bold",
            "Title on second page: Not bold",
            "Paragraph alignment: Left-aligned",
            "Minimum of 3 paragraphs",
            "References page present",
            "Proper in-text citations",
            "Title page present",
            "Title page text: Center aligned",
            "Page numbers: Top right"
        ],
        "Completed": []
    }

    def safe_append(value):
        checklist_data["Completed"].append("✅" if value else "❌")

    default_font_name = doc.styles['Normal'].font.name if 'Normal' in doc.styles else "Times New Roman"
    default_font_size = doc.styles['Normal'].font.size if 'Normal' in doc.styles else Pt(12)
    assumed_font_size = default_font_size or Pt(12)

    def is_correct_font(paragraph):
        for run in paragraph.runs:
            run_font = run.font.name or default_font_name
            run_size = run.font.size or assumed_font_size
            if run_font != 'Times New Roman' or run_size != Pt(12):
                return False
        return True

    safe_append(all(is_correct_font(p) for p in doc.paragraphs if p.text.strip()) if doc.paragraphs else False)
    safe_append(all(p.paragraph_format.line_spacing in [None, 2.0] for p in doc.paragraphs if p.text.strip()) if doc.paragraphs else False)
    safe_append(all(
        section.left_margin.inches == 1 and section.right_margin.inches == 1 and
        section.top_margin.inches == 1 and section.bottom_margin.inches == 1
        for section in doc.sections
    ) if doc.sections else False)

    first_paragraph = next((p for p in doc.paragraphs if p.text.strip()), None)
    safe_append(first_paragraph and first_paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER and any(run.bold for run in first_paragraph.runs))

    second_page_title = doc.paragraphs[1] if len(doc.paragraphs) > 1 else None
    safe_append(second_page_title and not any(run.bold for run in second_page_title.runs))

    body_paragraphs = [p for p in doc.paragraphs if len(p.text.strip()) > 50]
    safe_append(all(p.alignment in [WD_ALIGN_PARAGRAPH.LEFT, None] for p in body_paragraphs) if body_paragraphs else False)
    safe_append(len(body_paragraphs) >= 3)

    found_references = any(p.text.strip().lower() == 'references' for p in doc.paragraphs)
    safe_append(found_references)

    citation_pattern = re.compile(r'\. \([A-Za-z]+, \d{4}\)')
    safe_append(any(citation_pattern.search(p.text if p.text else "") for p in body_paragraphs) if body_paragraphs else False)

    title_page_paragraphs = [p for p in doc.paragraphs[:5] if p.text.strip()]
    safe_append(len(title_page_paragraphs) > 0 and all(len(p.text) < 100 for p in title_page_paragraphs))
    safe_append(title_page_paragraphs and all(p.alignment == WD_ALIGN_PARAGRAPH.CENTER for p in title_page_paragraphs))

    has_page_numbers = all(
        section.header and any(
            para.alignment == WD_ALIGN_PARAGRAPH.RIGHT and para.text.strip().isdigit()
            for para in section.header.paragraphs if para.text.strip()
        ) for section in doc.sections
    ) if doc.sections else False
    safe_append(has_page_numbers)

    return checklist_data

st.title("📄 Word Document Grading Checklist")

uploaded_file = st.file_uploader("Upload a .docx file", type=["docx"])

if uploaded_file:
    doc = docx.Document(uploaded_file)
    results = check_word_document(doc)

    st.subheader("✅ Checklist Results")

    # Show the results in a table
    st.write(
        "| Criteria | Status |\n"
        "|---------|--------|\n" +
        "\n".join(f"| {criterion} | {status} |" for criterion, status in zip(results["Grading Criteria"], results["Completed"]))
    )

    # Calculate percentage completion
    completed_count = results["Completed"].count("✅")
    total_criteria = len(results["Completed"])
    progress = completed_count / total_criteria

    st.subheader("📊 Compliance Score")
    st.progress(progress)

    # Expandable details section
    with st.expander("🔍 Detailed Checklist"):
        for i, criterion in enumerate(results["Grading Criteria"]):
            st.write(f"- {criterion}: {results['Completed'][i]}")
