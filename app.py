import streamlit as st
import docx
import re
import pandas as pd
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
            "Page numbers: Top right in header"
        ],
        "Completed": []
    }

    def safe_append(value):
        checklist_data["Completed"].append("Yes" if value else "No")

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

    # Page Number Check: Only in Header, Right-aligned
    def has_correct_page_numbers(doc):
        for section in doc.sections:
            if section.header:
                for para in section.header.paragraphs:
                    if para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                        text_content = para.text.strip()
                        if text_content.isdigit() or 'PAGE' in para._element.xml:
                            return True
        return False

    has_page_numbers = has_correct_page_numbers(doc)
    safe_append(has_page_numbers)

    return checklist_data

def display_results(checklist_data):
    total_yes = checklist_data["Completed"].count("Yes")
    total_items = len(checklist_data["Completed"])
    percentage_complete = (total_yes / total_items) * 100
    points = (total_yes / total_items) * 20

    col1, col2 = st.columns(2)
    
    with col1:
        if percentage_complete == 100:
            st.success(f"Completion Score: {percentage_complete:.1f}%")
        elif percentage_complete >= 80:
            st.warning(f"Completion Score: {percentage_complete:.1f}%")
        else:
            st.error(f"Completion Score: {percentage_complete:.1f}%")

    with col2:
        if points == 20:
            st.success(f"Points: {points:.1f}/20")
        elif points >= 16:
            st.warning(f"Points: {points:.1f}/20")
        else:
            st.error(f"Points: {points:.1f}/20")

    st.subheader("📋 Detailed Checklist")
    checklist_df = pd.DataFrame(checklist_data)
    st.table(checklist_df)

st.title("📄 Word Document Grading Checklist")

uploaded_file = st.file_uploader("Upload a .docx file", type=["docx"])

if uploaded_file:
    doc = docx.Document(uploaded_file)
    results = check_word_document(doc)
    display_results(results)
