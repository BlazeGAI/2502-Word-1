import streamlit as st
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def check_word_document(doc):
    checklist_data = {
        "Grading Criteria": [
            "Is the font Times New Roman, 12pt?",
            "Is line spacing set to double?",
            "Are margins set to 1 inch on all sides?",
            "Is the title centered and not bold?",
            "Are the paragraphs left-aligned?",
            "Are there at least 3 paragraphs?",
            "Is there a References page?",
            "Are in-text citations properly formatted?",
            "Is there a title page?",
            "Is the title page center aligned?"
        ],
        "Completed": []
    }

    # Retrieve document default font settings
    default_font_name = doc.styles['Normal'].font.name
    default_font_size = doc.styles['Normal'].font.size
    assumed_font_size = Pt(12) if default_font_size is None else default_font_size

    def is_correct_font(paragraph):
        for run in paragraph.runs:
            run_font = run.font.name or default_font_name
            run_size = run.font.size or assumed_font_size
            
            if run_font != 'Times New Roman' or run_size != Pt(12):
                return False
        return True
    
    correct_font = all(is_correct_font(p) for p in doc.paragraphs if p.text.strip())
    checklist_data["Completed"].append("Yes" if correct_font else "No")

    correct_spacing = all(
        p.paragraph_format.line_spacing in [None, 2.0] for p in doc.paragraphs if p.text.strip()
    )
    checklist_data["Completed"].append("Yes" if correct_spacing else "No")

    correct_margins = all(
        section.left_margin.inches == 1 and
        section.right_margin.inches == 1 and
        section.top_margin.inches == 1 and
        section.bottom_margin.inches == 1
        for section in doc.sections
    )
    checklist_data["Completed"].append("Yes" if correct_margins else "No")

    title_paragraph = doc.paragraphs[0] if doc.paragraphs else None
    title_centered = (
        title_paragraph and 
        title_paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER and
        not any(run.bold for run in title_paragraph.runs)
    )
    checklist_data["Completed"].append("Yes" if title_centered else "No")

    body_paragraphs = []
    found_references = False
    header_done = False
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        if not header_done and len(text) > 100:
            header_done = True
        if text.lower() == 'references':
            found_references = True
            continue
        if header_done and not found_references and len(text) > 100:
            body_paragraphs.append(text)

    # Improved left-aligned paragraph detection (ignoring short headers/titles)
    body_text_paragraphs = [p for p in doc.paragraphs if len(p.text.strip()) > 50]
    left_aligned = all(p.alignment in [WD_ALIGN_PARAGRAPH.LEFT, None] for p in body_text_paragraphs)
    checklist_data["Completed"].append("Yes" if left_aligned else "No")

    sufficient_paragraphs = len(body_paragraphs) >= 3
    checklist_data["Completed"].append("Yes" if sufficient_paragraphs else "No")

    has_references = any(p.text.strip().lower() == 'references' for p in doc.paragraphs)
    references_content = any(
        has_references and '(' in p.text and ')' in p.text for p in doc.paragraphs
    )
    checklist_data["Completed"].append("Yes" if (has_references and references_content) else "No")

    has_citations = any(
        '(' in p and ')' in p and any(str(year) in p for year in range(1900, 2025))
        for p in body_paragraphs
    )
    checklist_data["Completed"].append("Yes" if has_citations else "No")

    # Improved title page detection
    first_page_paragraphs = [p for p in doc.paragraphs[:5] if p.text.strip()]
    title_page_exists = len(first_page_paragraphs) > 0 and all(len(p.text) < 100 for p in first_page_paragraphs)
    checklist_data["Completed"].append("Yes" if title_page_exists else "No")

    title_page_centered = title_page_exists and all(p.alignment == WD_ALIGN_PARAGRAPH.CENTER for p in first_page_paragraphs)
    checklist_data["Completed"].append("Yes" if title_page_centered else "No")

    return checklist_data

st.title("Word Document Grading Checklist")

uploaded_file = st.file_uploader("Upload a .docx file", type=["docx"])

if uploaded_file:
    doc = docx.Document(uploaded_file)
    results = check_word_document(doc)
    
    st.subheader("Checklist Results")
    for i, criterion in enumerate(results["Grading Criteria"]):
        st.write(f"{criterion}: {results['Completed'][i]}")
