import streamlit as st
import docx
import re
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement


def check_word_document(doc):
    checklist_data = {
        "Grading Criteria": [
            "Is the font Times New Roman, 12pt?",
            "Is line spacing set to double?",
            "Are margins set to 1 inch on all sides?",
            "Is the title on the title page centered and bold?",
            "Is the title on the second page not bold?",
            "Are the paragraphs left-aligned?",
            "Are there at least 3 paragraphs?",
            "Is there a References page?",
            "Are in-text citations properly formatted?",
            "Is there a title page?",
            "Is the title page center aligned?",
            "Are there page numbers in the top right?"
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

    # Ensure title is first line of characters in document
    first_non_empty_paragraph = next((p for p in doc.paragraphs if p.text.strip()), None)
    title_bold_centered = (
        first_non_empty_paragraph and 
        first_non_empty_paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER and
        any(run.bold for run in first_non_empty_paragraph.runs)
    )
    checklist_data["Completed"].append("Yes" if title_bold_centered else "No")

    second_page_title = doc.paragraphs[1] if len(doc.paragraphs) > 1 else None
    title_not_bold = (
        second_page_title and
        not any(run.bold for run in second_page_title.runs)
    )
    checklist_data["Completed"].append("Yes" if title_not_bold else "No")

    body_paragraphs = []
    found_references = False
    header_done = False
    
    for p in doc.paragraphs:
        text = p.text.strip() if p.text else ""
        if not text:
            continue
        if not header_done and len(text) > 100:
            header_done = True
        if text.lower() == 'references':
            found_references = True
            continue
        if header_done and not found_references and len(text) > 100:
            body_paragraphs.append(p)

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

    # Improved in-text citation check using regex
    citation_pattern = re.compile(r'\. \([A-Za-z]+, \d{4}\)')
    has_citations = any(
        citation_pattern.search(p.text if p.text else "") for p in body_paragraphs
    )
    checklist_data["Completed"].append("Yes" if has_citations else "No")

    # Check for page numbers stored as PAGE fields in the header
    has_page_numbers = False
    for section in doc.sections:
        if section.header:
            for para in section.header.paragraphs:
                for run in para.runs:
                    for field in run._element.findall('.//w:fldSimple', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        if 'PAGE' in field.attrib.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instr', ''):
                            has_page_numbers = True
    checklist_data["Completed"].append("Yes" if has_page_numbers else "No")

    return checklist_data

st.title("Word Document Grading Checklist")

uploaded_file = st.file_uploader("Upload a .docx file", type=["docx"])

if uploaded_file:
    doc = docx.Document(uploaded_file)
    results = check_word_document(doc)
    
    st.subheader("Checklist Results")
    for i, criterion in enumerate(results["Grading Criteria"]):
        st.write(f"{criterion}: {results['Completed'][i]}")
