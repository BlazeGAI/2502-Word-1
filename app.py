import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def check_word_1(doc):
    checklist_data = {
        "Grading Criteria": [
            "Is the font Times New Roman, 12pt?",
            "Is line spacing set to double?",
            "Are margins set to 1 inch on all sides?",
            "Is the title centered and not bold?",
            "Are paragraphs properly indented?",
            "Are there at least 3 paragraphs?",
            "Is there a References page?",
            "Are in-text citations properly formatted?"
        ],
        "Completed": []
    }

    def is_correct_font(paragraph):
        """Check if the paragraph uses Times New Roman, 12pt font."""
        try:
            # Normalize font name comparison
            for run in paragraph.runs:
                # Check font name (case-insensitive and handles partial matches)
                run_font = run.font.name or paragraph.style.font.name
                if run_font and 'times new roman' not in run_font.lower():
                    return False
                
                # Check font size with some tolerance
                run_size = run.font.size or paragraph.style.font.size
                if run_size is not None:
                    # Allow small variations (e.g., ±0.5 points)
                    if abs(run_size.pt - 12) > 0.5:
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
        if header_done and not found_references and len(text) > 100 and not text.lower().startswith('in conclusion'):
            body_paragraphs.append(text)

    proper_indentation = True
    checklist_data["Completed"].append("Yes" if proper_indentation else "No")

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

    return checklist_data

def main():
    st.title("Word Document Checker")

    uploaded_file = st.file_uploader("Upload a DOCX file", type=["docx"])
    if uploaded_file is not None:
        try:
            doc = Document(uploaded_file)
        except Exception as e:
            st.error("Error processing the document. Verify the file is a valid DOCX file.")
            return

        result = check_word_1(doc)

        st.header("Checklist Results")
        for criterion, status in zip(result["Grading Criteria"], result["Completed"]):
            st.write(f"**{criterion}**: {status}")

if __name__ == "__main__":
    main()
