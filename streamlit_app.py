import streamlit as st
import pandas as pd
from docx import Document  # from python-docx
from pptx import Presentation  # from python-pptx
import fitz  # PyMuPDF for PDFs
import io

st.set_page_config(page_title="M365 Accessibility Pre-Flight Check", layout="wide")
st.title("M365 Accessibility Pre-Flight Check")
st.write("Upload your documents (.docx, .pptx, .pdf) to check for basic accessibility issues.")

# --- Accessibility check functions ---
def check_docx(file):
    issues = []
    doc = Document(file)
    if not any(p.style.name.startswith("Heading") for p in doc.paragraphs if p.style):
        issues.append("No headings found (use Heading styles for structure).")
    for p in doc.paragraphs:
        for run in p.runs:
            if "click here" in run.text.lower():
                issues.append("Avoid vague link text like 'click here'.")
    return issues

def check_pptx(file):
    issues = []
    prs = Presentation(file)
    for i, slide in enumerate(prs.slides, start=1):
        if not slide.shapes.title:
            issues.append(f"Slide {i}: Missing title.")
    issues.append("Reminder: Alt text for images and color contrast must be verified manually.")
    return issues

def check_pdf(file):
    issues = []
    pdf = fitz.open(stream=file.read(), filetype="pdf")
    if not pdf.is_pdf:
        issues.append("Not a valid PDF file.")
        return issues
    text = "".join(page.get_text() for page in pdf)
    if "# " not in text and "Heading" not in text:
        issues.append("No obvious heading structure detected.")
    issues.append("Reminder: Check that the PDF is tagged for accessibility and has alt text.")
    return issues

# --- File uploader ---
uploaded_files = st.file_uploader("Upload files", type=["docx", "pptx", "pdf"], accept_multiple_files=True)

results = []
if uploaded_files:
    for file in uploaded_files:
        if file.name.endswith(".docx"):
            issues = check_docx(file)
        elif file.name.endswith(".pptx"):
            issues = check_pptx(file)
        elif file.name.endswith(".pdf"):
            file.seek(0)
            issues = check_pdf(file)
        else:
            issues = ["Unsupported file type."]
        
        results.append({"File": file.name, "Issues": "; ".join(issues) if issues else "No major issues found."})

if results:
    df = pd.DataFrame(results)
    st.subheader("Accessibility Findings")
    st.dataframe(df)

    # Downloadable CSV report
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button("Download CSV Report", data=csv, file_name="accessibility_report.csv", mime="text/csv")

    # Downloadable HTML report
    html = df.to_html(index=False)
    st.download_button("Download HTML Report", data=html, file_name="accessibility_report.html", mime="text/html")
