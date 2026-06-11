import streamlit as st
import pdfplumber
from docx import Document

skills = [
    "python",
    "java",
    "sql",
    "machine learning",
    "html",
    "css",
    "javascript",
    "git",
    "github"
]

def read_pdf(file):
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            if page.extract_text():
                text += page.extract_text()
    return text

def read_docx(file):
    doc = Document(file)
    text = "\n".join([p.text for p in doc.paragraphs])
    return text

st.title("AI Resume Analyzer")

uploaded_file = st.file_uploader(
    "Upload Resume",
    type=["pdf", "docx"]
)

if uploaded_file:
    if uploaded_file.name.endswith(".pdf"):
        resume_text = read_pdf(uploaded_file)
    else:
        resume_text = read_docx(uploaded_file)

    resume_text = resume_text.lower()

    found = []
    missing = []

    for skill in skills:
        if skill in resume_text:
            found.append(skill)
        else:
            missing.append(skill)

    score = int((len(found) / len(skills)) * 100)

    st.subheader("ATS Score")
    st.progress(score / 100)
    st.write(f"Score: {score}%")

    st.subheader("Skills Found")
    st.success(", ".join(found) if found else "None")

    st.subheader("Missing Skills")
    st.error(", ".join(missing) if missing else "None")

    if score >= 80:
        st.success("Excellent Resume")
    elif score >= 60:
        st.warning("Good Resume - Add more skills")
    else:
        st.error("Resume Needs Improvement")