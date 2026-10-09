# Northstar Resume Intelligence

A Streamlit resume analyzer that helps job seekers review resume content, compare recognized skills with a target role, and download a concise PDF report.

## Features

- Upload PDF and DOCX resumes, including text contained in DOCX tables.
- Review skill coverage across programming, web development, data and AI, and tools and cloud.
- Compare a resume against built-in role profiles or skills found in a pasted job description.
- See missing role skills and job-description keywords to consider.
- Review document signals such as word count, contact details, common section headings, and frequently used keywords.
- Generate and download a PDF summary with role-match and skill-coverage details.
- Browse a responsive dashboard with category charts, an overview, role-match analysis, and a recruiter snapshot.

## Requirements

- Python 3.10 or later
- Packages listed in [`requirements.txt`](./requirements.txt)

## Run locally

From PowerShell in the project directory:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run analyzer.py
```

Open [http://localhost:8501](http://localhost:8501) in your browser. Upload a PDF or DOCX resume using the sidebar. Select a target role, paste a job description for a tailored comparison, or use both together. Open **Recruiter snapshot & report** to download the PDF.

## How matching works

The analyzer uses local, rule-based keyword matching against its built-in skill list and role profiles. The overall coverage score is the proportion of built-in skills detected in the resume. When a job description contains recognizable built-in skills, the target match score is the proportion of those skills detected in the resume. Without recognizable job-description skills, a selected role profile provides the comparison instead.

This is a keyword aid, not an ATS simulation, a measure of candidate quality, or a hiring recommendation. A match does not verify proficiency, and a missing keyword does not mean a person lacks that experience. Add skills and achievements only when they accurately represent your background.

## Privacy and limitations

The app does not intentionally save uploaded resumes or send them to an AI service. Files are processed by the Streamlit app during the session. If you deploy the app to a hosted server, uploads are processed by that server and its privacy and retention policies apply.

Scanned or image-only PDFs may not contain extractable text; this app does not perform OCR. Use a text-based PDF or DOCX for best results.
