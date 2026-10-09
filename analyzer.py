import html
import os
import re
import zipfile
from collections import Counter
from datetime import datetime
from io import BytesIO

import pdfplumber
import streamlit as st
from docx import Document
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


BRAND = "NORTHSTAR"
skills_by_category = {
    "Programming": ["python", "java", "sql", "javascript", "c++", "c#"],
    "Web development": [
        "html",
        "css",
        "react",
        "node.js",
        "flask",
        "django",
        "fastapi",
    ],
    "Data and AI": [
        "machine learning",
        "deep learning",
        "data analysis",
        "pandas",
        "numpy",
        "scikit-learn",
        "pytorch",
        "tensorflow",
        "nlp",
    ],
    "Tools and cloud": [
        "git",
        "github",
        "docker",
        "linux",
        "aws",
        "azure",
        "tableau",
        "agile",
        "jira",
    ],
}
all_skills = [skill for category in skills_by_category.values() for skill in category]

role_profiles = {
    "Software Engineer": [
        "python",
        "java",
        "sql",
        "javascript",
        "git",
        "github",
        "docker",
        "agile",
    ],
    "Data Scientist": [
        "python",
        "sql",
        "machine learning",
        "deep learning",
        "data analysis",
        "pandas",
        "numpy",
        "scikit-learn",
        "tableau",
    ],
    "Frontend Developer": [
        "javascript",
        "html",
        "css",
        "react",
        "git",
        "github",
    ],
    "Backend Developer": [
        "python",
        "java",
        "sql",
        "node.js",
        "flask",
        "django",
        "fastapi",
        "docker",
        "linux",
    ],
    "Cloud / DevOps Engineer": [
        "python",
        "linux",
        "aws",
        "azure",
        "docker",
        "git",
        "github",
        "agile",
    ],
}

stopwords = {
    "about", "after", "also", "are", "based", "been", "being", "build", "can",
    "candidate", "company", "develop", "development", "experience", "for",
    "from", "have", "into", "including", "looking", "must", "our", "projects",
    "qualifications", "required", "requirements", "resume", "role", "skills",
    "strong", "such", "team", "that", "the", "their", "this", "through",
    "understanding", "using", "various", "will", "with", "work", "working",
    "years", "you",
}


def normalize_text(text):
    text = text.lower().replace("\u2013", " ").replace("\u2014", " ")
    text = re.sub(r"[^a-z0-9+.#\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def contains_skill(text, skill):
    normalized_skill = normalize_text(skill)
    pattern = r"(?<![a-z0-9])" + re.escape(normalized_skill) + r"(?![a-z0-9])"
    return bool(re.search(pattern, text))


def _read_file_bytes(file):
    if isinstance(file, (str, os.PathLike)):
        with open(file, "rb") as source:
            return source.read()

    position = file.tell()
    file.seek(0)
    content = file.read()
    file.seek(position)
    return content


def read_pdf(file):
    text_parts = []
    with pdfplumber.open(BytesIO(_read_file_bytes(file))) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text_parts.append(page_text)
    return "\n".join(text_parts)


def read_docx(file):
    document = Document(BytesIO(_read_file_bytes(file)))
    content = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            content.append(" ".join(cell.text for cell in row.cells))
    return "\n".join(content)


def extract_top_keywords(text, limit=10):
    words = re.findall(r"\b[a-z][a-z0-9+#.-]{2,}\b", normalize_text(text))
    counts = Counter(word for word in words if word not in stopwords)
    return counts.most_common(limit)


def extract_skills(text):
    normalized = normalize_text(text)
    return sorted(skill for skill in all_skills if contains_skill(normalized, skill))


def summarize_resume(text):
    normalized = normalize_text(text)
    headings = [
        heading
        for heading in (
            "education",
            "experience",
            "employment",
            "projects",
            "certifications",
        )
        if re.search(rf"\b{heading}\b", normalized)
    ]
    return {
        "word_count": len(normalized.split()),
        "email_present": bool(
            re.search(r"\b[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}\b", text, re.I)
        ),
        "phone_present": bool(
            re.search(
                r"\b(?:\+?\d{1,3}[-.\s]?)?(?:\(?\d{3}\)?[-.\s]?)\d{3}[-.\s]?\d{4}\b",
                text,
            )
        ),
        "sections": headings,
    }


def build_job_recommendations(job_description, resume_text, job_skills):
    missing_skills = [skill for skill in job_skills if skill not in extract_skills(resume_text)]
    job_keywords = extract_top_keywords(job_description, limit=30)
    resume_words = set(re.findall(r"\b[a-z][a-z0-9+#.-]{2,}\b", normalize_text(resume_text)))
    extra_keywords = [
        word
        for word, _ in job_keywords
        if word not in resume_words and not any(contains_skill(word, skill) for skill in all_skills)
    ]
    return missing_skills, extra_keywords[:5]


def build_pdf_report(
    resume_name,
    role,
    resume_summary,
    overall_score,
    found_skills,
    job_skills,
    matched_skills,
    missing_skills,
    keyword_suggestions,
    category_scores,
):
    buffer = BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=20 * mm,
        leftMargin=20 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title="Northstar Resume Match Report",
    )
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="NorthstarTitle",
            parent=styles["Title"],
            alignment=TA_LEFT,
            textColor=colors.HexColor("#102a43"),
            fontSize=23,
            leading=28,
            spaceAfter=5,
        )
    )
    styles.add(
        ParagraphStyle(
            name="SectionHeading",
            parent=styles["Heading2"],
            textColor=colors.HexColor("#087e78"),
            fontSize=12,
            leading=16,
            spaceBefore=13,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ReportBody",
            parent=styles["BodyText"],
            fontSize=9,
            leading=14,
            textColor=colors.HexColor("#344b60"),
        )
    )

    def safe_list(values):
        return ", ".join(html.escape(value) for value in values) or "None identified"

    story = [
        Paragraph("NORTHSTAR", styles["ReportBody"]),
        Paragraph("Resume match report", styles["NorthstarTitle"]),
        Paragraph(
            f"Prepared {datetime.now().strftime('%B %d, %Y')} | "
            f"Resume: {html.escape(resume_name)}",
            styles["ReportBody"],
        ),
        Spacer(1, 10),
        Paragraph("At a glance", styles["SectionHeading"]),
    ]
    metrics = [
        ["Overall skill coverage", "Target role match", "Resume length"],
        [
            f"{overall_score}%",
            f"{int(len(matched_skills) / len(job_skills) * 100) if job_skills else 0}%",
            f"{resume_summary['word_count']} words",
        ],
    ]
    metric_table = Table(metrics, colWidths=[55 * mm, 55 * mm, 55 * mm])
    metric_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#edf5f4")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#087e78")),
                ("TEXTCOLOR", (0, 1), (-1, 1), colors.HexColor("#102a43")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("PADDING", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#d9e5e3")),
                ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d9e5e3")),
            ]
        )
    )
    story.extend(
        [
            metric_table,
            Paragraph("Recruiter snapshot", styles["SectionHeading"]),
            Paragraph(
                f"Target role: {html.escape(role)}. The resume contains "
                f"{resume_summary['word_count']} words, "
                f"{'an email address' if resume_summary['email_present'] else 'no detected email address'}, "
                f"and {'a phone number' if resume_summary['phone_present'] else 'no detected phone number'}. "
                f"Detected resume sections: {safe_list(resume_summary['sections'])}.",
                styles["ReportBody"],
            ),
            Paragraph("Skills aligned to the target", styles["SectionHeading"]),
            Paragraph(safe_list(matched_skills), styles["ReportBody"]),
            Paragraph("Priority skills to evidence", styles["SectionHeading"]),
            Paragraph(safe_list(missing_skills), styles["ReportBody"]),
            Paragraph("Additional job-description keywords to consider", styles["SectionHeading"]),
            Paragraph(safe_list(keyword_suggestions), styles["ReportBody"]),
            Paragraph("Overall skills found", styles["SectionHeading"]),
            Paragraph(safe_list(found_skills), styles["ReportBody"]),
            Paragraph("Coverage by skill area", styles["SectionHeading"]),
        ]
    )
    for category, score in category_scores.items():
        story.append(
            Paragraph(
                f"{html.escape(category)}: {score[0]} of {score[1]} skills",
                styles["ReportBody"],
            )
        )
    story.extend(
        [
            Spacer(1, 14),
            Paragraph(
                "This report is an automated keyword comparison, not a hiring decision. "
                "Only include skills and experience you can accurately demonstrate.",
                styles["ReportBody"],
            ),
        ]
    )
    document.build(story)
    return buffer.getvalue()


def render_styles():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
        :root { --ink:#102a43; --muted:#627d98; --teal:#087e78; --mint:#e8f6f2; }
        .stApp { background:linear-gradient(180deg,#f5f9fc 0%,#ffffff 340px); color:var(--ink); }
        [data-testid="stHeader"] { background:rgba(245,249,252,.9); }
        [data-testid="stSidebar"] { background:#102a43; }
        [data-testid="stSidebar"] * { color:#edf6f6; }
        [data-testid="stSidebar"] [data-baseweb="select"] * { color:#102a43; }
        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] { background:#183b56; border-color:#4a7188; }
        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] * { color:#edf6f6; }
        html, body, [class*="css"] { font-family:'DM Sans',sans-serif; }
        h1,h2,h3 { font-family:'Manrope',sans-serif; color:var(--ink); letter-spacing:-.035em; }
        .block-container { max-width:1200px; padding-top:1.2rem; padding-bottom:4rem; }
        .brandbar { display:flex; align-items:center; justify-content:space-between; padding:7px 0 22px; border-bottom:1px solid #e2ebf0; margin-bottom:25px; }
        .brandmark { display:flex; align-items:center; gap:10px; color:#102a43; font-weight:800; letter-spacing:.13em; font-size:14px; }
        .brandicon { width:31px; height:31px; display:grid; place-items:center; border-radius:10px; color:white; background:linear-gradient(145deg,#087e78,#35b69b); font-size:17px; }
        .brandnav { color:#627d98; font-size:13px; }
        .hero { background:linear-gradient(120deg,#102a43 0%,#16435b 67%,#087e78 150%); border-radius:22px; padding:32px 36px; color:white; margin:0 0 23px; box-shadow:0 16px 38px rgba(16,42,67,.12); }
        .hero-kicker { text-transform:uppercase; letter-spacing:.16em; font-size:11px; font-weight:700; color:#8ee0c9; margin-bottom:11px; }
        .hero h1 { color:white; font-size:clamp(28px,4vw,42px); line-height:1.12; margin:0 0 11px; }
        .hero p { color:#d5e6ed; font-size:15px; line-height:1.65; max-width:710px; margin:0; }
        .section-label { color:#087e78; text-transform:uppercase; letter-spacing:.12em; font-size:11px; font-weight:700; margin:22px 0 6px; }
        .stMetric { background:white; border:1px solid #e2ebf0; padding:15px 18px; border-radius:15px; box-shadow:0 4px 14px rgba(16,42,67,.035); }
        [data-testid="stMetricLabel"] { color:#627d98; }
        [data-testid="stMetricValue"] { color:#102a43; font-family:'Manrope',sans-serif; }
        div[data-testid="stProgress"] > div > div { background:linear-gradient(90deg,#087e78,#53c49f); }
        div[data-testid="stTabs"] button { font-weight:600; }
        div[data-testid="stTabs"] button[aria-selected="true"] { color:#087e78; }
        .note-card { background:#fff; border:1px solid #e2ebf0; border-radius:15px; padding:19px 20px; min-height:106px; }
        .note-card h4 { color:#102a43; font-size:15px; margin:0 0 7px; }
        .note-card p { color:#627d98; font-size:13px; line-height:1.55; margin:0; }
        .sidebar-brand { padding:6px 0 20px; border-bottom:1px solid rgba(255,255,255,.16); margin-bottom:15px; }
        .sidebar-brand strong { display:block; font-family:'Manrope',sans-serif; letter-spacing:.13em; font-size:14px; }
        .sidebar-brand span { color:#b7d0d9; font-size:12px; }
        .fine-print { font-size:12px; color:#627d98; line-height:1.55; }
        </style>
        """,
        unsafe_allow_html=True,
    )


st.set_page_config(
    page_title="Northstar | Resume Intelligence",
    page_icon="N",
    layout="wide",
    initial_sidebar_state="expanded",
)
render_styles()

with st.sidebar:
    st.markdown(
        '<div class="sidebar-brand"><strong>NORTHSTAR</strong>'
        "<span>Career intelligence studio</span></div>",
        unsafe_allow_html=True,
    )
    st.markdown("### Start an analysis")
    uploaded_file = st.file_uploader("Upload a resume", type=["pdf", "docx"])
    selected_role = st.selectbox("Target role", ["Use job description", *role_profiles])
    job_description = st.text_area(
        "Job description",
        height=180,
        placeholder="Paste the job description to see a tailored match...",
        help="A job description gives the most relevant skill and keyword comparison.",
    )
    st.caption("PDF and DOCX supported. Your resume is analyzed in this session.")

st.markdown(
    """
    <div class="brandbar">
      <div class="brandmark"><span class="brandicon">N</span> NORTHSTAR</div>
      <div class="brandnav">CAREER INTELLIGENCE &nbsp;·&nbsp; PRIVATE BY DESIGN</div>
    </div>
    <div class="hero">
      <div class="hero-kicker">Resume intelligence, made clear</div>
      <h1>Make your next move<br>with confidence.</h1>
      <p>See how your experience reads, identify the skills a role asks for, and leave with a practical plan to strengthen your application.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if uploaded_file is None:
    st.markdown('<div class="section-label">Your career workspace</div>', unsafe_allow_html=True)
    st.subheader("A clearer picture of your next opportunity")
    st.write(
        "Upload your resume from the sidebar to get a private, evidence-based review. "
        "Add a target role or paste a job description to make the analysis specific."
    )
    intro_cols = st.columns(3)
    intro_cards = [
        ("01 · Understand", "Review resume structure, contact details, keywords, and skill coverage."),
        ("02 · Compare", "See your alignment with a role profile or a specific job description."),
        ("03 · Take action", "Get a focused improvement plan and export a shareable PDF report."),
    ]
    for column, (title, description) in zip(intro_cols, intro_cards):
        with column:
            st.markdown(
                f'<div class="note-card"><h4>{title}</h4><p>{description}</p></div>',
                unsafe_allow_html=True,
            )
    st.markdown("")
    st.info("To begin, choose a PDF or DOCX resume in the left-hand panel.")
    st.stop()

file_name = uploaded_file.name or "resume"
if file_name.lower().endswith(".pdf"):
    resume_text = read_pdf(uploaded_file)
else:
    try:
        resume_text = read_docx(uploaded_file)
    except (ValueError, zipfile.BadZipFile):
        st.error("This DOCX file could not be read. Please upload a valid Word document.")
        st.stop()

if not resume_text.strip():
    st.error("No selectable text was found in this file. Try a text-based PDF or DOCX resume.")
    st.stop()

resume_skills = extract_skills(resume_text)
missing_overall = [skill for skill in all_skills if skill not in resume_skills]
overall_score = round(len(resume_skills) / len(all_skills) * 100) if all_skills else 0
resume_summary = summarize_resume(resume_text)
resume_keywords = extract_top_keywords(resume_text)

if job_description.strip():
    target_skills = extract_skills(job_description)
    target_label = (
        selected_role
        if selected_role != "Use job description"
        else "Custom job description"
    )
    if not target_skills and selected_role != "Use job description":
        target_skills = role_profiles[selected_role]
elif selected_role != "Use job description":
    target_skills = role_profiles[selected_role]
    target_label = f"{selected_role} match"
else:
    target_skills = []
    target_label = "Add a role or job description"

matched_skills = [skill for skill in target_skills if skill in resume_skills]
missing_job_skills, extra_keyword_suggestions = build_job_recommendations(
    job_description, resume_text, target_skills
) if job_description.strip() else (
    [skill for skill in target_skills if skill not in resume_skills],
    [],
)
job_score = round(len(matched_skills) / len(target_skills) * 100) if target_skills else 0

found_by_category = {
    category: [skill for skill in skills if skill in resume_skills]
    for category, skills in skills_by_category.items()
}
category_scores = {
    category: (len(found_by_category[category]), len(skills))
    for category, skills in skills_by_category.items()
}

st.markdown(
    f'<div class="section-label">Analysis for {html.escape(file_name)}</div>',
    unsafe_allow_html=True,
)
metric_cols = st.columns(4)
metric_cols[0].metric("Overall skill coverage", f"{overall_score}%")
metric_cols[1].metric("Target role match", f"{job_score}%" if target_skills else "—")
metric_cols[2].metric("Skills identified", len(resume_skills))
metric_cols[3].metric("Resume length", f"{resume_summary['word_count']:,} words")

if resume_summary["word_count"] < 250:
    st.warning("This resume is quite short. Check that all relevant sections and experience were extracted.")
elif resume_summary["word_count"] > 1000:
    st.info("This resume is longer than 1,000 words. Consider prioritizing the most relevant experience.")

tab_overview, tab_match, tab_insights, tab_report = st.tabs(
    ["Overview", "Role match", "Resume insights", "Recruiter snapshot & report"]
)

with tab_overview:
    overview_left, overview_right = st.columns([1.1, 1])
    with overview_left:
        st.markdown('<div class="section-label">Your profile at a glance</div>', unsafe_allow_html=True)
        st.subheader("A useful starting point")
        st.write(
            f"Detected **{len(resume_skills)} of {len(all_skills)}** skills across "
            f"{len(skills_by_category)} skill areas."
        )
        st.progress(overall_score / 100, text=f"Overall skill coverage · {overall_score}%")
        contact_items = [
            ("Email", resume_summary["email_present"]),
            ("Phone", resume_summary["phone_present"]),
        ]
        for label, present in contact_items:
            st.write(f"{'✓' if present else '○'} **{label}** {'detected' if present else 'not detected'}")
    with overview_right:
        st.markdown('<div class="section-label">Coverage by skill area</div>', unsafe_allow_html=True)
        st.bar_chart(
            {
                "Matched": [category_scores[name][0] for name in skills_by_category],
                "Remaining": [
                    category_scores[name][1] - category_scores[name][0]
                    for name in skills_by_category
                ],
            },
            horizontal=True,
            color=["#087e78", "#dce8ed"],
        )

    st.markdown('<div class="section-label">Your capabilities</div>', unsafe_allow_html=True)
    for category, skills in skills_by_category.items():
        found = found_by_category[category]
        absent = [skill for skill in skills if skill not in found]
        with st.expander(f"{category} · {len(found)} of {len(skills)} identified"):
            st.markdown(f"**Found:** {', '.join(found) if found else 'None identified'}")
            st.markdown(f"**Not detected:** {', '.join(absent) if absent else 'No gaps in this area'}")

with tab_match:
    st.markdown('<div class="section-label">A comparison, not a verdict</div>', unsafe_allow_html=True)
    st.subheader(target_label)
    if not target_skills:
        st.info("Choose a target role or paste a job description in the left panel to see a tailored match.")
    else:
        st.progress(job_score / 100, text=f"Skill overlap · {job_score}%")
        match_cols = st.columns(3)
        match_cols[0].metric("Role skills reviewed", len(target_skills))
        match_cols[1].metric("Found in resume", len(matched_skills))
        match_cols[2].metric("Not detected", len(missing_job_skills))
        st.markdown("**Matched skills**")
        st.success(", ".join(matched_skills) if matched_skills else "No target skills were detected in the resume.")
        st.markdown("**Prioritize these role skills**")
        if missing_job_skills:
            st.warning(", ".join(missing_job_skills))
            st.caption(
                "If you have this experience, add a specific example or achievement that demonstrates it. "
                "Do not add skills you cannot support."
            )
        else:
            st.success("Every skill detected in this target is also present in the resume.")

    if job_description.strip():
        if extra_keyword_suggestions:
            st.markdown("**Job-description terms not prominent in the resume**")
            st.write(", ".join(extra_keyword_suggestions))
            st.caption("Review these terms and naturally include them only where they accurately describe your work.")
        elif not target_skills:
            st.info("No library skills were identified in the pasted job description; check its text or choose a role profile.")
    elif selected_role == "Use job description":
        st.caption("For the most specific comparison, paste the full job description in the left panel.")

with tab_insights:
    insight_cols = st.columns(2)
    with insight_cols[0]:
        st.markdown('<div class="section-label">Document signals</div>', unsafe_allow_html=True)
        st.subheader("Resume essentials")
        st.write(f"**Word count:** {resume_summary['word_count']:,}")
        st.write(f"**Email address:** {'Detected' if resume_summary['email_present'] else 'Not detected'}")
        st.write(f"**Phone number:** {'Detected' if resume_summary['phone_present'] else 'Not detected'}")
        st.write(
            f"**Common sections found:** {', '.join(resume_summary['sections']) if resume_summary['sections'] else 'No common section headings detected'}"
        )
    with insight_cols[1]:
        st.markdown('<div class="section-label">Language in your resume</div>', unsafe_allow_html=True)
        st.subheader("Frequently used keywords")
        if resume_keywords:
            st.dataframe(
                [{"Keyword": word, "Mentions": count} for word, count in resume_keywords],
                hide_index=True,
                width="stretch",
            )
        else:
            st.write("No repeated keywords stood out in the extracted text.")

    st.markdown('<div class="section-label">Suggested next steps</div>', unsafe_allow_html=True)
    suggestions = []
    if missing_job_skills:
        suggestions.append(
            f"Review the target role's key skills: {', '.join(missing_job_skills[:5])}."
        )
    if not resume_summary["email_present"] or not resume_summary["phone_present"]:
        suggestions.append("Make your preferred contact details easy to find at the top of the resume.")
    if resume_summary["word_count"] < 250:
        suggestions.append("Check whether relevant experience, projects, or education details are missing.")
    if resume_summary["word_count"] > 1000:
        suggestions.append("Trim repeated or less relevant details to make the strongest evidence easier to scan.")
    if not suggestions:
        suggestions.append("Keep the resume focused on measurable achievements that support the target role.")
    for suggestion in suggestions:
        st.write(f"- {suggestion}")

with tab_report:
    st.markdown('<div class="section-label">A concise, evidence-led review</div>', unsafe_allow_html=True)
    st.subheader("Recruiter snapshot")
    if target_skills:
        snapshot = (
            f"For **{target_label.lower()}**, the resume contains **{len(matched_skills)} of "
            f"{len(target_skills)}** target skills ({job_score}% keyword overlap). "
        )
    else:
        snapshot = (
            f"The resume contains **{len(resume_skills)} identified skills** and "
            f"{resume_summary['word_count']:,} words. Add a target role for a role-specific match. "
        )
    contact_status = []
    if resume_summary["email_present"]:
        contact_status.append("email detected")
    if resume_summary["phone_present"]:
        contact_status.append("phone detected")
    snapshot += (
        f"Contact details: {', '.join(contact_status)}."
        if contact_status
        else "No email or phone number was detected in the extracted text."
    )
    st.info(snapshot)
    if resume_summary["sections"]:
        st.write(f"**Sections identified:** {', '.join(resume_summary['sections'])}")
    if matched_skills:
        st.write(f"**Relevant skills detected:** {', '.join(matched_skills)}")
    if missing_job_skills:
        st.write(f"**Skills to review:** {', '.join(missing_job_skills)}")
    st.caption(
        "This snapshot reflects text and keywords found in the uploaded file. "
        "It does not assess candidate quality or make hiring recommendations."
    )

    report_bytes = build_pdf_report(
        file_name,
        target_label,
        resume_summary,
        overall_score,
        resume_skills,
        target_skills,
        matched_skills,
        missing_job_skills,
        extra_keyword_suggestions,
        category_scores,
    )
    safe_filename = re.sub(r"[^a-zA-Z0-9_-]+", "-", os.path.splitext(file_name)[0]).strip("-")
    st.download_button(
        "Download your PDF report",
        data=report_bytes,
        file_name=f"{safe_filename or 'resume'}-northstar-report.pdf",
        mime="application/pdf",
        type="primary",
    )

with st.expander("Preview extracted resume text"):
    st.text_area("Extracted text", resume_text[:8000], height=260, label_visibility="collapsed")

st.markdown(
    '<p class="fine-print">Northstar Resume Intelligence · Your resume is processed in the app session and is not stored by this analyzer. '
    "Keyword matching is a guide, not a hiring decision.</p>",
    unsafe_allow_html=True,
)
