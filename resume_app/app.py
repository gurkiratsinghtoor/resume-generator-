import io
import re
from collections import Counter

import streamlit as st
from docx import Document
from pypdf import PdfReader

st.set_page_config(page_title="Resume Prep & Checker", page_icon="📄", layout="wide")

# ----------------------------- Config / word lists -----------------------------
SECTIONS = {
    "Summary / Objective": ["summary", "objective", "profile", "about me"],
    "Skills": ["skills", "technical skills", "core competencies"],
    "Experience": ["experience", "work history", "employment", "internship"],
    "Education": ["education", "academic"],
    "Projects": ["projects", "project work"],
    "Certifications / Awards": ["certifications", "certificates", "achievements", "awards"],
}

ACTION_VERBS = {
    "achieved", "analyzed", "automated", "built", "created", "collaborated", "delivered",
    "designed", "developed", "drove", "enhanced", "established", "implemented", "improved",
    "increased", "integrated", "launched", "led", "managed", "mentored", "optimized",
    "organized", "reduced", "resolved", "streamlined", "tested", "trained", "deployed",
    "architected", "migrated", "maintained", "coordinated", "planned", "wrote", "presented",
}

WEAK_PHRASES = [
    "responsible for", "worked on", "helped with", "duties included", "team player",
    "hard worker", "go-getter", "detail-oriented", "references available", "results-driven",
]

STOPWORDS = set("""a an the and or of to in for on with at by from as is are was were be been this that
these those it its you your we our they their will can may should must have has had not but if
than then so such into over under about across per using use used etc experience years year
work working ability able strong good excellent knowledge team role job candidate required
preferred responsibilities requirements skills including including plus new more other also""".split())

EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
PHONE_RE = re.compile(r"(\+?\d[\d\s().-]{8,}\d)")
LINK_RE = re.compile(r"(linkedin\.com|github\.com)", re.I)
BULLET_RE = re.compile(r"^\s*[-•*–▪●]\s*(.+)")


# ----------------------------- Helpers -----------------------------
def extract_text(uploaded):
    name = uploaded.name.lower()
    if name.endswith(".pdf"):
        reader = PdfReader(uploaded)
        return "\n".join((p.extract_text() or "") for p in reader.pages)
    if name.endswith(".docx"):
        doc = Document(uploaded)
        return "\n".join(p.text for p in doc.paragraphs)
    return uploaded.read().decode("utf-8", errors="ignore")


def get_bullets(text):
    return [m.group(1).strip() for line in text.splitlines() if (m := BULLET_RE.match(line))]


def check_resume(text):
    """Rule-based resume analysis. Returns (score, findings dict)."""
    low = text.lower()
    words = re.findall(r"\b\w+\b", text)
    n_words = len(words)
    bullets = get_bullets(text)
    score = 0
    good, warn = [], []

    # 1. Contact info (20)
    contact = {
        "Email": bool(EMAIL_RE.search(text)),
        "Phone": bool(PHONE_RE.search(text)),
        "LinkedIn / GitHub": bool(LINK_RE.search(text)),
    }
    pts = round(20 * sum(contact.values()) / 3)
    score += pts
    for k, v in contact.items():
        (good if v else warn).append(f"{k} {'found' if v else 'missing'}")

    # 2. Sections (30)
    found = {}
    for sec, keys in SECTIONS.items():
        found[sec] = any(re.search(rf"(^|\n)\s*{re.escape(k)}\b", low) for k in keys)
    core = ["Skills", "Experience", "Education"]
    core_pts = sum(10 for c in core if found[c])
    score += core_pts
    for sec, ok in found.items():
        if ok:
            good.append(f"Section present: {sec}")
        elif sec in core:
            warn.append(f"Missing important section: {sec}")

    # 3. Length (10)
    if 250 <= n_words <= 800:
        score += 10
        good.append(f"Good length ({n_words} words)")
    elif n_words < 250:
        score += 4
        warn.append(f"Too short ({n_words} words). Aim for 300-700.")
    else:
        score += 5
        warn.append(f"Quite long ({n_words} words). Try to keep it to 1-2 pages.")

    # 4. Action verbs (15)
    starts = [b.split()[0].lower().strip(",.") for b in bullets if b.split()]
    verb_hits = sum(1 for s in starts if s in ACTION_VERBS)
    if bullets:
        ratio = verb_hits / len(bullets)
        score += round(15 * min(ratio / 0.6, 1))
        if ratio >= 0.6:
            good.append(f"Strong action verbs ({verb_hits}/{len(bullets)} bullets)")
        else:
            warn.append(f"Only {verb_hits}/{len(bullets)} bullets start with an action verb "
                        "(e.g. Built, Led, Improved).")
    else:
        warn.append("No bullet points detected. Use bullets (- or •) for achievements.")

    # 5. Quantified achievements (15)
    quant = [b for b in bullets if re.search(r"\d|%|\$|₹", b)]
    if bullets:
        ratio = len(quant) / len(bullets)
        score += round(15 * min(ratio / 0.4, 1))
        if ratio >= 0.4:
            good.append(f"Measurable results in {len(quant)}/{len(bullets)} bullets")
        else:
            warn.append("Add numbers to achievements (e.g. 'reduced load time by 30%').")

    # 6. Weak phrases (10)
    weak = [p for p in WEAK_PHRASES if p in low]
    score += max(0, 10 - 3 * len(weak))
    if weak:
        warn.append("Weak/cliché phrases: " + ", ".join(f"'{w}'" for w in weak))
    else:
        good.append("No weak or cliché phrases")

    # Extra: personal pronouns
    if re.search(r"\b(i|my|me)\b", low):
        warn.append("Avoid first-person pronouns (I, my, me) in resumes.")

    return min(score, 100), {"good": good, "warn": warn, "words": n_words, "bullets": len(bullets)}


def keywords(text):
    toks = re.findall(r"[a-zA-Z][a-zA-Z+#.]{1,}", text.lower())
    return [t.strip(".") for t in toks if t.strip(".") not in STOPWORDS and len(t) > 2]


def keyword_match(resume, jd):
    jd_counts = Counter(keywords(jd))
    jd_top = [w for w, _ in jd_counts.most_common(30)]
    res_set = set(keywords(resume))
    matched = [w for w in jd_top if w in res_set]
    missing = [w for w in jd_top if w not in res_set]
    pct = round(100 * len(matched) / len(jd_top)) if jd_top else 0
    return pct, matched, missing


def build_resume_text(d):
    lines = [d["name"].upper()]
    contact = " | ".join(x for x in [d["email"], d["phone"], d["location"], d["links"]] if x)
    if contact:
        lines.append(contact)
    if d["summary"].strip():
        lines += ["", "SUMMARY", d["summary"].strip()]
    if d["skills"].strip():
        lines += ["", "SKILLS", d["skills"].strip()]
    if d["experience"]:
        lines += ["", "EXPERIENCE"]
        for e in d["experience"]:
            lines.append(f"{e['title']} - {e['company']} ({e['dates']})")
            lines += [f"- {b.strip().lstrip('-•* ')}" for b in e["bullets"].splitlines() if b.strip()]
            lines.append("")
    if d["projects"].strip():
        lines += ["PROJECTS"]
        lines += [f"- {b.strip().lstrip('-•* ')}" for b in d["projects"].splitlines() if b.strip()]
        lines.append("")
    if d["education"].strip():
        lines += ["EDUCATION", d["education"].strip()]
    if d["certs"].strip():
        lines += ["", "CERTIFICATIONS / ACHIEVEMENTS"]
        lines += [f"- {b.strip().lstrip('-•* ')}" for b in d["certs"].splitlines() if b.strip()]
    return "\n".join(lines).strip()


def build_docx(text):
    doc = Document()
    for i, line in enumerate(text.splitlines()):
        if not line.strip():
            continue
        if i == 0:
            doc.add_heading(line, level=0)
        elif line.isupper() and len(line) < 40:
            doc.add_heading(line.title(), level=2)
        elif line.startswith("- "):
            doc.add_paragraph(line[2:], style="List Bullet")
        else:
            doc.add_paragraph(line)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ----------------------------- UI -----------------------------
st.title("📄 Resume Preparation & Checking System")
st.caption("Build a clean resume, then check it for ATS-friendliness and job-description fit.")

tab_build, tab_check = st.tabs(["✍️ Build Resume", "✅ Check Resume"])

# ---------- Builder ----------
with tab_build:
    c1, c2 = st.columns([1, 1])
    with c1:
        st.subheader("Your details")
        name = st.text_input("Full name", "Your Name")
        a, b = st.columns(2)
        email = a.text_input("Email")
        phone = b.text_input("Phone")
        location = a.text_input("Location", "Khanna, Punjab")
        links = b.text_input("LinkedIn / GitHub")
        summary = st.text_area("Professional summary (2-3 lines)", height=90)
        skills = st.text_area("Skills (comma separated)", height=70,
                              placeholder="Python, SQL, Streamlit, Git, Communication")

        n_exp = st.number_input("Number of experiences / internships", 0, 6, 1)
        experience = []
        for i in range(int(n_exp)):
            with st.expander(f"Experience {i + 1}", expanded=(i == 0)):
                t = st.text_input("Job title", key=f"t{i}")
                co = st.text_input("Company", key=f"c{i}")
                dt = st.text_input("Dates", key=f"d{i}", placeholder="Jun 2024 - Present")
                bl = st.text_area("Achievements (one per line)", key=f"b{i}", height=100,
                                  placeholder="Built a dashboard that cut reporting time by 40%")
                experience.append({"title": t, "company": co, "dates": dt, "bullets": bl})

        projects = st.text_area("Projects (one per line)", height=90)
        education = st.text_area("Education", height=80,
                                 placeholder="B.Tech CSE - XYZ University (2021-2025), CGPA 8.2")
        certs = st.text_area("Certifications / Achievements (one per line)", height=70)

    data = dict(name=name, email=email, phone=phone, location=location, links=links,
                summary=summary, skills=skills, experience=experience, projects=projects,
                education=education, certs=certs)
    resume_text = build_resume_text(data)
    st.session_state["built_text"] = resume_text

    with c2:
        st.subheader("Preview")
        st.code(resume_text, language=None)
        d1, d2 = st.columns(2)
        d1.download_button("⬇️ Download .txt", resume_text, "resume.txt")
        d2.download_button("⬇️ Download .docx", build_docx(resume_text), "resume.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ---------- Checker ----------
with tab_check:
    st.subheader("Resume checker")
    src = st.radio("Resume source", ["Upload file", "Paste text", "Use resume from Builder tab"],
                   horizontal=True)
    text = ""
    if src == "Upload file":
        up = st.file_uploader("PDF, DOCX or TXT", type=["pdf", "docx", "txt"])
        if up:
            text = extract_text(up)
    elif src == "Paste text":
        text = st.text_area("Paste resume text", height=250)
    else:
        text = st.session_state.get("built_text", "")

    jd = st.text_area("Job description (optional, for keyword match)", height=150)

    if st.button("Analyze", type="primary"):
        if len(text.strip()) < 50:
            st.warning("Please provide a resume with some content first.")
        else:
            score, res = check_resume(text)
            m1, m2, m3 = st.columns(3)
            m1.metric("Resume score", f"{score}/100")
            m2.metric("Words", res["words"])
            m3.metric("Bullet points", res["bullets"])
            st.progress(score / 100)

            g, w = st.columns(2)
            with g:
                st.markdown("#### ✅ What's good")
                for x in res["good"]:
                    st.write("•", x)
            with w:
                st.markdown("#### ⚠️ Improve")
                for x in res["warn"] or ["Nothing major. Nice work!"]:
                    st.write("•", x)

            if jd.strip():
                pct, matched, missing = keyword_match(text, jd)
                st.markdown("---")
                st.markdown(f"#### 🎯 Job description match: **{pct}%**")
                st.progress(pct / 100)
                k1, k2 = st.columns(2)
                k1.markdown("**Matched keywords**")
                k1.write(", ".join(matched) or "None")
                k2.markdown("**Missing keywords (add if truthful)**")
                k2.write(", ".join(missing) or "None")
