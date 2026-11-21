import io
import re
import html
from datetime import date

import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from notes_engine import generate_all_notes


# ==================================================
# 1. PAGE CONFIGURATION
# ==================================================

st.set_page_config(
    page_title="Syllabus AI | Study notes from your syllabus",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ==================================================
# 2. SESSION STATE
# ==================================================

st.session_state.setdefault("generated_notes", None)
st.session_state.setdefault("syllabus_input", "")
st.session_state.setdefault("detail_level", "Balanced")
st.session_state.setdefault("language", "Simple English")
st.session_state.setdefault("subject", "")


# ==================================================
# 3. DESIGN SYSTEM
# ==================================================
# Idea: a student's notebook. Ruled paper, a red margin line,
# blue ink and a yellow highlighter. The syllabus box is the
# notebook page, everything else stays quiet.

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,500;12..96,700;12..96,800&family=Figtree:wght@400;500;600;700&display=swap');

:root {
    --paper: #F5F7FB;
    --card: #FFFFFF;
    --ink: #16213E;
    --muted: #5B6785;
    --line: #DCE3F0;
    --rule: #E3EAF6;
    --blue: #2F4BDC;
    --blue-deep: #2038B8;
    --blue-soft: #E8ECFD;
    --highlight: #FFE27A;
    --margin: #E8766B;
    --done: #1F8F6B;
    --display: 'Bricolage Grotesque', 'Segoe UI', sans-serif;
    --body: 'Figtree', 'Segoe UI', system-ui, sans-serif;
}

/* ---------- Base ---------- */

.stApp {
    background: var(--paper);
    color: var(--ink);
    font-family: var(--body);
}

.stApp p, .stApp li, .stApp label, .stApp textarea,
.stApp button, .stApp input {
    font-family: var(--body);
}

.block-container {
    max-width: 1080px;
    padding-top: 3rem;
    padding-bottom: 4rem;
}

#MainMenu, footer { visibility: hidden; }

header[data-testid="stHeader"] { background: transparent; }

.stApp h1, .stApp h2, .stApp h3, .stApp h4 {
    font-family: var(--display);
    color: var(--ink);
    letter-spacing: -0.02em;
}

/* ---------- Hero ---------- */

.sa-brand {
    display: inline-flex;
    align-items: center;
    gap: 10px;
    font-family: var(--display);
    font-weight: 700;
    font-size: 18px;
    color: var(--ink);
    margin-bottom: 34px;
}

.sa-brand-mark {
    width: 30px;
    height: 30px;
    border-radius: 8px;
    background: var(--blue);
    color: #fff;
    display: grid;
    place-items: center;
    font-size: 16px;
}

.sa-hero {
    display: grid;
    grid-template-columns: 1.05fr 0.95fr;
    gap: 56px;
    align-items: center;
    margin-bottom: 38px;
}

.sa-hero h1 {
    font-size: clamp(34px, 5vw, 54px);
    font-weight: 800;
    line-height: 1.08;
    letter-spacing: -0.035em;
    margin: 0 0 18px 0;
    padding: 0;
}

.sa-hero p {
    color: var(--muted);
    font-size: 17px;
    line-height: 1.7;
    max-width: 52ch;
    margin: 0;
}

.sa-page {
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 6px;
    box-shadow: 0 22px 44px rgba(22, 33, 62, 0.10);
    transform: rotate(1.3deg);
    padding: 18px 0 18px 0;
}

.sa-page-in {
    position: relative;
    padding: 0 18px 0 58px;
    background-image: repeating-linear-gradient(
        transparent, transparent 27px, var(--rule) 27px, var(--rule) 28px
    );
}

.sa-page-in::before {
    content: "";
    position: absolute;
    top: -18px;
    bottom: -18px;
    left: 44px;
    width: 1px;
    background: var(--margin);
}

.sa-row {
    height: 28px;
    line-height: 28px;
    font-size: 15px;
    color: var(--ink);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

.sa-row.small {
    font-size: 13px;
    color: var(--muted);
}

.sa-row mark {
    background: var(--highlight);
    color: var(--ink);
    padding: 1px 3px;
    border-radius: 2px;
}

/* ---------- Progress steps (a real sequence) ---------- */

.sa-steps {
    display: flex;
    margin: 6px 0 8px 0;
}

.sa-step {
    flex: 1;
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 14px;
    font-weight: 600;
    color: var(--muted);
    position: relative;
}

.sa-step:not(:last-child)::after {
    content: "";
    flex: 1;
    height: 2px;
    margin: 0 14px;
    background: var(--line);
}

.sa-step.done:not(:last-child)::after { background: var(--blue); }

.sa-dot {
    width: 28px;
    height: 28px;
    border-radius: 50%;
    border: 2px solid var(--line);
    background: var(--card);
    display: grid;
    place-items: center;
    font-size: 13px;
    flex-shrink: 0;
}

.sa-step.active { color: var(--blue); }
.sa-step.active .sa-dot { border-color: var(--blue); color: var(--blue); }

.sa-step.done { color: var(--ink); }
.sa-step.done .sa-dot {
    background: var(--blue);
    border-color: var(--blue);
    color: #fff;
}

/* ---------- Section headings ---------- */

.sa-h2 {
    font-family: var(--display);
    font-size: 26px;
    font-weight: 700;
    letter-spacing: -0.02em;
    color: var(--ink);
    margin: 46px 0 6px 0;
}

.sa-desc {
    color: var(--muted);
    font-size: 15px;
    line-height: 1.65;
    max-width: 64ch;
    margin-bottom: 18px;
}

/* ---------- Syllabus box (notebook page) ---------- */

.stTextArea [data-baseweb="textarea"] {
    background: var(--card) !important;
    border: 1px solid var(--line) !important;
    border-radius: 14px !important;
    overflow: hidden;
    box-shadow: 0 6px 20px rgba(22, 33, 62, 0.05);
}

.stTextArea [data-baseweb="textarea"]:focus-within {
    border-color: var(--blue) !important;
    box-shadow: 0 0 0 3px rgba(47, 75, 220, 0.14);
}

.stTextArea textarea {
    background-color: transparent !important;
    background-image:
        linear-gradient(90deg, transparent 44px, var(--margin) 44px, var(--margin) 45px, transparent 45px),
        repeating-linear-gradient(transparent, transparent 27px, var(--rule) 27px, var(--rule) 28px) !important;
    background-attachment: local !important;
    color: var(--ink) !important;
    font-size: 15px !important;
    line-height: 28px !important;
    padding: 14px 20px 14px 62px !important;
}

.stTextArea textarea::placeholder {
    color: #9AA3BC !important;
    opacity: 1 !important;
}

/* ---------- Detected units ---------- */

.sa-found {
    margin-top: 14px;
    font-size: 14px;
    color: var(--muted);
}

.sa-found strong { color: var(--ink); }

.sa-chips {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-top: 10px;
}

.sa-chip {
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 999px;
    padding: 6px 13px;
    font-size: 13px;
    color: var(--ink);
}

.sa-chip small {
    color: var(--muted);
    margin-left: 6px;
    font-size: 12px;
}

.sa-hint {
    margin-top: 14px;
    padding: 12px 14px;
    border-left: 3px solid var(--highlight);
    background: #FFFBEA;
    border-radius: 0 10px 10px 0;
    font-size: 14px;
    color: #5C4B00;
}

/* ---------- Select boxes ---------- */

[data-testid="stWidgetLabel"] p {
    color: var(--ink) !important;
    font-weight: 600 !important;
    font-size: 14px !important;
}

.stSelectbox [data-baseweb="select"] > div {
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 10px;
    min-height: 48px;
    color: var(--ink);
}

.stSelectbox [data-baseweb="select"] > div:focus-within {
    border-color: var(--blue);
    box-shadow: 0 0 0 3px rgba(47, 75, 220, 0.14);
}

/* ---------- Buttons ---------- */

.stButton button,
.stDownloadButton button {
    border-radius: 10px !important;
    font-weight: 600 !important;
    min-height: 46px;
    transition: background-color 0.15s ease, border-color 0.15s ease;
}

.stButton button p,
.stDownloadButton button p {
    color: inherit !important;
    font-weight: 600 !important;
}

.stButton button:focus-visible,
.stDownloadButton button:focus-visible {
    outline: 3px solid rgba(47, 75, 220, 0.35) !important;
    outline-offset: 2px;
}

.stButton button:is([kind="primary"], [data-testid="stBaseButton-primary"]) {
    background: var(--blue) !important;
    color: #fff !important;
    border: 1px solid var(--blue) !important;
    width: 100%;
    min-height: 54px;
    font-size: 16px !important;
}

.stButton button:is([kind="primary"], [data-testid="stBaseButton-primary"]):hover {
    background: var(--blue-deep) !important;
    border-color: var(--blue-deep) !important;
}

.stButton button:is([kind="secondary"], [data-testid="stBaseButton-secondary"]) {
    background: var(--card) !important;
    color: var(--ink) !important;
    border: 1px solid var(--line) !important;
}

.stButton button:is([kind="secondary"], [data-testid="stBaseButton-secondary"]):hover {
    background: var(--blue-soft) !important;
    border-color: var(--blue) !important;
    color: var(--blue-deep) !important;
}

.stDownloadButton button {
    width: 100%;
    background: var(--card) !important;
    color: var(--ink) !important;
    border: 1px solid var(--line) !important;
}

.stDownloadButton button:hover {
    background: var(--blue-soft) !important;
    border-color: var(--blue) !important;
    color: var(--blue-deep) !important;
}

/* ---------- Summary strip ---------- */

.sa-summary {
    display: flex;
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 14px;
    margin-bottom: 20px;
    overflow: hidden;
}

.sa-summary > div {
    flex: 1;
    padding: 16px 20px;
}

.sa-summary > div + div { border-left: 1px solid var(--line); }

.sa-summary span {
    display: block;
    font-size: 13px;
    color: var(--muted);
    margin-bottom: 2px;
}

.sa-summary b {
    font-family: var(--display);
    font-size: 21px;
    font-weight: 700;
    color: var(--ink);
}

/* ---------- Generated notes ---------- */

[data-testid="stExpander"] {
    background: var(--card);
    border: 1px solid var(--line) !important;
    border-radius: 14px;
    margin-bottom: 12px;
    overflow: hidden;
}

[data-testid="stExpander"] summary {
    padding: 14px 18px;
}

[data-testid="stExpander"] summary p {
    font-family: var(--display);
    font-size: 18px;
    font-weight: 700;
    color: var(--ink);
}

[data-testid="stExpander"] [data-testid="stMarkdownContainer"] {
    font-size: 15.5px;
    line-height: 1.75;
    color: #26314F;
}

[data-testid="stExpander"] [data-testid="stMarkdownContainer"] h2,
[data-testid="stExpander"] [data-testid="stMarkdownContainer"] h3 {
    font-size: 19px;
    margin-top: 1.2em;
}

[data-testid="stExpander"] strong {
    background: linear-gradient(transparent 62%, rgba(255, 226, 122, 0.8) 62%);
    color: var(--ink);
}

/* ---------- Misc ---------- */

div[data-testid="stAlert"] { border-radius: 10px; }

hr { border-color: var(--line) !important; }

.sa-footer {
    text-align: center;
    color: #8590AA;
    font-size: 13px;
    line-height: 1.8;
    padding-top: 48px;
}

/* ---------- Responsive ---------- */

@media (max-width: 860px) {
    .sa-hero { grid-template-columns: 1fr; gap: 34px; }
    .sa-page { transform: none; }
    .sa-step span.sa-label { display: none; }
    .sa-summary { flex-direction: column; }
    .sa-summary > div + div { border-left: none; border-top: 1px solid var(--line); }
}

@media (max-width: 600px) {
    .block-container { padding-left: 1rem; padding-right: 1rem; padding-top: 2rem; }
    .sa-h2 { font-size: 22px; margin-top: 36px; }
}

@media (prefers-reduced-motion: reduce) {
    * { transition: none !important; }
}
</style>
""",
    unsafe_allow_html=True,
)


# ==================================================
# 4. HELPERS
# ==================================================

SAMPLE_SYLLABUS = """Subject: Generative Artificial Intelligence

Unit 1: Introduction to Generative AI
- Introduction to Artificial Intelligence
- Meaning and concepts of Generative AI
- Traditional AI vs Generative AI
- Applications of Generative AI
- Types of generative models

Unit 2: Large Language Models
- Introduction to Large Language Models
- Transformer architecture
- Attention mechanism
- Encoder and decoder
- Applications of LLMs

Unit 3: Prompt Engineering
- Introduction to prompt engineering
- Zero-shot prompting
- Few-shot prompting
- Chain-of-thought prompting
- Prompt design techniques
"""

UNIT_RE = re.compile(
    r"^\s*(unit|module|chapter)\s*[-:.]?\s*([0-9]+|[ivxIVX]+)\b\s*[:\-–.]?\s*(.*)$",
    re.IGNORECASE,
)


def analyse_syllabus(text):
    """Quick preview only: finds the subject and unit headings."""
    subject = ""
    units = []

    for raw in text.splitlines():
        line = raw.strip()

        if not line:
            continue

        if line.lower().startswith("subject"):
            subject = line.split(":", 1)[-1].strip()
            continue

        match = UNIT_RE.match(line)

        if match:
            word = match.group(1).capitalize()
            number = match.group(2)
            title = match.group(3).strip()
            units.append(
                {
                    "label": f"{word} {number}",
                    "title": title,
                    "topics": 0,
                }
            )
        elif units:
            units[-1]["topics"] += 1

    return subject, units


def load_sample():
    st.session_state.syllabus_input = SAMPLE_SYLLABUS


def clear_all():
    st.session_state.syllabus_input = ""
    st.session_state.generated_notes = None


def render_steps(current):
    steps = ["Paste syllabus", "Choose style", "Get your notes"]
    parts = []

    for number, label in enumerate(steps, start=1):
        if number < current:
            state, mark = "done", "✓"
        elif number == current:
            state, mark = "active", str(number)
        else:
            state, mark = "", str(number)

        parts.append(
            f'<div class="sa-step {state}">'
            f'<div class="sa-dot">{mark}</div>'
            f'<span class="sa-label">{label}</span>'
            f"</div>"
        )

    st.markdown(
        f'<div class="sa-steps">{"".join(parts)}</div>',
        unsafe_allow_html=True,
    )


def section(title, description):
    st.markdown(
        f'<div class="sa-h2">{title}</div>'
        f'<div class="sa-desc">{description}</div>',
        unsafe_allow_html=True,
    )


# ==================================================
# 5. HERO
# ==================================================

st.markdown(
    """
<div class="sa-brand"><div class="sa-brand-mark">✎</div>Syllabus AI</div>
<div class="sa-hero">
<div>
<h1>Turn your syllabus into notes you can revise from</h1>
<p>Paste the syllabus from your university. Syllabus AI writes clear notes for every unit, in simple English, Hinglish or Hindi, and gives you a Word file to keep.</p>
</div>
<div class="sa-page" aria-hidden="true">
<div class="sa-page-in">
<div class="sa-row small">From your syllabus</div>
<div class="sa-row">Unit 2: Large Language Models</div>
<div class="sa-row">Transformer architecture, attention</div>
<div class="sa-row small">Your notes</div>
<div class="sa-row"><b>Transformer:</b> reads a whole sentence at once.</div>
<div class="sa-row"><mark>Attention</mark> lets each word find the words that matter.</div>
</div>
</div>
</div>
""",
    unsafe_allow_html=True,
)


# ==================================================
# 6. PROGRESS STEPS
# ==================================================

if st.session_state.generated_notes:
    current_step = 4
elif st.session_state.syllabus_input.strip():
    current_step = 2
else:
    current_step = 1

render_steps(current_step)


# ==================================================
# 7. SYLLABUS INPUT
# ==================================================

section(
    "Paste your syllabus",
    "Copy it from your university PDF or portal. Unit headings like "
    "“Unit 1: Introduction” help us keep your notes organised.",
)

btn1, btn2, _spacer = st.columns([1.2, 1, 4])

with btn1:
    st.button("Load sample syllabus", on_click=load_sample)

with btn2:
    st.button("Clear", on_click=clear_all)

syllabus = st.text_area(
    "Syllabus text",
    height=300,
    key="syllabus_input",
    label_visibility="collapsed",
    placeholder="""Subject: Example Subject

Unit 1: Introduction
- Topic 1
- Topic 2

Unit 2: Core Concepts
- Topic 3
- Topic 4

Paste your complete syllabus here...""",
)

# Live preview of what we found
if syllabus.strip():
    subject_preview, units_preview = analyse_syllabus(syllabus)

    if units_preview:
        total_topics = sum(u["topics"] for u in units_preview)
        subject_text = (
            f" for <strong>{html.escape(subject_preview)}</strong>"
            if subject_preview
            else ""
        )

        chips = ""

        for unit in units_preview:
            name = unit["title"] or unit["label"]
            name = name if len(name) <= 36 else name[:35] + "…"
            chips += (
                f'<span class="sa-chip">{html.escape(unit["label"])}: '
                f"{html.escape(name)}"
                f'<small>{unit["topics"]} lines</small></span>'
            )

        st.markdown(
            f'<div class="sa-found">Found <strong>{len(units_preview)} '
            f"units</strong>{subject_text}, about {total_topics} topic lines."
            f'</div><div class="sa-chips">{chips}</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="sa-hint">No unit headings found. Add lines like '
            "“Unit 1: Title” so each unit gets its own notes.</div>",
            unsafe_allow_html=True,
        )


# ==================================================
# 8. SETTINGS
# ==================================================

section(
    "Choose how the notes read",
    "Pick the length and the language that suit how you study.",
)

col1, col2 = st.columns(2)

with col1:
    detail_level = st.selectbox(
        "Detail level",
        ["Balanced", "Short revision notes", "Detailed exam notes"],
    )

with col2:
    language = st.selectbox(
        "Language",
        ["Simple English", "Hinglish", "Hindi"],
    )


# ==================================================
# 9. WORD DOCUMENT GENERATOR
# ==================================================

INK = RGBColor(22, 33, 62)
BLUE = RGBColor(47, 75, 220)
BODY = RGBColor(38, 49, 79)
MUTED = RGBColor(91, 103, 133)


def add_runs(paragraph, text):
    """Adds text to a paragraph and turns **bold** into real bold."""
    for part in re.split(r"(\*\*[^*]+\*\*)", text):
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part:
            paragraph.add_run(part)

    return paragraph


def create_docx(notes, subject="", detail="", lang=""):

    document = Document()

    section_ = document.sections[0]
    section_.top_margin = Inches(0.8)
    section_.bottom_margin = Inches(0.8)
    section_.left_margin = Inches(0.9)
    section_.right_margin = Inches(0.9)

    styles = document.styles

    styles["Normal"].font.name = "Calibri"
    styles["Normal"].font.size = Pt(11)
    styles["Normal"].font.color.rgb = BODY

    for level, size in ((1, 20), (2, 15), (3, 12.5)):
        heading_style = styles[f"Heading {level}"]
        heading_style.font.name = "Calibri"
        heading_style.font.size = Pt(size)
        heading_style.font.bold = True
        heading_style.font.color.rgb = BLUE if level == 1 else INK

    # Cover page
    for _ in range(6):
        document.add_paragraph()

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run(subject or "Study Notes")
    title_run.bold = True
    title_run.font.size = Pt(32)
    title_run.font.color.rgb = INK

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_run = subtitle.add_run("Unit-wise study notes")
    subtitle_run.font.size = Pt(16)
    subtitle_run.font.color.rgb = BLUE

    info_bits = [f"{len(notes)} units"]

    if detail:
        info_bits.append(detail)

    if lang:
        info_bits.append(lang)

    info_bits.append(date.today().strftime("%d %B %Y"))

    info = document.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_run = info.add_run("   |   ".join(info_bits))
    info_run.font.size = Pt(10.5)
    info_run.font.color.rgb = MUTED

    document.add_page_break()

    for heading, content in notes:

        document.add_heading(heading, level=1)

        for line in content.splitlines():

            line = line.strip()

            if not line:
                continue

            if line.startswith("### "):
                document.add_heading(line[4:].strip(), level=3)

            elif line.startswith("## "):
                document.add_heading(line[3:].strip(), level=2)

            elif line.startswith("# "):
                document.add_heading(line[2:].strip(), level=2)

            elif line.startswith(("- ", "* ")):
                add_runs(
                    document.add_paragraph(style="List Bullet"),
                    line[2:].strip(),
                )

            elif re.match(r"^\d+[\.\)]\s+", line):
                add_runs(
                    document.add_paragraph(style="List Number"),
                    re.sub(r"^\d+[\.\)]\s+", "", line),
                )

            else:
                add_runs(document.add_paragraph(), line)

        document.add_page_break()

    output = io.BytesIO()
    document.save(output)
    output.seek(0)

    return output.getvalue()


def create_markdown(notes, subject=""):
    parts = [f"# {subject or 'Study Notes'}"]

    for heading, content in notes:
        parts.append(f"## {heading}\n\n{content.strip()}")

    return "\n\n---\n\n".join(parts)


# ==================================================
# 10. GENERATE
# ==================================================

section(
    "Generate your notes",
    "Each unit is written separately, so longer syllabi can take a minute or two.",
)

if st.button("Generate my notes", type="primary"):

    if not syllabus.strip():

        st.warning("Paste your syllabus above, then try again.")

    else:

        try:

            progress_bar = st.progress(0)
            status = st.empty()

            def update_progress(done, total, heading):

                percentage = int(done / total * 100)

                progress_bar.progress(percentage)

                status.info(
                    f"Writing notes for {heading} "
                    f"({done} of {total} units done)"
                )

            with st.spinner("Writing your notes..."):

                notes = generate_all_notes(
                    syllabus,
                    progress_callback=update_progress,
                    detail_level=detail_level,
                    language=language,
                )

            subject_name, _units = analyse_syllabus(syllabus)

            st.session_state.generated_notes = notes
            st.session_state.detail_level = detail_level
            st.session_state.language = language
            st.session_state.subject = subject_name

            progress_bar.progress(100)

            status.success("Your notes are ready. Scroll down to read them.")

        except Exception as error:

            st.error(
                f"We couldn't generate your notes: {error}. "
                "Check your connection and API key, then try again."
            )


# ==================================================
# 11. RESULTS
# ==================================================

if st.session_state.generated_notes:

    notes = st.session_state.generated_notes
    subject_name = st.session_state.get("subject", "")

    section(
        "Your notes",
        "Open a unit to read it. Key terms are highlighted for quick revision.",
    )

    st.markdown(
        f"""
<div class="sa-summary">
<div><span>Units</span><b>{len(notes)}</b></div>
<div><span>Detail</span><b>{html.escape(st.session_state.get("detail_level", "Balanced"))}</b></div>
<div><span>Language</span><b>{html.escape(st.session_state.get("language", "Simple English"))}</b></div>
</div>
""",
        unsafe_allow_html=True,
    )

    for index, (heading, content) in enumerate(notes, start=1):

        with st.expander(heading, expanded=(index == 1)):

            st.markdown(content)

    section(
        "Keep your notes",
        "Download them to revise offline or print.",
    )

    dl1, dl2 = st.columns(2)

    with dl1:
        st.download_button(
            label="Download as Word (.docx)",
            data=create_docx(
                notes,
                subject_name,
                st.session_state.get("detail_level", ""),
                st.session_state.get("language", ""),
            ),
            file_name="Syllabus_AI_Unit_Wise_Notes.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
        )

    with dl2:
        st.download_button(
            label="Download as Markdown (.md)",
            data=create_markdown(notes, subject_name),
            file_name="Syllabus_AI_Unit_Wise_Notes.md",
            mime="text/markdown",
        )


# ==================================================
# 12. FOOTER
# ==================================================

st.markdown(
    """
<div class="sa-footer">
<div>Syllabus AI, powered by Groq</div>
<div>Notes are AI-generated. Check important facts against your textbook.</div>
</div>
""",
    unsafe_allow_html=True,
)