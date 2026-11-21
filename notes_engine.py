
import os
import re
import time
from typing import List, Tuple

from dotenv import load_dotenv
from groq import Groq

load_dotenv()

API_KEY = os.getenv("GROQ_API_KEY")
MODEL = os.getenv("GROQ_MODEL", "openai/gpt-oss-120b")


def split_into_units(syllabus: str) -> List[Tuple[str, str]]:
    """Detect and split syllabus into individual units."""

    pattern = re.compile(
        r"(?im)^\s*(unit|module)\s*[-:–]?\s*"
        r"([0-9]+|[ivxlcdm]+)\b[^\n]*"
    )

    matches = list(pattern.finditer(syllabus))

    if not matches:
        raise ValueError(
            "No units found. Use headings such as Unit 1, Unit 2 or Module 1."
        )

    units = []

    for index, match in enumerate(matches):
        start = match.end()
        end = (
            matches[index + 1].start()
            if index + 1 < len(matches)
            else len(syllabus)
        )

        heading = match.group(0).strip()
        content = syllabus[start:end].strip()

        if not content:
            raise ValueError(f"No topics found under {heading}.")

        units.append((heading, content))

    return units


def generate_unit_notes(
    unit_heading: str,
    syllabus_content: str,
    detail_level: str = "Balanced",
    language: str = "Simple English",
) -> str:
    """Generate exam-oriented notes for one unit using Groq."""

    if not API_KEY:
        raise RuntimeError(
            "Groq API key not found. Add GROQ_API_KEY to your .env file."
        )

    if not syllabus_content.strip():
        raise ValueError(f"No syllabus content found for {unit_heading}.")

    client = Groq(
        api_key=API_KEY,
        timeout=90.0,
        max_retries=2,
    )

    detail_instructions = {
        "Balanced": "Create balanced notes with clear explanations and examples.",
        "Short revision notes": "Create concise revision notes focusing on definitions, key points and important facts.",
        "Detailed exam notes": "Create detailed notes with thorough explanations, examples and exam-oriented coverage.",
    }

    selected_detail = detail_instructions.get(
        detail_level,
        detail_instructions["Balanced"]
    )

    prompt = f"""
Create high-quality university study notes.

UNIT: {unit_heading}

SYLLABUS:
{syllabus_content}

REQUIREMENTS:

1. Cover every topic and subtopic in the supplied syllabus.
2. Use {language} throughout the notes.
3. {selected_detail}
4. Explain concepts in simple, student-friendly language.
5. Include definitions and suitable examples.
6. Include advantages, disadvantages and applications wherever relevant.
7. Use headings, subheadings and bullet points.
8. Highlight important terminology.
9. Include MCQ-oriented facts and important points.
10. Add 5 short-answer questions and 3 long-answer questions.
11. Finish with a quick revision summary.
12. Do not invent unrelated syllabus topics.
13. Use Markdown formatting.

Make the notes accurate, organised and useful for university examinations.
"""

    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "system",
                "content": "You are an expert university professor and study-notes creator."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        temperature=0.3,
        max_completion_tokens=6000,
    )

    result = response.choices[0].message.content

    if not result or not result.strip():
        raise RuntimeError(
            f"Groq returned an empty response for {unit_heading}."
        )

    return result.strip()


def generate_all_notes(
    syllabus: str,
    progress_callback=None,
    detail_level: str = "Balanced",
    language: str = "Simple English",
) -> List[Tuple[str, str]]:
    """Generate notes for all detected syllabus units."""

    units = split_into_units(syllabus)
    all_notes = []

    for index, (heading, content) in enumerate(units, start=1):

        if progress_callback:
            progress_callback(index - 1, len(units), heading)

        notes = None
        last_error = None

        for attempt in range(3):
            try:
                notes = generate_unit_notes(
                    heading,
                    content,
                    detail_level,
                    language,
                )
                break

            except Exception as error:
                last_error = error

                if attempt < 2:
                    time.sleep(2 ** attempt)

        if notes is None:
            raise RuntimeError(
                f"Failed to generate notes for {heading}: {last_error}"
            )

        all_notes.append((heading, notes))

        if progress_callback:
            progress_callback(index, len(units), heading)

    return all_notes
