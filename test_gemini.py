
from notes_engine import generate_all_notes

sample_syllabus = """
Subject: Generative AI

Unit 1: Introduction to Generative AI
Introduction to Artificial Intelligence
Generative AI concepts
Applications of Generative AI
Types of generative models

Unit 2: Large Language Models
Introduction to LLMs
Transformer architecture
Attention mechanism
Applications of LLMs
"""

try:
    print("Starting unit-wise notes generation...\n")

    def show_progress(done, total, heading):
        print(f"Progress: {done}/{total} | {heading}")

    notes = generate_all_notes(
        sample_syllabus,
        progress_callback=show_progress,
        detail_level="Balanced",
        language="Simple English"
    )

    print("\nNOTES GENERATED SUCCESSFULLY!")
    print("=" * 50)

    for heading, content in notes:
        print(f"\n{heading}\n")
        print(content)
        print("\n" + "=" * 50)

except Exception as error:
    print("\nGENERATION FAILED!")
    print(type(error).__name__)
    print(error)
