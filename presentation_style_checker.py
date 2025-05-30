import streamlit as st
from pptx import Presentation
import re

# Style rules
STYLE_RULES = {
    "font": "Arial",
    "font_size": 24,
    "character_limit": 300,
    "forbidden_spellings": ["z. B.", "Du"],
    "required_spellings": {
        "z. B.": "z.B.",
        "Du": "du"
    }
}

def normalize_text(text):
    return re.sub(r"\s+", " ", text.strip()).lower()

def check_slide(slide, slide_number):
    results = []
    character_count = 0

    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text = run.text.strip()
                    normalized_text = normalize_text(text)
                    character_count += len(text)

                    # Debug: Print found text
                    print(f"[Slide {slide_number}] Found text: '{text}'")

                    # Font checks
                    font = run.font.name or "Unknown"
                    size = run.font.size.pt if run.font.size else None

                    if font != STYLE_RULES["font"]:
                        results.append([slide_number, "Font", font])
                    if size and int(size) != STYLE_RULES["font_size"]:
                        results.append([slide_number, "Font Size", f"{size:.1f} pt"])

                    # Forbidden spellings
                    for forbidden in STYLE_RULES["forbidden_spellings"]:
                        if forbidden.lower() in normalized_text:
                            results.append([slide_number, "Forbidden Spelling", forbidden])

                    # Required spellings (wrong usage)
                    for wrong, correct in STYLE_RULES["required_spellings"].items():
                        if wrong.lower() in normalized_text and correct.lower() not in normalized_text:
                            results.append([slide_number, "Incorrect Spelling", wrong])

    if character_count > STYLE_RULES["character_limit"]:
        results.append([slide_number, "Character Limit", f"{character_count} characters"])

    return results

def main():
    st.title("🔍 PowerPoint Style Checker")
    uploaded_file = st.file_uploader("Upload a PowerPoint file (.pptx)", type="pptx")

    if uploaded_file:
        prs = Presentation(uploaded_file)
        all_results = []

        for i, slide in enumerate(prs.slides, start=1):
            st.write(f"📄 Checking Slide {i}...")
            all_results.extend(check_slide(slide, i))

        if all_results:
            st.success(f"✅ Checked {len(prs.slides)} slides. Issues found:")
            st.table(all_results)
        else:
            st.success("✅ All slides passed the style check! 🎉")

if __name__ == "__main__":
    main()
