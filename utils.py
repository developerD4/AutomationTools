import pandas as pd
from docx import Document
from docx.shared import RGBColor
from spellchecker import SpellChecker

def get_font_properties(run):
    return {
        "font_family": run.font.name,
        "font_size": run.font.size.pt if run.font.size else None,
        "font_color": run.font.color.rgb if run.font.color and run.font.color.rgb else None,
        "bold": run.font.bold
    }

def expected_props(level):
    if level == 1:
        return {"font_family": "Times New Roman", "font_size": 16, "font_color": RGBColor(0, 0, 0), "bold": True}
    elif level == 2:
        return {"font_family": "Times New Roman", "font_size": 14, "font_color": RGBColor(0, 0, 0), "bold": True}
    elif level in [3, 4, 5]:
        return {"font_family": "Times New Roman", "font_size": 12, "font_color": RGBColor(0, 0, 0), "bold": True}
    return {}

def run_tests_on_doc(doc_path):
    document = Document(doc_path)
    result_data = []

    spell = SpellChecker()
    current_page = 1  # Placeholder

    for idx, para in enumerate(document.paragraphs):
        para_text = para.text.strip()

        # === Heading style validation ===
        if para.style.name.lower().startswith("heading"):
            try:
                level = int(para.style.name.split()[-1])
            except ValueError:
                continue

            run = para.runs[0] if para.runs else None
            if not run:
                result_data.append({
                    "Test Case Name": f"Heading Level {level} Font",
                    "Test Case Description": f"Validate font attributes for Heading Level {level}",
                    "Category": f"Heading Level {level}",
                    "Page Number": current_page,
                    "Text": para_text,
                    "Property": "Run",
                    "Expected": "Font run exists",
                    "Actual": "No font run",
                    "Status": "FAIL"
                })
                continue

            actual = get_font_properties(run)
            expected = expected_props(level)

            for prop, expected_val in expected.items():
                actual_val = actual.get(prop)
                result_data.append({
                    "Test Case Name": f"Heading Level {level} Font",
                    "Test Case Description": f"Validate font attributes for Heading Level {level}",
                    "Category": f"Heading Level {level}",
                    "Page Number": current_page,
                    "Text": para_text,
                    "Property": prop,
                    "Expected": str(expected_val),
                    "Actual": str(actual_val),
                    "Status": "PASS" if actual_val == expected_val else "FAIL"
                })

            if level == 6 and ("page" in run.text.lower() or any(char.isdigit() for char in run.text)):
                result_data.append({
                    "Test Case Name": "Page Number Location",
                    "Test Case Description": "Check if page number is placed on top-left in Heading Level 6",
                    "Category": "Page Number Check",
                    "Page Number": current_page,
                    "Text": para_text,
                    "Property": "Content",
                    "Expected": "Page number label present",
                    "Actual": "Valid" if run.text.strip() else "Empty",
                    "Status": "PASS" if run.text.strip() else "FAIL"
                })

        # === Spelling Check ===
        words = para_text.split()
        misspelled = spell.unknown(words)
        for word in misspelled:
            suggestion = spell.correction(word)
            result_data.append({
                "Test Case Name": "Spelling Check",
                "Test Case Description": "Validate if there are spelling mistakes in the paragraph",
                "Category": "Spelling Check",
                "Page Number": current_page,
                "Text": para_text,
                "Property": "Spelling",
                "Expected": suggestion,
                "Actual": word,
                "Status": "FAIL"
            })

    df = pd.DataFrame(result_data)
    excel_path = "test_report_detailed.xlsx"
    df.to_excel(excel_path, index=False)
    return excel_path
