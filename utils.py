import pandas as pd
from docx import Document
from docx.shared import RGBColor
import win32com.client
# from spellchecker import SpellChecker  # Uncomment if using spelling check

def rgb_to_hex(rgb_color):
    """Convert RGBColor to hex string for comparison"""
    if not rgb_color:
        return None
    return "{:02X}{:02X}{:02X}".format(rgb_color[0], rgb_color[1], rgb_color[2])

def get_font_properties(run, para):
    para_style_font = para.style.font

    def get_value(prop_name):
        run_val = getattr(run.font, prop_name, None)
        para_val = getattr(para_style_font, prop_name, None)
        return run_val if run_val is not None else para_val

    # Font color handling
    run_color = run.font.color.rgb if run.font.color and run.font.color.rgb else None
    para_color = para_style_font.color.rgb if para_style_font.color and para_style_font.color.rgb else None
    final_color = run_color if run_color else para_color

    return {
        "font_family": get_value('name'),
        "font_size": get_value('size').pt if get_value('size') else None,
        "font_color": rgb_to_hex(final_color) if final_color else None,
        "bold": get_value('bold')
    }

def expected_props(level):
    if level == 1:
        return {"font_family": "Times New Roman", "font_size": 16, "font_color": "000000", "bold": True}
    elif level == 2:
        return {"font_family": "Times New Roman", "font_size": 14, "font_color": "000000", "bold": True}
    elif level in [3, 4, 5]:
        return {"font_family": "Times New Roman", "font_size": 12, "font_color": "000000", "bold": True}
    return {}

def get_paragraph_page_map(doc_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(doc_path)

    para_to_page = {}
    for i, para in enumerate(doc.Paragraphs):
        para_to_page[i] = para.Range.Information(3)  # wdActiveEndPageNumber = 3

    doc.Close(False)
    word.Quit()
    return para_to_page

def run_tests_on_doc(doc_path):
    document = Document(doc_path)
    para_to_page_map = get_paragraph_page_map(doc_path)
    result_data = []
    # spell = SpellChecker()

    for idx, para in enumerate(document.paragraphs):
        para_text = para.text.strip()
        current_page = para_to_page_map.get(idx, "Unknown")

        # === Heading Font Validation ===
        if para.style.name.lower().startswith("heading"):
            try:
                level = int(para.style.name.split()[-1])
            except ValueError:
                continue

            run = para.runs[0] if para.runs else None
            expected = expected_props(level)

            if not run:
                result_data.append({
                    "Test Case Name": f"Heading Level {level} Font",
                    "Test Case Description": f"Validate font attributes for Heading Level {level}",
                    "Category": f"Heading Level {level}",
                    "Page Number": current_page,
                    "Text": para_text,
                    "Property": "run",
                    "Expected": "Font run exists",
                    "Actual": "Missing",
                    "Status": "FAIL"
                })
                continue

            actual = get_font_properties(run, para)

            for prop, expected_val in expected.items():
                actual_val = actual.get(prop)
                status = "PASS" if actual_val == expected_val else "FAIL"
                result_data.append({
                    "Test Case Name": f"Heading Level {level} Font",
                    "Test Case Description": f"Validate font attributes for Heading Level {level}",
                    "Category": f"Heading Level {level}",
                    "Page Number": current_page,
                    "Text": para_text,
                    "Property": prop,
                    "Expected": str(expected_val),
                    "Actual": str(actual_val) if actual_val is not None else "Missing",
                    "Status": status
                })

            # === Page number content validation for Heading 6 ===
            if level == 6 and ("page" in para_text.lower() or any(char.isdigit() for char in para_text)):
                result_data.append({
                    "Test Case Name": "Page Number Position",
                    "Test Case Description": "Validate page number in Heading Level 6",
                    "Category": "Page Number",
                    "Page Number": current_page,
                    "Text": para_text,
                    "Property": "Page number content",
                    "Expected": "Page number present",
                    "Actual": "Present" if para_text else "Missing",
                    "Status": "PASS" if para_text else "FAIL"
                })

        # === Spelling Mistake Check ===
        # words = [word.strip(".,:;!?()[]") for word in para_text.split()]
        # misspelled = spell.unknown(words)
        # for word in misspelled:
        #     result_data.append({
        #         "Test Case Name": "Spelling Check",
        #         "Test Case Description": "Check for spelling mistakes in document",
        #         "Category": "Spelling",
        #         "Page Number": current_page,
        #         "Text": para_text,
        #         "Property": "Spelling",
        #         "Expected": "Correct spelling",
        #         "Actual": word,
        #         "Status": "FAIL"
        #     })

    # Save result to Excel
    df = pd.DataFrame(result_data)
    excel_path = "test_report_detailed.xlsx"
    df.to_excel(excel_path, index=False)
    return excel_path
