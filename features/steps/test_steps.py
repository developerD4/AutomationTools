from behave import given, when, then
from utils import run_tests_on_doc
import os

@given("a Word document from the input folder")
def step_given_doc(context):
    context.doc_path = "C:\\Assignment\\word_doc_tester\\input\\sample.docx"

@when("the document is tested for heading styles and page number")
def step_when_test_doc(context):
    context.results = run_tests_on_doc(context.doc_path)

@then("a test report is generated as an Excel file")
def step_then_generate_report(context):
    excel_file = context.results
    print(f"Test report saved as: {excel_file}")
    # If you want to open the Excel file automatically after the test run:
    os.startfile(excel_file)  # On Windows; you can adjust for other OS
