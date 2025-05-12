Feature: Word Document Format Testing

  Scenario: Validate styles and page numbers in the Word document
    Given a Word document from the input folder
    When the document is tested for heading styles and page number
    Then a test report is generated as an Excel file