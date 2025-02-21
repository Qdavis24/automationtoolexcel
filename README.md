# Document Automation Tool

A Python-based automation tool that transforms Excel RFI (Request for Information) data into structured Word documents. The system intelligently maps questions from Excel spreadsheets to corresponding sections in Word templates.

## Features

- Automated extraction of RFI questions from Excel spreadsheets
- Smart section mapping based on document structure  
- Color-based row filtering for data cleanup
- Template-based Word document generation
- Configurable via YAML for easy maintenance

## Configuration

The system uses a YAML file for configuration:

```yaml
filepaths:
 data_spreadsheet: "./files/input.xlsx"
 document_template: "./files/template.docx"
 final_document: "./files/completed_review.docx"

excel:
 standard_column: "C"  # Column containing section identifiers
 rfi_column: "K"      # Column containing RFI questions 
 sheet_name: "FINAL verification"
 advanced:
   row_shift: 3       # Number of rows to skip from top
   header: 1          # Header row position
   ignore_color: "FF00B050"  # Rows with this color will be ignored
