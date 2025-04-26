# Crash Test Report Converter

A Python tool that extracts vehicle crash test data and injury metrics from a structured PDF and outputs it into a formatted Excel file with conditional color highlights.

## Features

- Extracts vehicle details (ID, type, powertrain, weight, etc.)
- Parses injury metrics like HIC, neck, pelvis, etc.
- Highlights:
  - Injury ≤ 80% → Yellow (Acceptable)
  - Injury > 100% → Red (Poor)
- Works with multi-page PDFs
- Simple file selection using GUI

## Usage

1. Run the script:
    ```bash
    python Convert.py
    ```
2. Choose a PDF file when prompted.
3. Excel file will be saved in the same location as the PDF.

## Dependencies

Install all required packages using:

```bash
pip install -r requirements.txt
# Crash-Report-Converter
