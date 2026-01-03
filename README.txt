MK Line Length Check Script

This tool checks per-line character limits in memoQ bilingual DOCX files
and generates a standalone HTML report with highlighted overflows.

The script is designed for non-technical users.
It automatically checks required dependencies at startup and installs
them if they are missing.

REQUIREMENTS
- Windows
- Python 3.9 or newer
- Internet connection (required only on first run, for dependency install)

No manual installation of Python libraries is required.
The script will automatically install the needed dependency if missing.

USAGE
1. Export a bilingual DOCX file from memoQ.
2. Put the following files in the same folder:
   - MK_Line_Length_Check_3.24.py
   - your memoQ bilingual .docx file
3. Open Command Prompt or PowerShell in that folder.
4. Run:
   python MK_Line_Length_Check_3.24.py
5. On first run, the script will check required dependencies.
   If something is missing, it will install it automatically.
6. When prompted, enter the maximum number of characters per line.
7. Open the generated HTML report in your browser.

OUTPUT
- One standalone HTML report is created in the same folder.
- Characters exceeding the configured line limit are highlighted in yellow.
- The report supports inline editing for visual review.

NOTES
- Inline editing in the HTML report is for review purposes only.
- Editing the report does NOT modify the original DOCX file.

LICENSE
MIT License
