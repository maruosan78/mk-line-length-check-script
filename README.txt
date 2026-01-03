MK Line Length Check Script

This tool checks per-line character limits in memoQ bilingual DOCX files
and generates an HTML report with highlighted overflows.

REQUIREMENTS
- Windows
- Python 3.9 or newer
- python-docx library

If needed, install the dependency:
pip install python-docx

USAGE
1. Export a bilingual DOCX file from memoQ.
2. Put the following files in the same folder:
   - MK_Line_Length_Check_3.23.py
   - your memoQ bilingual .docx file
3. Open Command Prompt or PowerShell in that folder.
4. Run:
   python MK_Line_Length_Check_3.23.py
5. Enter the maximum number of characters per line when asked.
6. Open the generated HTML report in your browser.

OUTPUT
- An HTML report is created in the same folder.
- Characters exceeding the limit are highlighted in yellow.

LICENSE
MIT License
