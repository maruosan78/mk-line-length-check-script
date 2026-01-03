# mk-line-length-check-script

A Python tool that checks per-line character limits in memoQ bilingual DOCX files and generates a standalone HTML report with overflow highlighting.

## What this script does
- Automatically detects a memoQ bilingual DOCX file in the current folder
- Prompts the user for a maximum character limit per line
- Replaces memoQ-style tags with real line breaks (`<...>`, `[...]`, `{...}`)
- Checks each logical line separately
- Flags segments where at least one line exceeds the configured limit
- Generates a standalone dark-mode HTML report
- Overflow characters are highlighted in yellow
- The Target column allows optional inline editing and reset for review

## Requirements
- Python 3.9 or newer
- python-docx library

Install the required dependency: `pip install python-docx`

## Usage
1. Export a bilingual DOCX file from memoQ.
2. Place the following files in the same folder:
   - `MK_Line_Length_Check_3.23.py`
   - your memoQ bilingual `.docx` file
3. Run the script: `python MK_Line_Length_Check_3.23.py`
4. Enter the maximum number of characters allowed per line when prompted.
5. Open the generated HTML report in your browser and review highlighted overflow characters.

## Supported input
- memoQ bilingual DOCX files only
- The DOCX file must contain a bilingual table with the following columns:
  - ID
  - Source
  - Target

## Output
- One standalone HTML report per run
- The report is generated even if no violations are found

## License
MIT License
