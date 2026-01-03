# mk-line-length-check-script

A Python tool that checks per-line character limits in memoQ bilingual DOCX files and generates a standalone HTML report with overflow highlighting.

The script is designed to be usable by non-technical users.
It automatically checks required dependencies at startup and installs them if they are missing.

## What this script does
- Automatically detects a memoQ bilingual DOCX file in the current folder
- Automatically checks required dependencies on startup
- Installs missing dependencies if needed (no manual setup required)
- Prompts the user for a maximum character limit per line
- Replaces memoQ-style tags with real line breaks (`<...>`, `[...]`, `{...}`)
- Checks each logical line separately
- Flags segments where at least one line exceeds the configured limit
- Generates a standalone dark-mode HTML report
- Overflow characters are highlighted in yellow
- The Target column allows optional inline editing for visual review

## Requirements
- Python 3.9 or newer
- Windows
- Internet connection (required only on first run, for automatic dependency installation)

No manual installation of Python libraries is required.
The script will automatically install the needed dependency if it is missing.

## Usage
1. Export a bilingual DOCX file from memoQ.
2. Place the following files in the same folder:
   - `MK_Line_Length_Check_3.24.py`
   - your memoQ bilingual `.docx` file
3. Open Command Prompt or PowerShell in that folder.
4. Run the script:
   - `python MK_Line_Length_Check_3.24.py`
5. On first run, the script will check required dependencies and install them automatically if needed.
6. When prompted, enter the maximum number of characters allowed per line.
7. Open the generated HTML report in your browser and review highlighted overflow characters.

## Supported input
- memoQ bilingual DOCX files only
- The DOCX file must contain a bilingual table with the following columns:
  - ID
  - Source
  - Target

## Output
- One standalone HTML report per run
- Characters exceeding the configured line limit are highlighted in yellow
- The report supports inline editing for visual review purposes

## Notes
- Inline editing in the HTML report is for review only.
- Editing the report does **not** modify the original DOCX file.

## License
MIT License
