import os
import re
import html
import sys
import subprocess
from typing import List, Dict, Tuple, Any


VERSION = "3.23"


def ensure_python_docx() -> Any:
    """
    Ensure python-docx is installed and importable.
    If missing, attempt to install it via pip for the current Python interpreter.
    Returns the imported Document callable/class from python-docx.
    """
    try:
        from docx import Document  # type: ignore
        print("Dependency check OK: python-docx is available.")
        return Document
    except Exception:
        print("Dependency missing: python-docx is not installed.")
        print("Attempting to install python-docx automatically...")

    # Try installing with the same interpreter that runs this script
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--user", "python-docx"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        # Second attempt without silencing output (more informative for users)
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "--user", "python-docx"]
            )
        except Exception as e:
            print("Error: automatic installation of python-docx failed.")
            print("Please install it manually and run the script again:")
            print("  pip install python-docx")
            raise e

    # Verify import after install
    try:
        from docx import Document  # type: ignore
        print("Dependency installed OK: python-docx is now available.")
        return Document
    except Exception as e:
        print("Error: python-docx still cannot be imported after installation.")
        print("Please install it manually and run the script again:")
        print("  pip install python-docx")
        raise e


def normalize_with_linebreaks(text: str) -> str:
    """
    Replace memoQ-style tags with real line breaks so we can count characters
    per logical line.

    Tags covered:
    - <...>
    - [...]
    - {...}
    """
    if not text:
        return ""

    tag_pattern = r"(?:<[^>]+>|\[[^\]]+?\]|\{[^}]+\})"
    text = re.sub(tag_pattern, "\n", text)

    # Normalize line endings
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return text


def find_bilingual_table(doc: Any) -> Tuple[Any, Any, Any]:
    """
    Find the bilingual table and the header row that contains 'ID'.
    Return (table, header_row_index, id_col_index).

    We look for:
    - any table
    - any row in that table whose cells contain 'ID' and at least 3 columns.
    """
    for table in doc.tables:
        for row_index, row in enumerate(table.rows):
            header_cells = [c.text.strip() for c in row.cells]
            if "ID" in header_cells and len(header_cells) >= 3:
                id_idx = header_cells.index("ID")
                return table, row_index, id_idx
    return None, None, None


def analyze_document(Document: Any, docx_path: str, char_limit: int) -> List[Dict]:
    """
    Analyze the DOCX file and return a list of violations.

    Each violation is a dict with:
    - id                (string, full ID from memoQ, but report shows only leading segment number)
    - source            (source text)
    - target            (original target text)
    - max_len           (maximum line length in the segment)
    - segment_len       (total length of the segment - sum of non-empty lines)
    - line_lengths      (list[int] - per non-empty line, kept for future use)
    - highlighted_target_html (target with overflow chars wrapped in <span>)
    """
    print(f"\nLoading DOCX document: {docx_path}")
    doc = Document(docx_path)

    table, header_row_index, id_idx = find_bilingual_table(doc)
    if table is None:
        print("No table with 'ID' and at least three columns was found.")
        return []

    print(
        f"Found bilingual table. Header row index: {header_row_index}, "
        f"'ID' column index: {id_idx}"
    )

    # Source column = first column to the right of ID
    src_idx = id_idx + 1
    # Target column = first column to the right of source
    trg_idx = src_idx + 1

    header_cells = [c.text.strip() for c in table.rows[header_row_index].cells]
    if trg_idx >= len(header_cells):
        print("Could not determine target column (not enough columns).")
        return []

    violations: List[Dict] = []

    # Iterate over segment rows (rows after the header)
    for row in table.rows[header_row_index + 1:]:
        seg_id = row.cells[id_idx].text.strip()
        src_text = row.cells[src_idx].text
        trg_text_original = row.cells[trg_idx].text

        trg_processed = normalize_with_linebreaks(trg_text_original)
        lines = trg_processed.split("\n")

        non_empty_lines = [l for l in lines if l.strip() != ""]
        if not non_empty_lines:
            continue

        lengths = [len(l) for l in non_empty_lines]
        max_len = max(lengths)
        segment_len = sum(lengths)

        if max_len <= char_limit:
            continue

        # Build HTML with highlighted overflow characters
        highlighted_lines: List[str] = []
        for line in lines:
            if line == "":
                highlighted_lines.append("<br>")
                continue

            length = len(line)
            if length <= char_limit:
                highlighted_lines.append(html.escape(line) + "<br>")
            else:
                prefix = line[:char_limit]
                overflow = line[char_limit:]
                highlighted_line = (
                    f"{html.escape(prefix)}"
                    f"<span class=\"overflow\">{html.escape(overflow)}</span><br>"
                )
                highlighted_lines.append(highlighted_line)

        highlighted_target_html = "".join(highlighted_lines)

        violations.append(
            {
                "id": seg_id,
                "source": src_text,
                "target": trg_text_original,
                "max_len": max_len,
                "segment_len": segment_len,
                "line_lengths": lengths,
                "highlighted_target_html": highlighted_target_html,
            }
        )

    print(
        f"Found {len(violations)} segments with at least one line longer than "
        f"{char_limit} characters."
    )
    return violations


def build_html_report(
    violations: List[Dict],
    output_path: str,
    char_limit: int,
    source_filename: str,
) -> None:
    """
    Build a standalone dark-mode HTML report with highlighted overflow characters
    and small edit/reset icons in the Target column.
    """
    escaped_filename = html.escape(os.path.basename(source_filename))
    total_segments = len(violations)

    if total_segments == 0:
        summary_text = (
            f"NO SEGMENTS EXCEED THE LIMIT OF {char_limit} CHARACTERS PER LINE."
        )
    else:
        summary_text = (
            f"Found {total_segments} segments with at least one line longer "
            f"than {char_limit} characters."
        )

    rows_html: List[str] = []
    for v in violations:
        full_id = v["id"]
        id_short = full_id.split()[0] if full_id else ""
        id_html = html.escape(id_short)

        src_html = html.escape(v["source"])
        target_html = v["highlighted_target_html"]

        data_original_html = html.escape(target_html, quote=True)

        target_cell_html = f"""
        <div class="target-container" data-original-html="{data_original_html}">
            <div class="target-text" contenteditable="false">
                {target_html}
            </div>
            <div class="target-actions">
                <button class="icon-btn edit-btn" title="Toggle edit">✏️</button>
                <button class="icon-btn reset-btn" title="Reset to original">⟲</button>
            </div>
        </div>
        """

        line_length_html = str(v["max_len"])
        segment_length_html = str(v.get("segment_len", sum(v["line_lengths"])))

        row = f"""
        <tr>
            <td class="cell-id">{id_html}</td>
            <td class="cell-source">{src_html}</td>
            <td class="cell-target">{target_cell_html}</td>
            <td class="cell-maxlen">{line_length_html}</td>
            <td class="cell-linelens">{segment_length_html}</td>
        </tr>
        """
        rows_html.append(row)

    rows_block = "\n".join(rows_html)

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>MK Line Length Check - limit {char_limit}</title>
<style>
    body {{
        margin: 0;
        padding: 0;
        font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        background-color: #050608;
        color: #f5f5f5;
    }}
    header {{
        padding: 20px 32px;
        background: linear-gradient(135deg, #10131a, #151a24);
        border-bottom: 1px solid #262b36;
    }}
    h1 {{
        margin: 0 0 4px 0;
        font-size: 22px;
        letter-spacing: 0.03em;
    }}
    .version-pill {{
        display: inline-block;
        margin-left: 8px;
        padding: 2px 8px;
        border-radius: 999px;
        background: #1f2937;
        color: #cbd5f5;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }}
    .subheader {{
        font-size: 13px;
        color: #9aa0b5;
    }}
    main {{
        padding: 16px 32px 32px 32px;
    }}
    .summary {{
        margin-bottom: 16px;
        font-size: 14px;
        font-weight: 600;
        color: #e5e7eb;
    }}
    .table-wrapper {{
        border-radius: 16px;
        border: 1px solid #1f2933;
        background: radial-gradient(circle at top left, #111827 0, #020617 45%);
        box-shadow: 0 14px 40px rgba(0,0,0,0.8);
        overflow: hidden;
    }}
    table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
    }}
    thead {{
        background: linear-gradient(90deg, #111827, #020617);
    }}
    th, td {{
        padding: 8px 10px;
        border-bottom: 1px solid #111827;
        vertical-align: top;
    }}
    th {{
        text-align: left;
        font-weight: 600;
        color: #9ca3af;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        white-space: nowrap;
    }}
    tbody tr:nth-child(even) {{
        background-color: rgba(15,23,42,0.7);
    }}
    tbody tr:nth-child(odd) {{
        background-color: rgba(3,7,18,0.9);
    }}
    tbody tr:hover {{
        background-color: rgba(55,65,81,0.8);
    }}
    .cell-id {{
        width: 60px;
        font-weight: 600;
        color: #e5e7eb;
        white-space: nowrap;
    }}
    .cell-source {{
        width: 30%;
        color: #d1d5db;
    }}
    .cell-target {{
        width: 40%;
        color: #f9fafb;
    }}
    .cell-maxlen {{
        width: 90px;
        text-align: right;
        font-variant-numeric: tabular-nums;
        color: #f97316;
        font-weight: 600;
    }}
    .cell-linelens {{
        width: 110px;
        text-align: right;
        font-variant-numeric: tabular-nums;
        color: #9ca3af;
    }}
    .overflow {{
        background-color: #ffff00;
        color: #000;
        font-weight: 600;
        padding: 0 1px;
    }}
    .footer-note {{
        margin-top: 18px;
        font-size: 11px;
        color: #6b7280;
    }}
    .target-container {{
        position: relative;
    }}
    .target-text[contenteditable="true"] {{
        outline: 1px dashed #6ec1ff;
        outline-offset: 2px;
        background-color: rgba(15, 23, 42, 0.6);
    }}
    .target-actions {{
        margin-top: 6px;
        display: flex;
        gap: 4px;
        font-size: 11px;
    }}
    .icon-btn {{
        border: 1px solid #374151;
        background-color: #111827;
        color: #e5e7eb;
        border-radius: 999px;
        padding: 2px 6px;
        cursor: pointer;
        font-size: 11px;
        line-height: 1;
    }}
    .icon-btn:hover {{
        background-color: #1f2937;
    }}
</style>
<script>
document.addEventListener('click', function(event) {{
    const editBtn = event.target.closest('.edit-btn');
    const resetBtn = event.target.closest('.reset-btn');

    if (editBtn) {{
        const container = editBtn.closest('.target-container');
        if (!container) return;
        const textDiv = container.querySelector('.target-text');
        if (!textDiv) return;

        const isEditable = textDiv.getAttribute('contenteditable') === 'true';
        if (isEditable) {{
            textDiv.setAttribute('contenteditable', 'false');
        }} else {{
            textDiv.setAttribute('contenteditable', 'true');
            textDiv.focus();
        }}
    }}

    if (resetBtn) {{
        const container = resetBtn.closest('.target-container');
        if (!container) return;
        const textDiv = container.querySelector('.target-text');
        if (!textDiv) return;

        const originalHtml = container.getAttribute('data-original-html');
        if (originalHtml != null) {{
            textDiv.innerHTML = originalHtml;
        }}
        textDiv.setAttribute('contenteditable', 'false');
    }}
}});
</script>
</head>
<body>
<header>
    <div>
        <h1>MK Line Length Check<span class="version-pill">v{VERSION}</span></h1>
        <div class="subheader">
            Source file: <strong>{escaped_filename}</strong> ·
            Limit per line: <strong>{char_limit} characters</strong>
        </div>
    </div>
</header>
<main>
    <div class="summary">
        <strong>{summary_text}</strong>
    </div>

    <div class="table-wrapper">
        <table id="resultsTable">
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Source</th>
                    <th>Target (overflow highlighted)</th>
                    <th>Line length</th>
                    <th>Segment length</th>
                </tr>
            </thead>
            <tbody>
                {rows_block}
            </tbody>
        </table>
    </div>

    <div class="footer-note">
        Overflow characters beyond the configured limit are highlighted in yellow.
    </div>
</main>
</body>
</html>
"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"HTML report written to: {output_path}")


def main():
    print(f"=== memoQ Line Length Checker - v{VERSION} (DOCX + HTML) ===")

    # Ensure dependency before doing anything else
    Document = ensure_python_docx()

    current_dir = os.getcwd()
    files = os.listdir(current_dir)

    docx_files = [f for f in files if f.lower().endswith(".docx")]
    rtf_like_files = [f for f in files if f.lower().endswith(".rtf") or f.lower().endswith(".rtx")]

    if docx_files:
        chosen_file = docx_files[0]
    elif rtf_like_files:
        chosen_file = rtf_like_files[0]
    else:
        print("Error: no DOCX or RTF/RTX file found in the current folder.")
        return

    input_path = os.path.join(current_dir, chosen_file)
    print(f"Detected file in current folder: {input_path}")

    if not os.path.isfile(input_path):
        print("Error: detected file does not exist. Please check the folder.")
        return

    ext = os.path.splitext(input_path)[1].lower()
    if ext != ".docx":
        print("Error: this version works with DOCX only. Please export a bilingual DOCX from memoQ.")
        return

    try:
        limit_str = input("Enter character limit per line: ").strip()
        char_limit = int(limit_str)
    except ValueError:
        print("Error: character limit must be an integer.")
        return

    violations = analyze_document(Document, input_path, char_limit)

    base, _ = os.path.splitext(input_path)
    output_path = f"{base}_length_report_{char_limit}.html"

    build_html_report(violations, output_path, char_limit, input_path)


if __name__ == "__main__":
    main()
