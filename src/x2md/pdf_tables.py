"""PDF table extraction and markdown replacement.

Uses pdfplumber to extract tables from PDFs, converts them to HTML, and
replaces broken markdown-format tables in the markitdown-generated output.
"""

import re
from pathlib import Path


def _table_to_html(table: list) -> str:
    """Convert a pdfplumber 2D table list to an HTML <table> string.

    Args:
        table: 2D list of cell values from pdfplumber.

    Returns:
        HTML table string, or empty string if the table is empty.
    """
    if not table or not table[0]:
        return ""

    # Clean: remove all-None rows and columns
    cleaned = []
    for row in table:
        if row and any(cell is not None and str(cell).strip() for cell in row):
            cleaned.append(
                [str(cell).strip() if cell is not None else "" for cell in row]
            )

    if not cleaned:
        return ""

    has_header = len(cleaned) >= 2

    lines = ["<table>"]
    for row_idx, row in enumerate(cleaned):
        lines.append("  <tr>")
        tag = "th" if (has_header and row_idx == 0) else "td"
        for cell in row:
            cell_text = cell.replace("\n", "<br>").replace("|", "&#124;")
            lines.append(f"    <{tag}>{cell_text}</{tag}>")
        lines.append("  </tr>")
    lines.append("</table>")
    return "\n".join(lines)


def extract_tables_from_pdf(pdf_path: str | Path) -> list[tuple[int, str]]:
    """Extract all tables from a PDF using pdfplumber.

    Args:
        pdf_path: Path to the input PDF file.

    Returns:
        List of (page_number, html_table_string) tuples in page order.

    Raises:
        ImportError: If pdfplumber is not installed.
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError(
            "pdfplumber is required for PDF table extraction. "
            "Install with: pip install x2md[pdf]"
        )

    pdf_path = Path(pdf_path)
    all_tables = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue
                html = _table_to_html(table)
                if html:
                    all_tables.append((page_num, html))

    return all_tables


def replace_tables_in_markdown(
    md_path: str | Path, pdf_tables: list[tuple[int, str]]
) -> int:
    """Replace broken markdown tables with pdfplumber-extracted HTML tables.

    Strategy:
        1. First try strict markdown table regex (header row + separator + data rows).
        2. Fall back to loose regex (any consecutive lines containing ``|``).
        3. One-to-one replacement in document order.

    Args:
        md_path: Path to the markdown file to process.
        pdf_tables: List of (page_num, html_table) tuples from pdfplumber.

    Returns:
        Number of tables replaced.
    """
    md_path = Path(md_path)
    content = md_path.read_text(encoding="utf-8")

    if not pdf_tables:
        return 0

    # Strict regex: well-formed markdown tables
    md_table_re = re.compile(
        r"(?:^\|.*\|\s*\n)"  # header row
        r"^\|[\-\s:|]+\|\s*\n"  # separator row
        r"(?:^\|.*\|\s*(?:\n|$))*",  # data rows
        re.MULTILINE,
    )

    md_tables = md_table_re.findall(content)
    if not md_tables:
        # Loose regex fallback: any consecutive lines containing |
        loose_re = re.compile(r"(?:^.*\|.*\n){2,}", re.MULTILINE)
        md_tables = loose_re.findall(content)

    replaced = 0
    for i, md_table in enumerate(md_tables):
        if i < len(pdf_tables):
            _, html_table = pdf_tables[i]
            replacement = f"\n\n{html_table}\n\n"
            content = content.replace(md_table, replacement, 1)
            replaced += 1

    md_path.write_text(content, encoding="utf-8")
    return replaced
