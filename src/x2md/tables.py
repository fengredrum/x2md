"""DOCX table to HTML conversion — extracted from docx2md.py.

Parses w:tbl XML elements from .docx files and converts them to HTML <table>
strings with proper colspan/rowspan handling.
"""

import re
import zipfile
from pathlib import Path


def _extract_cell_paragraphs(cell_xml: str) -> list:
    """Extract all paragraph HTML text from a w:tc XML element.

    Each w:p becomes a paragraph (separated by <br>),
    w:r with w:b and w:i are converted to <b>/<i> tags.
    Returns a list of paragraph text strings.
    """
    paragraphs = []
    para_matches = re.findall(r"<w:p[ >].*?</w:p>", cell_xml, re.DOTALL)
    if not para_matches:
        para_matches = re.findall(r"<w:p\s*/>", cell_xml, re.DOTALL)

    for p_xml in para_matches:
        runs = re.findall(r"<w:r[ >].*?</w:r>", p_xml, re.DOTALL)
        if not runs:
            paragraphs.append("")
            continue

        run_texts = []
        for r_xml in runs:
            is_bold = bool(
                re.search(r"<w:b\s*/>", r_xml) or re.search(r"<w:b[ >]", r_xml)
            )
            is_italic = bool(
                re.search(r"<w:i\s*/>", r_xml) or re.search(r"<w:i[ >]", r_xml)
            )

            texts = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", r_xml)
            text = "".join(texts)

            if not text:
                continue

            if is_bold:
                text = f"<b>{text}</b>"
            if is_italic:
                text = f"<i>{text}</i>"
            run_texts.append(text)

        paragraphs.append("".join(run_texts))

    return paragraphs


def _extract_cell_text(cell_xml: str) -> str:
    """Extract cell text from w:tc XML, with multi-paragraph cells
    joined by <br>."""
    paragraphs = _extract_cell_paragraphs(cell_xml)
    return "<br>".join(paragraphs) if paragraphs else ""


def _merge_adjacent_tags(text: str) -> str:
    """Merge adjacent same-name HTML tags to reduce fragmentation.

    Example: <b>a</b><b>b</b> → <b>ab</b>
    """
    for tag in ("b", "i"):
        pattern = re.compile(rf"</{tag}><{tag}>")
        while pattern.search(text):
            text = pattern.sub("", text)
    return text


def parse_table_to_html(table_xml: str) -> str:
    """Convert a single w:tbl XML element to an HTML <table> string.

    Correctly handles:
    - w:gridSpan → colspan
    - w:vMerge (restart/continue) → rowspan
    - Bold/italic formatting → <b>/<i>
    """
    rows_xml = re.findall(r"<w:tr[ >].*?</w:tr>", table_xml, re.DOTALL)
    if not rows_xml:
        return ""

    # ── First pass: analyze vMerge, compute rowspan ──
    open_merges: dict = {}
    rowspans: dict = {}

    for row_idx, row_xml in enumerate(rows_xml):
        cells_xml = re.findall(r"<w:tc[ >].*?</w:tc>", row_xml, re.DOTALL)
        col_idx = 0

        for cell_xml in cells_xml:
            gs = re.search(r'<w:gridSpan[^>]*w:val="(\d+)"', cell_xml)
            colspan = int(gs.group(1)) if gs else 1

            vm_restart = bool(re.search(r'<w:vMerge[^>]*w:val="restart"', cell_xml))
            has_vmerge = "<w:vMerge" in cell_xml
            vm_continue = has_vmerge and not vm_restart

            if vm_restart:
                open_merges[col_idx] = {"start_row": row_idx, "count": 1}
            elif vm_continue:
                if col_idx in open_merges:
                    open_merges[col_idx]["count"] += 1
            else:
                if col_idx in open_merges:
                    info = open_merges.pop(col_idx)
                    if info["count"] > 1:
                        rowspans[(info["start_row"], col_idx)] = info["count"]

            col_idx += colspan

    # Close any remaining open merges
    for col_idx, info in open_merges.items():
        if info["count"] > 1:
            rowspans[(info["start_row"], col_idx)] = info["count"]

    # ── Second pass: generate HTML ──
    # Track continuation cells to skip
    skip_cells: set = set()
    for (start_row, col), rs in rowspans.items():
        for offset in range(1, rs):
            skip_cells.add((start_row + offset, col))

    html_parts = ["<table>"]

    for row_idx, row_xml in enumerate(rows_xml):
        html_parts.append("  <tr>")
        cells_xml = re.findall(r"<w:tc[ >].*?</w:tc>", row_xml, re.DOTALL)
        col_idx = 0

        for cell_xml in cells_xml:
            gs = re.search(r'<w:gridSpan[^>]*w:val="(\d+)"', cell_xml)
            colspan = int(gs.group(1)) if gs else 1

            has_vmerge = "<w:vMerge" in cell_xml
            vm_restart = bool(re.search(r'<w:vMerge[^>]*w:val="restart"', cell_xml))
            vm_continue = has_vmerge and not vm_restart

            if vm_continue:
                col_idx += colspan
                continue

            attrs = []
            if colspan > 1:
                attrs.append(f'colspan="{colspan}"')

            rs = rowspans.get((row_idx, col_idx))
            if rs and rs > 1:
                attrs.append(f'rowspan="{rs}"')

            # Determine if header cell: first row or all-bold content
            tag = "td"
            if row_idx == 0:
                tag = "th"
            else:
                paragraphs = _extract_cell_paragraphs(cell_xml)
                cell_text = "".join(paragraphs)
                if (
                    cell_text.startswith("<b>")
                    and cell_text.endswith("</b>")
                    and cell_text.count("<b>") == 1
                ):
                    tag = "th"

            text = _extract_cell_text(cell_xml)

            attr_str = (" " + " ".join(attrs)) if attrs else ""
            html_parts.append(f"    <{tag}{attr_str}>{text}</{tag}>")

            col_idx += colspan

        html_parts.append("  </tr>")

    html_parts.append("</table>")
    result = "\n".join(html_parts)
    return _merge_adjacent_tags(result)


def extract_tables(docx_path: str | Path) -> list:
    """Extract all tables from a .docx file's document.xml.

    Args:
        docx_path: Path to the .docx file.

    Returns:
        List of HTML table strings in document order.
    """
    docx_path = Path(docx_path)
    with zipfile.ZipFile(docx_path) as z:
        if "word/document.xml" not in z.namelist():
            return []
        xml_content = z.read("word/document.xml").decode("utf-8")

    table_matches = re.findall(r"<w:tbl[ >].*?</w:tbl>", xml_content, re.DOTALL)
    html_tables = []
    for tbl_xml in table_matches:
        html = parse_table_to_html(tbl_xml)
        if html:
            html_tables.append(html)

    return html_tables


def replace_markdown_tables(md_path: str | Path, html_tables: list) -> None:
    """Replace Markdown table blocks in a file with HTML tables.

    Reads the file at md_path, replaces markdown tables one-to-one with
    HTML tables from the list, and writes the result back.
    """
    md_path = Path(md_path)
    content = md_path.read_text(encoding="utf-8")
    new_content = _replace_markdown_tables_in_content(content, html_tables)
    md_path.write_text(new_content, encoding="utf-8")


def _replace_markdown_tables_in_content(md_content: str, html_tables: list) -> str:
    """Replace Markdown table blocks with HTML tables in a string.

    Matches markdown tables one-to-one with HTML tables. If there are fewer
    HTML tables than markdown tables, remaining markdown tables stay as-is.
    """
    if not html_tables:
        return md_content

    md_table_re = re.compile(
        r"(?:^\|.*\|\s*\n)"  # header row
        r"^\|[-\s:|]+\|\s*\n"  # separator row
        r"(?:^\|.*\|\s*(?:\n|$))*",  # data rows
        re.MULTILINE,
    )

    tables_found = md_table_re.findall(md_content)
    if not tables_found:
        return md_content

    result = md_content
    for i, md_table in enumerate(tables_found):
        if i < len(html_tables):
            replacement = "\n\n" + html_tables[i] + "\n\n"
            result = result.replace(md_table, replacement, 1)

    return result
