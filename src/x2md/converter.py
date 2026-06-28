"""Core conversion pipeline: .docx → Markdown with image extraction
and HTML table post-processing."""

from pathlib import Path

from markitdown import MarkItDown

from x2md.images import extract_images, fix_references
from x2md.tables import extract_tables, replace_markdown_tables


def _docx_to_markdown_raw(docx_path: Path, output_path: Path) -> Path:
    """Convert a .docx file to Markdown using markitdown.

    Args:
        docx_path: Path to the input .docx file.
        output_path: Path to write the output .md file.

    Returns:
        The output Path.
    """
    md = MarkItDown()
    result = md.convert(str(docx_path))

    output_path.write_text(result.text_content, encoding="utf-8")
    return output_path


def convert(
    input_path: str | Path,
    output_dir: str | Path | None = None,
    extract_images_flag: bool = True,
    html_tables: bool = True,
) -> Path:
    """Convert a .docx file to Markdown with optional image and table handling.

    Args:
        input_path: Path to the input .docx file.
        output_dir: Output directory (default: same as input file).
        extract_images_flag: If True, extract embedded images and fix references.
        html_tables: If True, convert tables to HTML (colspan/rowspan support).

    Returns:
        Path to the generated .md file.

    Raises:
        FileNotFoundError: If the input file does not exist.
        ValueError: If the input file is not a .docx file.
    """
    input_path = Path(input_path).resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if input_path.suffix.lower() != ".docx":
        raise ValueError(
            f"Unsupported format: {input_path.suffix}. Expected .docx"
        )

    output_dir = Path(output_dir) if output_dir else input_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    output_name = input_path.stem + ".md"
    output_path = output_dir / output_name
    doc_stem = input_path.stem

    # Step 1: Convert with markitdown
    _docx_to_markdown_raw(input_path, output_path)

    # Step 2: Extract images and fix references
    if extract_images_flag:
        extract_images(input_path, output_dir, doc_stem)
        fix_references(output_path, input_path, doc_stem)

    # Step 3: Replace markdown tables with HTML tables
    if html_tables:
        html_table_list = extract_tables(input_path)
        if html_table_list:
            replace_markdown_tables(output_path, html_table_list)

    return output_path
