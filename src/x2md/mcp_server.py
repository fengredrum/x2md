"""MCP server exposing .docx and .pdf to Markdown conversion tools.

Usage:
    uv run x2md-mcp
"""

from pathlib import Path

from mcp.server.fastmcp import FastMCP

from x2md.converter import convert

mcp = FastMCP(
    "x2md",
    instructions="Convert .docx, .pdf, and .xlsx/.xls files to Markdown with image extraction and HTML tables",
)


@mcp.tool()
def convert_docx_to_markdown(
    input_path: str,
    output_dir: str | None = None,
    extract_images: bool = True,
    html_tables: bool = True,
) -> str:
    """Convert a .docx file to Markdown with optional image extraction
    and HTML table conversion.

    Images are extracted and saved to: <output_dir>/images/<doc_name>/fig_N.ext
    Tables are converted to HTML format preserving colspan and rowspan.

    Args:
        input_path: Absolute path to the input .docx file.
        output_dir: Directory for the output .md file and images subdirectory.
                    Defaults to the same directory as the input file.
        extract_images: Extract embedded images and fix references.
        html_tables: Convert tables to HTML (with colspan/rowspan).
                     When False, uses basic markdown tables.

    Returns:
        A summary of the conversion result.
    """
    result_path = convert(
        input_path=input_path,
        output_dir=output_dir,
        extract_images_flag=extract_images,
        html_tables=html_tables,
    )

    # Build summary
    file_size = result_path.stat().st_size

    images_dir = result_path.parent / "images" / Path(input_path).stem
    img_count = (
        len([f for f in images_dir.iterdir() if f.is_file()])
        if images_dir.exists()
        else 0
    )

    return (
        f"Conversion complete.\n"
        f"  Input:  {input_path}\n"
        f"  Output: {result_path}\n"
        f"  Size:   {file_size:,} bytes\n"
        f"  Images: {img_count} extracted"
    )


@mcp.tool()
def convert_pdf_to_markdown(
    input_path: str,
    output_dir: str | None = None,
    extract_images: bool = True,
    fix_tables: bool = True,
    force_ocr: bool = False,
    dpi: int = 300,
) -> str:
    """Convert a .pdf file to Markdown with optional image extraction,
    table repair, and OCR for scanned documents.

    Automatically detects whether the PDF is text-based or scanned.
    Text-based PDFs use markitdown + pdfplumber for table repair.
    Scanned PDFs use tesseract OCR with Chinese language support.

    Images are extracted and saved to: <output_dir>/images/<doc_name>/fig_N.ext
    Tables are converted to HTML format.

    Args:
        input_path: Absolute path to the input .pdf file.
        output_dir: Directory for the output .md file and images subdirectory.
                    Defaults to the same directory as the input file.
        extract_images: Extract embedded images and fix references.
        fix_tables: Use pdfplumber to extract and repair tables as HTML.
        force_ocr: Force OCR even if the PDF has a text layer.
        dpi: OCR rendering DPI (default: 300).

    Returns:
        A summary of the conversion result.
    """
    from x2md.pdf_converter import convert_pdf

    try:
        result_path = convert_pdf(
            input_path=input_path,
            output_dir=output_dir,
            extract_images=extract_images,
            fix_tables=fix_tables,
            force_ocr=force_ocr,
            dpi=dpi,
        )
    except ImportError as e:
        return (
            f"PDF conversion dependencies are not installed.\n"
            f"  {e}\n"
            f"  Install with: pip install x2md[pdf]\n"
            f"  Or: uv sync --extra pdf"
        )

    # Build summary
    file_size = result_path.stat().st_size

    images_dir = result_path.parent / "images" / Path(input_path).stem
    img_count = (
        len([f for f in images_dir.iterdir() if f.is_file()])
        if images_dir.exists()
        else 0
    )

    return (
        f"Conversion complete.\n"
        f"  Input:  {input_path}\n"
        f"  Output: {result_path}\n"
        f"  Size:   {file_size:,} bytes\n"
        f"  Images: {img_count} extracted"
    )


@mcp.tool()
def convert_excel_to_markdown(
    input_path: str,
    output_dir: str | None = None,
) -> str:
    """Convert a .xlsx/.xls file to Markdown with HTML tables.

    Each worksheet becomes a ## section. Tables are converted to HTML
    format preserving merged cells (colspan/rowspan) and cell styles
    (bold, italic, font size, colors, background, alignment, borders).

    .xlsx files use openpyxl (full style support).
    .xls files use xlrd (requires pip install x2md[xls]).

    Args:
        input_path: Absolute path to the input .xlsx or .xls file.
        output_dir: Directory for the output .md file.
                    Defaults to the same directory as the input file.

    Returns:
        A summary of the conversion result.
    """
    from x2md.excel_converter import convert_excel

    try:
        result_path = convert_excel(
            input_path=input_path,
            output_dir=output_dir,
        )
    except ImportError as e:
        return (
            f"Excel conversion dependency missing.\n"
            f"  {e}\n"
            f"  Install with: pip install x2md[xls]"
        )

    file_size = result_path.stat().st_size
    return (
        f"Conversion complete.\n"
        f"  Input:  {input_path}\n"
        f"  Output: {result_path}\n"
        f"  Size:   {file_size:,} bytes"
    )


def main():
    """Entry point — run the MCP server over stdio transport."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
