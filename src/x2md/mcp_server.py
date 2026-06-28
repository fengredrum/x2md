"""MCP server exposing .docx to Markdown conversion tools.

Usage:
    uv run x2md-mcp
"""

from pathlib import Path

from mcp.server.fastmcp import FastMCP

from x2md.converter import convert

mcp = FastMCP(
    "x2md",
    instructions="Convert .docx files to Markdown with image extraction and HTML tables",
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


def main():
    """Entry point — run the MCP server over stdio transport."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
