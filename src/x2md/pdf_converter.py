"""Core PDF-to-Markdown conversion pipeline.

Mirrors the DOCX ``converter.py`` pattern: validate input, detect PDF type,
branch on text vs scanned, delegate to sub-modules for tables/images/OCR,
apply shared postprocessing.
"""

from pathlib import Path

from markitdown import MarkItDown

from x2md.pdf_images import (
    extract_images_from_pdf,
    fix_pdf_image_references,
    insert_image_references,
)
from x2md.pdf_tables import extract_tables_from_pdf, replace_tables_in_markdown
from x2md.pdf_ocr import ocr_scanned_pdf
from x2md.postprocess import postprocess_markdown


def _detect_pdf_type(pdf_path: str) -> str:
    """Detect whether a PDF is text-based or scanned.

    Uses pymupdf to extract text from the first 3 pages.
    Returns ``"text"`` if total character count ≥ 100, otherwise ``"scanned"``.

    Raises:
        ImportError: If pymupdf is not installed.
    """
    try:
        import fitz
    except ImportError:
        raise ImportError(
            "pymupdf (fitz) is required for PDF type detection. "
            "Install with: pip install x2md[pdf]"
        )

    doc = fitz.open(pdf_path)
    check_pages = min(3, doc.page_count)
    total_chars = 0
    for i in range(check_pages):
        text = doc[i].get_text()
        total_chars += len(text.strip())
    doc.close()

    if total_chars < 100:
        return "scanned"
    return "text"


def _convert_text_pdf(input_path: str, output_path: str) -> str:
    """Convert a text-based PDF to raw markdown using markitdown.

    Args:
        input_path: Path to the input PDF file.
        output_path: Path to write the output .md file.

    Returns:
        The output path string.
    """
    md = MarkItDown()
    result = md.convert(input_path)
    Path(output_path).write_text(result.text_content, encoding="utf-8")
    return output_path


def convert_pdf(
    input_path: str | Path,
    output_dir: str | Path | None = None,
    extract_images: bool = True,
    fix_tables: bool = True,
    force_ocr: bool = False,
    dpi: int = 300,
) -> Path:
    """Convert a PDF file to Markdown with optional image extraction,
    table repair, and OCR for scanned documents.

    Pipeline:
        1. Validate input (exists, .pdf extension).
        2. Detect PDF type (text vs scanned) using pymupdf.
        3. Branch:
           - **Text PDF:** markitdown → tables → images → postprocess
           - **Scanned PDF:** tesseract OCR → postprocess
        4. Return path to the generated .md file.

    Args:
        input_path: Path to the input .pdf file.
        output_dir: Output directory (default: same directory as input).
        extract_images: Extract embedded images and fix references.
        fix_tables: Use pdfplumber to extract and repair tables as HTML.
        force_ocr: Force OCR mode even if the PDF has a text layer.
        dpi: OCR rendering DPI (default 300).

    Returns:
        Path to the output .md file.

    Raises:
        FileNotFoundError: If the input file does not exist.
        ValueError: If the input file is not a .pdf file.
        ImportError: If required PDF dependencies are not installed,
                     with a message suggesting ``pip install x2md[pdf]``.
    """
    input_path = Path(input_path).resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if input_path.suffix.lower() != ".pdf":
        raise ValueError(
            f"Unsupported format: {input_path.suffix}. Expected .pdf"
        )

    output_dir = Path(output_dir) if output_dir else input_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    output_name = input_path.stem + ".md"
    output_path = output_dir / output_name
    doc_stem = input_path.stem

    # ── Detect PDF type ──
    if force_ocr:
        pdf_type = "scanned"
    else:
        try:
            pdf_type = _detect_pdf_type(str(input_path))
        except Exception:
            pdf_type = "text"

    # ── Convert ──
    if pdf_type == "scanned":
        # OCR branch
        ocr_scanned_pdf(
            str(input_path), str(output_path), str(output_dir), doc_stem, dpi
        )
    else:
        # Text branch
        # Step 1: markitdown base conversion
        _convert_text_pdf(str(input_path), str(output_path))

        # Step 2: Table repair
        if fix_tables:
            tables = extract_tables_from_pdf(str(input_path))
            if tables:
                replace_tables_in_markdown(str(output_path), tables)

        # Step 3: Image extraction
        if extract_images:
            extracted = extract_images_from_pdf(
                str(input_path), str(output_dir), doc_stem
            )
            image_count = len(extracted)
            fix_pdf_image_references(str(output_path), doc_stem, image_count)
            # Insert image references when markitdown produced none
            insert_image_references(str(output_path), doc_stem, extracted)

    # Step 4: Shared postprocessing (always applied for PDF)
    postprocess_markdown(output_path)

    return output_path
