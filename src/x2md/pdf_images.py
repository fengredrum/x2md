"""Image extraction and reference repair for PDF files.

Extracts embedded images from PDFs using pymupdf (fitz), renames them in
page order (fig_1.png, fig_2.png, ...), fixes truncated base64 references
in the markitdown-generated markdown, and inserts image references into
the markdown when markitdown produces no image output at all.
"""

import os
import re
from pathlib import Path
from urllib.parse import quote

# ── Regex patterns (from pdf2md.py) ────────────────────────────────────
TRUNCATED_BASE64_RE = re.compile(
    r"!\[([^\]]*)\]\(data:image/([^;]+);base64(\.\.\.?)\)"
)
IMAGE_RE = re.compile(r"!\[([^\]]*)\]\(data:image/[^;]+;base64,[A-Za-z0-9+/=]+\)")

# Pattern to find "图 N" figure references in Chinese documents
FIGURE_REF_RE = re.compile(r"图\s*(\d+)")


def _img_ref(doc_stem: str, fig_idx: int, ext: str) -> str:
    """Build a URL-encoded image reference path.

    URL-encodes the document stem so that markdown renderers (including
    VS Code's webview-based preview) can resolve paths containing Chinese
    characters or other special Unicode characters.

    Returns a path like: ``images/<encoded_stem>/fig_N.ext``
    """
    return f"images/{quote(doc_stem, safe='')}/fig_{fig_idx}.{ext}"


def extract_images_from_pdf(
    pdf_path: str | Path, output_dir: str | Path, doc_stem: str
) -> list[dict]:
    """Extract embedded images from a PDF using pymupdf.

    Images are saved to: <output_dir>/images/<doc_stem>/fig_N.ext
    Images smaller than 5KB are skipped (likely icons/decorations).

    Args:
        pdf_path: Path to the input PDF file.
        output_dir: Base output directory.
        doc_stem: Document name stem (used for subdirectory name).

    Returns:
        List of dicts with keys: fig_idx, ext, page_num, page_text.
        ``page_text`` is the first 200 chars of text from that page,
        used as an anchor to find the insertion point in the markdown.

    Raises:
        ImportError: If pymupdf is not installed.
    """
    try:
        import fitz
    except ImportError:
        raise ImportError(
            "pymupdf (fitz) is required for PDF image extraction. "
            "Install with: pip install x2md[pdf]"
        )

    pdf_path = Path(pdf_path)
    output_dir = Path(output_dir)
    target_dir = output_dir / "images" / doc_stem
    target_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(str(pdf_path))
    extracted = []
    fig_idx = 0

    for page_num in range(doc.page_count):
        page = doc[page_num]
        image_list = page.get_images(full=True)

        # Collect page text for anchoring
        page_text = page.get_text().strip()

        for img_info in image_list:
            xref = img_info[0]
            base_image = doc.extract_image(xref)
            if base_image:
                fig_idx += 1
                ext = base_image["ext"]
                img_bytes = base_image["image"]

                # Skip tiny images (likely icons/decorations)
                if len(img_bytes) < 5000:
                    continue

                img_name = f"fig_{fig_idx}.{ext}"
                img_path = target_dir / img_name
                img_path.write_bytes(img_bytes)

                extracted.append({
                    "fig_idx": fig_idx,
                    "ext": ext,
                    "page_num": page_num,
                    "page_text": page_text[:200],
                })

    doc.close()
    return extracted


def fix_pdf_image_references(
    md_path: str | Path, doc_stem: str, image_count: int
) -> None:
    """Replace truncated base64 image references with local file paths.

    If no images were extracted, truncated base64 references are stripped.
    Otherwise, they are replaced one-to-one with images/<doc_stem>/fig_N.ext
    paths in document order.

    Args:
        md_path: Path to the markdown file to fix.
        doc_stem: Document name stem for building image paths.
        image_count: Number of images extracted (0 = strip references).
    """
    md_path = Path(md_path)
    content = md_path.read_text(encoding="utf-8")

    if image_count == 0:
        # Strip truncated base64 references entirely
        content = TRUNCATED_BASE64_RE.sub("", content)
        md_path.write_text(content, encoding="utf-8")
        return

    truncated = TRUNCATED_BASE64_RE.findall(content)
    if not truncated:
        return

    def replacer(match, counter=[0]):
        alt = match.group(1)
        ext = match.group(2) if match.group(2) else "png"
        counter[0] += 1
        fig_name = f"fig_{counter[0]}.{ext}"
        img_ref = _img_ref(doc_stem, counter[0], ext)
        return f"![{alt}]({img_ref})"

    content = TRUNCATED_BASE64_RE.sub(replacer, content)
    md_path.write_text(content, encoding="utf-8")


def insert_image_references(
    md_path: str | Path,
    doc_stem: str,
    extracted_images: list[dict],
) -> int:
    """Insert image references into the markdown at appropriate positions.

    Used when markitdown produces no image references for a PDF's embedded
    images. For each extracted image, locates the page's text anchor in the
    markdown and inserts a ``![fig_N](images/...)`` reference.

    Strategy:
        1. For each extracted image, find the page's text snippet in the
           markdown to determine the insertion region.
        2. Look for a ``图 N`` reference near the insertion point.
        3. Insert the image reference after the ``图 N ...`` line.
        4. If no ``图 N`` is found, insert after the first paragraph on
           that page.

    Args:
        md_path: Path to the markdown file.
        doc_stem: Document name stem for building image paths.
        extracted_images: List of dicts from :func:`extract_images_from_pdf`.

    Returns:
        Number of image references inserted.
    """
    if not extracted_images:
        return 0

    md_path = Path(md_path)
    content = md_path.read_text(encoding="utf-8")

    inserted = 0
    # Process images in reverse order to preserve string positions
    # when inserting text into the middle of the content
    for img_info in reversed(extracted_images):
        fig_idx = img_info["fig_idx"]
        ext = img_info["ext"]
        page_text = img_info["page_text"]

        img_ref = _img_ref(doc_stem, fig_idx, ext)
        img_md = f"![图 {fig_idx}]({img_ref})"

        # Find insertion point: locate the page's text in the markdown
        insert_pos = _find_insertion_point(content, page_text, fig_idx)

        if insert_pos >= 0:
            # Insert image reference as its own paragraph before the
            # insertion point
            content = (
                content[:insert_pos]
                + f"\n{img_md}\n\n"
                + content[insert_pos:]
            )
            inserted += 1

    md_path.write_text(content, encoding="utf-8")
    return inserted


def _find_insertion_point(
    content: str, page_text: str, fig_idx: int
) -> int:
    """Find the best position to insert an image reference.

    Args:
        content: The full markdown content.
        page_text: Text from the PDF page containing the image.
        fig_idx: The figure index (for matching to ``图 N`` text).

    Returns:
        Character position for insertion, or -1 if no good spot found.
    """
    # Strategy 1: Find "图 N ..." text in the markdown
    # Look for the figure reference specifically
    tu_pattern = re.compile(rf"图\s*{fig_idx}\b")
    tu_match = tu_pattern.search(content)
    if tu_match:
        # Insert after the line containing "图 N"
        end_of_line = content.find("\n", tu_match.end())
        if end_of_line < 0:
            end_of_line = len(content)
        return end_of_line

    # Strategy 2: Find the page text anchor in the markdown
    if page_text:
        # Try to match the first substantial sentence (at least 30 chars)
        sentences = re.split(r"[。\n]", page_text)
        for sentence in sentences:
            sentence = sentence.strip()
            if len(sentence) >= 30:
                pos = content.find(sentence)
                if pos >= 0:
                    # Insert at the beginning of the paragraph containing
                    # this sentence
                    para_start = content.rfind("\n\n", 0, pos)
                    if para_start >= 0:
                        return para_start + 2
                    return pos

    return -1
