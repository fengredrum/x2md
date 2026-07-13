"""OCR pipeline for scanned PDFs.

Renders each page to an image, runs tesseract OCR with Chinese language
support, and produces markdown with embedded page images and formatted text.
"""

import os
from pathlib import Path
from urllib.parse import quote


def _setup_tesseract_env() -> None:
    """Configure tesseract environment variables for local installs.

    Looks for tesseract under ~/.local/bin, ~/.local/lib, and
    ~/.local/share/tessdata, adding them to PATH, LD_LIBRARY_PATH,
    and TESSDATA_PREFIX respectively.
    """
    local_bin = os.path.expanduser("~/.local/bin")
    local_lib = os.path.expanduser("~/.local/lib")
    local_tessdata = os.path.expanduser("~/.local/share/tessdata")

    if os.path.exists(os.path.join(local_bin, "tesseract")):
        os.environ["PATH"] = local_bin + os.pathsep + os.environ.get("PATH", "")
    if os.path.exists(local_lib):
        current_ld = os.environ.get("LD_LIBRARY_PATH", "")
        if local_lib not in current_ld:
            os.environ["LD_LIBRARY_PATH"] = local_lib + (
                os.pathsep + current_ld if current_ld else ""
            )
    if os.path.exists(local_tessdata):
        os.environ["TESSDATA_PREFIX"] = local_tessdata


def format_ocr_text(raw_text: str) -> str:
    """Format raw OCR output for better readability.

    Rules:
        - Strip extra whitespace per line.
        - Merge short lines that don't end with Chinese sentence-ending
          punctuation (continuing lines).
        - Detect heading-like lines (short, starting with keywords) and
          add ``### `` prefix.
        - Preserve paragraph spacing.

    Args:
        raw_text: Raw OCR output from tesseract.

    Returns:
        Formatted markdown string.
    """
    lines = raw_text.split("\n")
    formatted = []
    prev_empty = False

    title_keywords = [
        "一、", "二、", "三、", "四、", "五、",
        "六、", "七、", "八、", "九、", "十、",
        "（一）", "（二）", "（三）",
        "（四）", "（五）",
        "1.", "2.", "3.", "4.", "5.",
        "培养目标", "毕业要求",
        "核心课程", "课程体系",
        "专业特色", "专业概况",
        "师资", "就业", "培养模式",
        "实践", "主干学科", "学制",
        "授予学位", "学分", "第",
    ]

    buffer = ""
    for line in lines:
        stripped = line.strip()
        if not stripped:
            if buffer:
                formatted.append(buffer)
                buffer = ""
            if not prev_empty:
                formatted.append("")
                prev_empty = True
            continue
        prev_empty = False

        # Detect heading-like lines
        is_title = any(
            stripped.startswith(kw) and len(stripped) < 40
            for kw in title_keywords
        )

        if is_title:
            if buffer:
                formatted.append(buffer)
                buffer = ""
            formatted.append(f"\n### {stripped}")
        else:
            # Merge with previous line if buffer doesn't end with
            # sentence-ending punctuation
            if buffer and not buffer.endswith(
                ("。", "，", "；", "：", "）", ")", "》")
            ):
                buffer += stripped
            else:
                if buffer:
                    formatted.append(buffer)
                buffer = stripped

    if buffer:
        formatted.append(buffer)

    return "\n".join(formatted)


def ocr_scanned_pdf(
    pdf_path: str | Path,
    output_path: str | Path,
    output_dir: str | Path,
    doc_stem: str,
    dpi: int = 300,
) -> str:
    """Process a scanned PDF page-by-page with tesseract OCR.

    For each page:
        1. Render to PNG at the specified DPI.
        2. Run tesseract with ``chi_sim`` language and ``--psm 6``.
        3. Format OCR output with :func:`format_ocr_text`.
        4. Embed page image reference and OCR text in markdown.

    Args:
        pdf_path: Path to the input PDF file.
        output_path: Path to write the output .md file.
        output_dir: Base output directory (for images subdirectory).
        doc_stem: Document name stem for image paths.
        dpi: Rendering DPI for page images (default 300).

    Returns:
        The output path as a string.

    Raises:
        ImportError: If required dependencies are not installed.
    """
    try:
        import fitz
        import pytesseract
        from PIL import Image
        import io
    except ImportError as e:
        raise ImportError(
            f"OCR dependencies are required for scanned PDF processing: {e}. "
            "Install with: pip install x2md[pdf]"
        )

    _setup_tesseract_env()
    pytesseract.pytesseract.tesseract_cmd = os.path.expanduser(
        "~/.local/bin/tesseract"
    )

    output_path = Path(output_path)
    output_dir = Path(output_dir)

    # Image output directory
    img_dir = output_dir / "images" / doc_stem
    img_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(str(pdf_path))
    total_pages = doc.page_count

    md_parts = []
    md_parts.append(f"# {doc_stem}\n")
    md_parts.append(
        f"> ⚠️ "
        f"本文档由扫描版 PDF 经 OCR "
        f"识别生成，可能存在识别"
        f"误差。每页附原始图片供"
        f"对照。\n"
    )
    md_parts.append(f"> 总页数：{total_pages}\n")

    for page_num in range(total_pages):
        page = doc[page_num]

        # Render page to image
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")

        # Save page image
        page_img_name = f"page_{page_num + 1:03d}.png"
        page_img_path = img_dir / page_img_name
        page_img_path.write_bytes(img_bytes)

        # OCR
        img = Image.open(io.BytesIO(img_bytes))
        try:
            ocr_text = pytesseract.image_to_string(
                img, lang="chi_sim", config="--psm 6"
            )
        except Exception as e:
            ocr_text = f"[OCR 错误: {e}]"

        # Format and assemble
        formatted = format_ocr_text(ocr_text)
        md_parts.append(f"\n## 第 {page_num + 1} 页\n")
        md_parts.append(
            f"![第 {page_num + 1} 页](images/{quote(doc_stem, safe='')}/{page_img_name})\n"
        )
        md_parts.append(formatted)
        md_parts.append("")

    doc.close()

    output_path.write_text("\n".join(md_parts), encoding="utf-8")
    return str(output_path)
