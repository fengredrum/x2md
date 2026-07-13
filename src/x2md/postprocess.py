"""Shared Markdown postprocessing utilities.

Operations that clean up markdown output — useful for both DOCX and PDF
conversion pipelines (currently applied automatically in the PDF pipeline).
"""

import re
from pathlib import Path


def postprocess_markdown(md_path: str | Path) -> None:
    """Clean up markdown formatting issues in place.

    Operations:
        1. Compress 4+ consecutive blank lines to 2
        2. Remove extra whitespace between CJK characters (OCR artifact)
        3. Ensure blank lines before all heading lines
    """
    md_path = Path(md_path)
    content = md_path.read_text(encoding="utf-8")

    # Compress consecutive blank lines (4+ → 2)
    content = re.sub(r"\n{4,}", "\n\n\n", content)

    # Remove extra whitespace between CJK characters (OCR residue)
    content = re.sub(r"([一-鿿])\s+([一-鿿])", r"\1\2", content)

    # Ensure blank lines before headings
    content = re.sub(r"([^\n])\n(#{1,6}\s)", r"\1\n\n\2", content)

    md_path.write_text(content, encoding="utf-8")
