"""Image extraction and reference repair for .docx files.

Extracts embedded images from a .docx ZIP archive, renames them in document
order (fig_1.png, fig_2.png, ...), and fixes truncated base64 references
in the generated markdown.
"""

import os
import re
import zipfile
from pathlib import Path
from urllib.parse import quote
from xml.etree import ElementTree as ET


# ── Regex patterns (from docx2md.py) ──────────────────────────────────
TRUNCATED_BASE64_RE = re.compile(
    r"!\[([^\]]*)\]\(data:image/([^;]+);base64(\.\.\.?)\)"
)

IMAGE_RE = re.compile(r"!\[([^\]]*)\]\(data:image/[^;]+;base64,[A-Za-z0-9+/=]+\)")


def get_image_order_in_document(docx_path: str | Path) -> list:
    """Parse document.xml to determine the order images appear in the document.

    Returns:
        List of (alt_text, embed_rId, position) tuples in document order.
    """
    docx_path = Path(docx_path)
    image_order = []
    with zipfile.ZipFile(docx_path) as z:
        content = z.read("word/document.xml").decode("utf-8")

    for m in re.finditer(
        r"<wp:(?:inline|anchor).*?</wp:(?:inline|anchor)>", content, re.DOTALL
    ):
        snippet = m.group()
        name_m = re.search(r'name="([^"]*)"', snippet)
        embed_m = re.search(r'r:embed="([^"]*)"', snippet)
        name = name_m.group(1) if name_m else "图片"
        rid = embed_m.group(1) if embed_m else None
        image_order.append((name, rid, m.start()))

    return image_order


def extract_images(
    docx_path: str | Path, output_dir: str | Path, doc_stem: str
) -> int:
    """Extract images from a .docx file, naming them in document order.

    Images are saved to: <output_dir>/images/<doc_stem>/fig_1.png, fig_2.png, ...

    Args:
        docx_path: Path to the .docx file.
        output_dir: Base output directory.
        doc_stem: Document name stem (used for subdirectory name).

    Returns:
        Number of images found in the document.
    """
    docx_path = Path(docx_path)
    output_dir = Path(output_dir)
    target_dir = output_dir / "images" / doc_stem
    target_dir.mkdir(parents=True, exist_ok=True)

    # 1. Get image order from document
    image_order = get_image_order_in_document(docx_path)

    # 2. Read relationships to map rId → original filename
    rid_to_filename = {}
    with zipfile.ZipFile(docx_path) as z:
        rels_path = "word/_rels/document.xml.rels"
        if rels_path in z.namelist():
            root = ET.fromstring(z.read(rels_path))
            for rel in root:
                rid = rel.get("Id")
                target = rel.get("Target", "")
                if "image" in target.lower():
                    rid_to_filename[rid] = os.path.basename(target)

    # 3. Extract and rename images in document order
    extracted = []
    with zipfile.ZipFile(docx_path) as z:
        for idx, (name, rid, _) in enumerate(image_order, 1):
            if rid and rid in rid_to_filename:
                original_filename = rid_to_filename[rid]
                ext = os.path.splitext(original_filename)[1] or ".png"
                new_name = f"fig_{idx}{ext}"
                out_path = target_dir / new_name

                source_path = f"word/media/{original_filename}"
                if source_path in z.namelist():
                    with z.open(source_path) as src, open(out_path, "wb") as dst:
                        dst.write(src.read())
                    extracted.append((new_name, out_path.stat().st_size))

    return len(image_order)


def fix_references(
    md_path: str | Path, docx_path: str | Path, doc_stem: str
) -> None:
    """Replace truncated base64 image references in markdown with local file paths.

    Image reference format: images/<doc_stem>/fig_N.ext

    Strategy:
        1. If markdown has full base64 images → skip (already embedded)
        2. If truncated base64 (ending with ...) → replace with file reference
        3. Match markdown images to document images in order
    """
    md_path = Path(md_path)
    docx_path = Path(docx_path)
    content = md_path.read_text(encoding="utf-8")

    truncated = TRUNCATED_BASE64_RE.findall(content)
    full_base64 = IMAGE_RE.findall(content)

    if not truncated and not full_base64:
        return

    if full_base64:
        return

    # Get image order from document
    doc_images = get_image_order_in_document(docx_path)

    # Map rId → original filename for extension
    rid_to_filename = {}
    with zipfile.ZipFile(docx_path) as z:
        rels_path = "word/_rels/document.xml.rels"
        if rels_path in z.namelist():
            root = ET.fromstring(z.read(rels_path))
            for rel in root:
                rid = rel.get("Id")
                target = rel.get("Target", "")
                if "image" in target.lower():
                    rid_to_filename[rid] = os.path.basename(target)

    # Build image reference paths in document order
    image_refs = []
    for idx, (name, rid, _) in enumerate(doc_images, 1):
        if rid and rid in rid_to_filename:
            ext = os.path.splitext(rid_to_filename[rid])[1] or ".png"
            img_ref = f"images/{quote(doc_stem, safe='')}/fig_{idx}{ext}"
            image_refs.append((name, img_ref))

    # Replace truncated base64 references with local file references
    def replacer(match, counter=[0]):
        alt = match.group(1)
        if counter[0] < len(image_refs):
            doc_name, img_ref = image_refs[counter[0]]
            img_alt = alt if alt and not alt.startswith("图片") else doc_name
            counter[0] += 1
            return f"![{img_alt}]({img_ref})"
        else:
            return match.group(0)

    content = TRUNCATED_BASE64_RE.sub(replacer, content)
    md_path.write_text(content, encoding="utf-8")
