#!/usr/bin/env python3
"""pdf2md.py — 将 PDF 文件转换为 Markdown，正确处理表格和图片，扫描版使用 OCR。

特性：
    - 自动检测 PDF 类型（文本型 / 扫描版）
    - 文本型：markitdown 基础转换 + pdfplumber 精确表格修复
    - 扫描版：tesseract OCR 识别 + 格式排版还原
    - 自动提取内嵌图片，修复图片引用
    - 表格以 HTML 格式输出，正确保留结构

依赖：
    markitdown[all]     文档转 Markdown（通过 uv 管理）
    pdfplumber          PDF 表格提取
    pymupdf (fitz)      PDF 渲染与图片提取
    pytesseract         Tesseract OCR Python 封装
    tesseract-ocr       系统 OCR 引擎（需单独安装）
    tesseract-ocr-chi-sim  中文语言包

用法：
    uv run python scripts/pdf2md.py <input.pdf> [options]

选项：
    --output-dir DIR    输出目录（默认：PDF 所在目录）
    --no-images         不提取图片
    --no-table-fix      不修复表格（使用 markitdown 原生输出）
    --force-ocr         强制使用 OCR（即使 PDF 含文本层）
    --dpi N             OCR 渲染 DPI（默认：300）

示例：
    uv run python scripts/pdf2md.py "人陪方案/参考/广州航海学院_无人驾驶航空器系统工程_2025版.pdf"
    uv run python scripts/pdf2md.py "人陪方案/参考/华南理工大学_低空技术与工程（盖章扫描版）.pdf"
"""

import argparse
import os
import re
import sys
from pathlib import Path


# ── 常量 ──────────────────────────────────────────────────────────────
BASE64_DATA_URI_RE = re.compile(r'!\[([^\]]*)\]\(data:image/[^;]+;base64,\.\.\.?\)')
IMAGE_RE = re.compile(r'!\[([^\]]*)\]\(data:image/[^;]+;base64,[A-Za-z0-9+/=]+\)')
TRUNCATED_BASE64_RE = re.compile(
    r'!\[([^\]]*)\]\(data:image/([^;]+);base64(\.\.\.?)\)'
)
# markitdown 可能产生 reference: 风格的图片引用
REF_IMAGE_RE = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')


# ── 环境设置 ──────────────────────────────────────────────────────────
def setup_tesseract_env():
    """配置 tesseract 环境变量（使用本地安装的 tesseract）。"""
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


# ── 依赖检查 ──────────────────────────────────────────────────────────
def ensure_dependencies():
    """确保所有必要的 Python 包已安装。"""
    deps = {
        "markitdown": "markitdown[all]",
        "pdfplumber": "pdfplumber",
        "fitz": "pymupdf",
        "pytesseract": "pytesseract",
        "PIL": "Pillow",
    }
    missing = []
    for import_name, pip_name in deps.items():
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pip_name)

    if missing:
        print(f"[setup] 安装缺失的依赖: {' '.join(missing)}")
        os.system(f"{sys.executable} -m pip install -q {' '.join(missing)}")


# ── PDF 类型检测 ──────────────────────────────────────────────────────
def detect_pdf_type(pdf_path: str) -> str:
    """检测 PDF 是文本型还是扫描版。

    使用 pymupdf 提取前三页文本，若总字符数 < 100 则判定为扫描版。
    返回 "text" 或 "scanned"。
    """
    import fitz

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


# ── 文本型 PDF 处理 ───────────────────────────────────────────────────
def convert_text_pdf(input_path: str, output_path: str) -> str:
    """使用 markitdown 转换文本型 PDF → Markdown。"""
    from markitdown import MarkItDown

    md = MarkItDown()
    result = md.convert(input_path)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(result.text_content)

    print(f"[markitdown] 文本转换完成 → {output_path}")
    return output_path


# ── 表格提取与修复 ────────────────────────────────────────────────────
def extract_tables_from_pdf(pdf_path: str) -> list:
    """使用 pdfplumber 从 PDF 中提取所有表格，返回 HTML 表格列表。

    返回:
        [(page_num, html_table), ...] — 按页序排列
    """
    import pdfplumber

    all_tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue
                html = _table_to_html(table)
                if html:
                    all_tables.append((page_num, html))
    return all_tables


def _table_to_html(table: list) -> str:
    """将 pdfplumber 提取的二维表格列表转为 HTML <table> 字符串。"""
    if not table or not table[0]:
        return ""

    # 清理：移除全 None 的行和列
    cleaned = []
    for row in table:
        if row and any(cell is not None and str(cell).strip() for cell in row):
            cleaned.append([str(cell).strip() if cell is not None else "" for cell in row])

    if not cleaned:
        return ""

    # 判断第一行是否是表头
    has_header = len(cleaned) >= 2

    lines = ["<table>"]
    for row_idx, row in enumerate(cleaned):
        lines.append("  <tr>")
        tag = "th" if (has_header and row_idx == 0) else "td"
        for cell in row:
            # 清理单元格文本
            cell_text = cell.replace("\n", "<br>").replace("|", "&#124;")
            lines.append(f"    <{tag}>{cell_text}</{tag}>")
        lines.append("  </tr>")
    lines.append("</table>")
    return "\n".join(lines)


def replace_tables_in_markdown(md_path: str, pdf_tables: list) -> int:
    """将 markdown 中的破表替换为 pdfplumber 提取的 HTML 表格。

    替换策略：按出现顺序，将 markdown 中检测到的表格块替换为 HTML 表格。
    返回替换数量。
    """
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    if not pdf_tables:
        print("[tables] 未检测到表格")
        return 0

    # 匹配 markdown 表格块（使用 | 分隔的表格）
    md_table_re = re.compile(
        r'(?:^\|.*\|\s*\n)'          # header row
        r'^\|[\-\s:|]+\|\s*\n'       # separator row
        r'(?:^\|.*\|\s*(?:\n|$))*',  # data rows
        re.MULTILINE
    )

    md_tables = md_table_re.findall(content)
    if not md_tables:
        # 尝试更宽松的匹配：连续的带有 | 的行（含格式错乱的）
        loose_re = re.compile(
            r'(?:^.*\|.*\n){2,}', re.MULTILINE
        )
        md_tables = loose_re.findall(content)

    replaced = 0
    for i, md_table in enumerate(md_tables):
        if i < len(pdf_tables):
            _, html_table = pdf_tables[i]
            replacement = f"\n\n{html_table}\n\n"
            content = content.replace(md_table, replacement, 1)
            replaced += 1

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"[tables] 已替换 {replaced}/{len(pdf_tables)} 个表格为 HTML 格式")
    return replaced


# ── 图片提取 ──────────────────────────────────────────────────────────
def extract_images_from_pdf(pdf_path: str, output_dir: str, doc_stem: str) -> int:
    """使用 pymupdf 提取 PDF 中的嵌入图片。

    图片保存到 output_dir/images/<doc_stem>/fig_N.ext。
    返回提取的图片数量。
    """
    import fitz

    target_dir = os.path.join(output_dir, "images", doc_stem)
    os.makedirs(target_dir, exist_ok=True)

    doc = fitz.open(pdf_path)
    extracted = []
    fig_idx = 0

    for page_num in range(doc.page_count):
        page = doc[page_num]
        image_list = page.get_images(full=True)

        for img_info in image_list:
            xref = img_info[0]
            base_image = doc.extract_image(xref)
            if base_image:
                fig_idx += 1
                ext = base_image["ext"]
                img_bytes = base_image["image"]

                # 跳过太小（可能是图标/装饰）的图片（< 5KB）
                if len(img_bytes) < 5000:
                    continue

                img_name = f"fig_{fig_idx}.{ext}"
                img_path = os.path.join(target_dir, img_name)
                with open(img_path, "wb") as f:
                    f.write(img_bytes)
                extracted.append((img_name, len(img_bytes)))

    doc.close()

    if extracted:
        print(f"[images] 提取 {len(extracted)} 张图片 → {target_dir}/")
        for name, size in extracted[:10]:
            print(f"         {name} ({size:,} bytes)")
        if len(extracted) > 10:
            print(f"         ... 共 {len(extracted)} 张")

    return len(extracted)


def fix_image_references(md_path: str, doc_stem: str, image_count: int):
    """修复 markdown 中的图片引用。

    将截断的 base64 data URI 替换为本地文件路径引用。
    """
    if image_count == 0:
        print("[images] 无嵌入图片，跳过引用修复")
        # 清理可能残留的截断 base64 引用
        with open(md_path, "r", encoding="utf-8") as f:
            content = f.read()
        content = TRUNCATED_BASE64_RE.sub("", content)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(content)
        return

    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    truncated = TRUNCATED_BASE64_RE.findall(content)
    if not truncated:
        print("[images] 无截断图片引用需修复")
        return

    print(f"[images] 修复 {min(len(truncated), image_count)} 处图片引用")

    def replacer(match, counter=[0]):
        alt = match.group(1)
        ext = match.group(2) if match.group(2) else "png"
        counter[0] += 1
        fig_name = f"fig_{counter[0]}.{ext}"
        img_ref = f"images/{doc_stem}/{fig_name}"
        return f"![{alt}]({img_ref})"

    content = TRUNCATED_BASE64_RE.sub(replacer, content)

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(content)


# ── 扫描版 PDF 处理（OCR）─────────────────────────────────────────────
def ocr_scanned_pdf(pdf_path: str, output_path: str, output_dir: str, doc_stem: str, dpi: int = 300):
    """使用 tesseract OCR 处理扫描版 PDF。

    流程：
    1. 逐页渲染为图片
    2. tesseract 中文 OCR
    3. 格式排版还原
    4. 图片引用嵌入
    """
    import fitz
    from PIL import Image
    import pytesseract
    import io

    pytesseract.pytesseract.tesseract_cmd = os.path.expanduser("~/.local/bin/tesseract")

    # 图片输出目录
    img_dir = os.path.join(output_dir, "images", doc_stem)
    os.makedirs(img_dir, exist_ok=True)

    doc = fitz.open(pdf_path)
    total_pages = doc.page_count
    print(f"[ocr] 共 {total_pages} 页，逐页 OCR 处理中（DPI={dpi}）...")

    md_parts = []
    md_parts.append(f"# {doc_stem}\n")
    md_parts.append(f"> ⚠️ 本文档由扫描版 PDF 经 OCR 识别生成，可能存在识别误差。每页附原始图片供对照。\n")
    md_parts.append(f"> 总页数：{total_pages}\n")

    for page_num in range(total_pages):
        page = doc[page_num]
        print(f"  [ocr] 处理第 {page_num + 1}/{total_pages} 页...", end=" ")

        # 渲染页面为图片
        mat = fitz.Matrix(dpi / 72, dpi / 72)  # 300 DPI
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")

        # 保存页面图片
        page_img_name = f"page_{page_num + 1:03d}.png"
        page_img_path = os.path.join(img_dir, page_img_name)
        with open(page_img_path, "wb") as f:
            f.write(img_bytes)

        # OCR 识别
        img = Image.open(io.BytesIO(img_bytes))
        try:
            ocr_text = pytesseract.image_to_string(img, lang="chi_sim", config="--psm 6")
        except Exception as e:
            ocr_text = f"[OCR 错误: {e}]"

        # 格式排版
        formatted = format_ocr_text(ocr_text)

        # 组装：每页前面放图片引用，后面放 OCR 文本
        md_parts.append(f"\n## 第 {page_num + 1} 页\n")
        md_parts.append(f"![第 {page_num + 1} 页](images/{doc_stem}/{page_img_name})\n")
        md_parts.append(formatted)
        md_parts.append("")

        print(f"OK ({len(ocr_text)} 字符)")

    doc.close()

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_parts))

    print(f"[ocr] OCR 转换完成 → {output_path}")
    return output_path


def format_ocr_text(raw_text: str) -> str:
    """对 OCR 原始文本进行格式排版。

    规则：
    - 去除行首行尾多余空白
    - 合并过短的行（可能的换行错误）
    - 检测可能的标题行（短行/关键词开头）
    - 保留段落间距
    """
    lines = raw_text.split("\n")
    formatted = []
    prev_empty = False

    # 常见的标题关键词（中文公文）
    title_keywords = [
        "一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、", "十、",
        "（一）", "（二）", "（三）", "（四）", "（五）",
        "1.", "2.", "3.", "4.", "5.",
        "培养目标", "毕业要求", "核心课程", "课程体系", "专业特色",
        "专业概况", "师资", "就业", "培养模式", "实践",
        "主干学科", "学制", "授予学位", "学分", "第",
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

        # 检测标题行
        is_title = False
        for kw in title_keywords:
            if stripped.startswith(kw) and len(stripped) < 40:
                is_title = True
                break

        if is_title:
            if buffer:
                formatted.append(buffer)
                buffer = ""
            formatted.append(f"\n### {stripped}")
        else:
            # 判断是否与前一行合并
            if buffer and not buffer.endswith(("。", "，", "；", "：", "）", ")", "》")):
                buffer += stripped
            else:
                if buffer:
                    formatted.append(buffer)
                buffer = stripped

    if buffer:
        formatted.append(buffer)

    return "\n".join(formatted)


# ── 后处理 ────────────────────────────────────────────────────────────
def postprocess_markdown(md_path: str):
    """清理 markdown 文件的格式问题。"""
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    # 压缩连续空行（超过2个连续空行 → 2个空行）
    content = re.sub(r'\n{4,}', '\n\n\n', content)

    # 修复中文文本中的多余空格（OCR 残留）
    content = re.sub(r'([一-鿿])\s+([一-鿿])', r'\1\2', content)

    # 确保标题前后有空行
    content = re.sub(r'([^\n])\n(#{1,6}\s)', r'\1\n\n\2', content)

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"[post] 后处理完成")


# ── 主流程 ────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="将 PDF 转换为 Markdown，正确处理表格和图片，扫描版使用 OCR",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="输入的 PDF 文件路径")
    parser.add_argument("--output-dir", "-o", default=None,
                        help="输出目录（默认：PDF 所在目录）")
    parser.add_argument("--no-images", action="store_true",
                        help="不提取图片")
    parser.add_argument("--no-table-fix", action="store_true",
                        help="不修复表格（使用 markitdown 原生输出）")
    parser.add_argument("--force-ocr", action="store_true",
                        help="强制使用 OCR（即使 PDF 含文本层）")
    parser.add_argument("--dpi", type=int, default=300,
                        help="OCR 渲染 DPI（默认：300）")

    args = parser.parse_args()

    # 验证输入
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"[error] 文件不存在: {args.input}")
        sys.exit(1)
    if input_path.suffix.lower() != '.pdf':
        print(f"[error] 仅支持 .pdf 文件，收到: {input_path.suffix}")
        sys.exit(1)

    # 设置 tesseract 环境
    setup_tesseract_env()

    # 确保依赖
    ensure_dependencies()

    # 确定输出路径
    if args.output_dir:
        output_dir = str(Path(args.output_dir))
    else:
        output_dir = str(input_path.parent)
    output_name = input_path.stem + ".md"
    output_path = os.path.join(output_dir, output_name)
    doc_stem = input_path.stem

    # ── 开始转换 ──
    print(f"\n{'='*60}")
    print(f"[convert] 输入: {input_path}")
    print(f"[convert] 输出: {output_path}")
    print(f"{'='*60}\n")

    # 检测 PDF 类型
    if args.force_ocr:
        pdf_type = "scanned"
        print("[detect] 强制使用 OCR 模式")
    else:
        try:
            pdf_type = detect_pdf_type(str(input_path))
            print(f"[detect] PDF 类型: {'扫描版 (OCR)' if pdf_type == 'scanned' else '文本型'}")
        except Exception as e:
            print(f"[detect] 类型检测失败: {e}，按文本型处理")
            pdf_type = "text"

    if pdf_type == "scanned":
        # OCR 分支
        ocr_scanned_pdf(str(input_path), output_path, output_dir, doc_stem, args.dpi)
    else:
        # 文本型分支
        # Step 1: markitdown 基础转换
        convert_text_pdf(str(input_path), str(output_path))

        # Step 2: 表格修复
        if not args.no_table_fix:
            try:
                tables = extract_tables_from_pdf(str(input_path))
                if tables:
                    replace_tables_in_markdown(str(output_path), tables)
                else:
                    print("[tables] 未检测到表格")
            except Exception as e:
                print(f"[tables] 表格处理失败: {e}")

        # Step 3: 图片提取
        image_count = 0
        if not args.no_images:
            try:
                image_count = extract_images_from_pdf(str(input_path), output_dir, doc_stem)
                fix_image_references(str(output_path), doc_stem, image_count)
            except Exception as e:
                print(f"[images] 图片处理失败: {e}")

    # Step 4: 后处理
    try:
        postprocess_markdown(output_path)
    except Exception as e:
        print(f"[post] 后处理失败: {e}")

    # ── 完成 ──
    print(f"\n{'='*60}")
    print(f"[done] 转换完成！")
    print(f"  Markdown: {output_path}")
    if not args.no_images:
        img_dir = os.path.join(output_dir, "images", doc_stem)
        print(f"  图片:     {img_dir}/")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
