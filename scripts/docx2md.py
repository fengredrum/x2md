#!/usr/bin/env python3
"""docx2md.py — 将 .docx 文件转换为 Markdown，并正确提取嵌入图片。

特性：
    - 使用 markitdown 进行基础转换
    - 表格以 HTML 格式输出，正确保留单元格合并信息（colspan/rowspan）
    - 自动提取并修复图片引用
    - 图片按文档出现顺序命名：images/<文档名>/fig_1.png, fig_2.png, ...

依赖：
    markitdown[all]     文档转 Markdown（通过 uv 管理）

用法：
    uv run python scripts/docx2md.py <input.docx> [options]

选项：
    --output-dir DIR    输出目录（默认：.docx 所在目录）
    --llm-model MODEL   使用 OpenAI 兼容 LLM 分析内容（需设置环境变量）
    --llm-api-base URL  API 端点（默认: https://api.deepseek.com）
    --llm-api-key KEY   API 密钥（默认读 DEEPSEEK_API_KEY 环境变量）
    --no-images         不提取图片
    --no-html-tables    禁用 HTML 表格转换（使用 markitdown 原生的 Markdown 表格）

环境变量：
    DEEPSEEK_API_KEY    DeepSeek API 密钥（使用 --llm-model 时必需）

示例：
    uv run python scripts/docx2md.py "参考材料/低空飞行器工程技术专业（本科）人才培养方案_0603zh.docx"
    uv run python scripts/docx2md.py "参考材料/xxx.docx" --output-dir ./output
    uv run python scripts/docx2md.py "参考材料/xxx.docx" --llm-model deepseek-v4-flash
"""

import argparse
import os
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


# ── 常量 ──────────────────────────────────────────────────────────────
BASE64_DATA_URI_RE = re.compile(r'!\[([^\]]*)\]\(data:image/[^;]+;base64,\.\.\.?\)')
IMAGE_RE = re.compile(r'!\[([^\]]*)\]\(data:image/[^;]+;base64,[A-Za-z0-9+/=]+\)')

# markitdown 输出的截断 base64 模式（base64 后直接跟 ...，无逗号）
TRUNCATED_BASE64_RE = re.compile(
    r'!\[([^\]]*)\]\(data:image/([^;]+);base64(\.\.\.?)\)'
)


# ── DOCX 表格 → HTML 转换 ──────────────────────────────────────────────

def _extract_cell_paragraphs(cell_xml: str) -> list:
    """从 w:tc XML 中提取所有段落的 HTML 文本。

    每个 w:p 成为一个段落（用 <br> 分隔），
    w:r 中的 w:b 和 w:i 转换为 <b>/<i> 标签。
    返回段落文本列表。
    """
    paragraphs = []
    # 匹配 w:p 元素（可能包含属性）
    para_matches = re.findall(r'<w:p[ >].*?</w:p>', cell_xml, re.DOTALL)
    if not para_matches:
        para_matches = re.findall(r'<w:p\s*/>', cell_xml, re.DOTALL)

    for p_xml in para_matches:
        runs = re.findall(r'<w:r[ >].*?</w:r>', p_xml, re.DOTALL)
        if not runs:
            # 空段落
            paragraphs.append('')
            continue

        run_texts = []
        for r_xml in runs:
            # 检查格式
            is_bold = bool(re.search(r'<w:b\s*/>', r_xml) or re.search(r'<w:b[ >]', r_xml))
            is_italic = bool(re.search(r'<w:i\s*/>', r_xml) or re.search(r'<w:i[ >]', r_xml))

            # 提取文本（w:t 元素的内容）
            texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', r_xml)
            text = ''.join(texts)

            if not text:
                continue

            if is_bold:
                text = f'<b>{text}</b>'
            if is_italic:
                text = f'<i>{text}</i>'
            run_texts.append(text)

        paragraphs.append(''.join(run_texts))

    return paragraphs


def _extract_cell_text(cell_xml: str) -> str:
    """从 w:tc XML 提取单元格文本，多段落用 <br> 分隔。"""
    paragraphs = _extract_cell_paragraphs(cell_xml)
    return '<br>'.join(paragraphs) if paragraphs else ''


def _parse_table_grid(table_xml: str) -> list:
    """解析 w:tblGrid，返回网格列宽列表（用于参考，单位 dxa）。"""
    grid_match = re.search(r'<w:tblGrid[^>]*>(.*?)</w:tblGrid>', table_xml, re.DOTALL)
    if not grid_match:
        return []
    cols = re.findall(r'<w:gridCol[^>]*w:w="(\d+)"', grid_match.group(1))
    return [int(c) for c in cols]


def parse_table_to_html(table_xml: str) -> str:
    """将单个 w:tbl XML 元素转换为 HTML <table> 字符串。

    正确处理：
    - w:gridSpan → colspan
    - w:vMerge (restart/continue) → rowspan
    - 粗体/斜体格式 → <b>/<i>
    """
    rows_xml = re.findall(r'<w:tr[ >].*?</w:tr>', table_xml, re.DOTALL)
    if not rows_xml:
        return ''

    # ── 第一遍：分析 vMerge，计算 rowspan ──
    # open_merges: col_index → {'start_row': int, 'count': int}
    open_merges: dict = {}
    # rowspans: (row, col) → rowspan_value
    rowspans: dict = {}

    for row_idx, row_xml in enumerate(rows_xml):
        cells_xml = re.findall(r'<w:tc[ >].*?</w:tc>', row_xml, re.DOTALL)
        col_idx = 0

        for cell_xml in cells_xml:
            # colspan
            gs = re.search(r'<w:gridSpan[^>]*w:val="(\d+)"', cell_xml)
            colspan = int(gs.group(1)) if gs else 1

            # vMerge
            vm_restart = bool(re.search(r'<w:vMerge[^>]*w:val="restart"', cell_xml))
            has_vmerge = '<w:vMerge' in cell_xml
            vm_continue = has_vmerge and not vm_restart

            if vm_restart:
                open_merges[col_idx] = {'start_row': row_idx, 'count': 1}
            elif vm_continue:
                if col_idx in open_merges:
                    open_merges[col_idx]['count'] += 1
            else:
                # 没有 vMerge → 关闭该列上打开的合并
                if col_idx in open_merges:
                    info = open_merges.pop(col_idx)
                    if info['count'] > 1:
                        rowspans[(info['start_row'], col_idx)] = info['count']

            col_idx += colspan

        # 检查：该行中有没有某列的 vMerge 继续但未出现 continuation cell
        # 这种情况下合并可能已结束（该行该列没有对应的 continuation cell）
        # 我们在行末检查：若某列在当前行没有出现 continuation，则合并结束
        # （注意：这需要知道该行的完整列布局）

    # 关闭所有仍打开的合并
    for col_idx, info in open_merges.items():
        if info['count'] > 1:
            rowspans[(info['start_row'], col_idx)] = info['count']

    # ── 第二遍：生成 HTML ──
    # 跟踪需要跳过的 continuation cells
    skip_cells: set = set()
    for (start_row, col), rs in rowspans.items():
        for offset in range(1, rs):
            skip_cells.add((start_row + offset, col))

    html_parts = ['<table>']

    for row_idx, row_xml in enumerate(rows_xml):
        html_parts.append('  <tr>')
        cells_xml = re.findall(r'<w:tc[ >].*?</w:tc>', row_xml, re.DOTALL)
        col_idx = 0

        for cell_xml in cells_xml:
            # colspan
            gs = re.search(r'<w:gridSpan[^>]*w:val="(\d+)"', cell_xml)
            colspan = int(gs.group(1)) if gs else 1

            # vMerge
            vm_restart = bool(re.search(r'<w:vMerge[^>]*w:val="restart"', cell_xml))
            has_vmerge = '<w:vMerge' in cell_xml
            vm_continue = has_vmerge and not vm_restart

            # 跳过 continuation cells
            if vm_continue:
                col_idx += colspan
                continue

            # 构建属性
            attrs = []
            if colspan > 1:
                attrs.append(f'colspan="{colspan}"')

            rs = rowspans.get((row_idx, col_idx))
            if rs and rs > 1:
                attrs.append(f'rowspan="{rs}"')

            # 判断是否为表头：第一行 或 单元格内文本全部加粗
            tag = 'td'
            if row_idx == 0:
                tag = 'th'
            else:
                paragraphs = _extract_cell_paragraphs(cell_xml)
                cell_text = ''.join(paragraphs)
                # 如果整个单元格内容被 <b> 包裹，视为表头
                if cell_text.startswith('<b>') and cell_text.endswith('</b>') and cell_text.count('<b>') == 1:
                    tag = 'th'

            # 提取文本
            text = _extract_cell_text(cell_xml)

            attr_str = (' ' + ' '.join(attrs)) if attrs else ''
            html_parts.append(f'    <{tag}{attr_str}>{text}</{tag}>')

            col_idx += colspan

        html_parts.append('  </tr>')

    html_parts.append('</table>')
    result = '\n'.join(html_parts)
    return _merge_adjacent_tags(result)


def extract_tables_from_docx(docx_path: str) -> list:
    """从 .docx 文件的 document.xml 中提取所有表格，返回 HTML 字符串列表。

    返回:
        [html_table_1, html_table_2, ...] — 按文档出现顺序排列。
    """
    with zipfile.ZipFile(docx_path) as z:
        if 'word/document.xml' not in z.namelist():
            return []
        xml_content = z.read('word/document.xml').decode('utf-8')

    table_matches = re.findall(r'<w:tbl[ >].*?</w:tbl>', xml_content, re.DOTALL)
    html_tables = []
    for tbl_xml in table_matches:
        html = parse_table_to_html(tbl_xml)
        if html:
            html_tables.append(html)

    return html_tables


def _merge_adjacent_tags(text: str) -> str:
    """合并相邻的同名 HTML 标签，减少碎片化。

    例如: <b>a</b><b>b</b> → <b>ab</b>
          <i>x</i><i>y</i> → <i>xy</i>
    """
    for tag in ('b', 'i'):
        # 重复直到没有更多合并
        pattern = re.compile(rf'</{tag}><{tag}>')
        while pattern.search(text):
            text = pattern.sub('', text)
    return text


def replace_markdown_tables_with_html(md_content: str, html_tables: list) -> str:
    """将 Markdown 中的表格块替换为 HTML 表格。

    按文档出现顺序一一对应替换。如果 HTML 表格数量少于 Markdown 表格，
    则剩余的 Markdown 表格保持原样。
    """
    if not html_tables:
        return md_content

    # 匹配 Markdown 表格块：
    # - 表头行（以 | 开头，以 | 结尾）
    # - 分隔行（|---|---|）
    # - 数据行（以 | 开头，以 | 结尾；最后一行可能以 EOF 结尾无换行符）
    md_table_re = re.compile(
        r'(?:^\|.*\|\s*\n)'            # header row
        r'^\|[-\s:|]+\|\s*\n'         # separator row
        r'(?:^\|.*\|\s*(?:\n|$))*',    # data rows (last may end with EOF)
        re.MULTILINE
    )

    tables_found = md_table_re.findall(md_content)
    if not tables_found:
        return md_content

    # 一一对应替换
    result = md_content
    for i, md_table in enumerate(tables_found):
        if i < len(html_tables):
            # 确保 HTML 表格前后有空行，以保持 Markdown 可读性
            replacement = '\n\n' + html_tables[i] + '\n\n'
            result = result.replace(md_table, replacement, 1)

    return result


def install_markitdown():
    """确保 markitdown[all] 已安装。"""
    try:
        import markitdown  # noqa: F401
    except ImportError:
        print("[setup] 安装 markitdown[all] ...")
        os.system(f"{sys.executable} -m pip install -q 'markitdown[all]'")


def convert_with_markitdown(input_path: str, output_path: str, keep_data_uris: bool = False):
    """使用 markitdown CLI 转换 .docx → .md。

    注意：
        - 默认行为会截断 base64 图片（变成 data:image/png;base64...）
        - --keep-data-uris 保留完整 base64，但会使 .md 文件膨胀
        - 本脚本推荐不用 --keep-data-uris，改用手动提取图片
    """
    from markitdown import MarkItDown

    md = MarkItDown()
    result = md.convert(input_path)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(result.text_content)

    print(f"[markitdown] 转换完成 → {output_path}")
    return output_path


def extract_images_from_docx(docx_path: str, output_dir: str, doc_stem: str) -> int:
    """从 .docx 中提取图片，按文档出现顺序重命名为 fig_1, fig_2, ...

    图片保存到 output_dir/images/doc_stem/ 目录下（例如：images/论证报告/fig_1.png）。

    返回:
        提取的图片总数。
    """
    target_dir = os.path.join(output_dir, "images", doc_stem)
    os.makedirs(target_dir, exist_ok=True)

    # 1. 获取文档中图片的出现顺序
    image_order = get_image_order_in_document(docx_path)

    # 2. 读取关系文件，建立 rId → 原始文件名 映射
    rid_to_filename = {}
    with zipfile.ZipFile(docx_path) as z:
        rels_path = 'word/_rels/document.xml.rels'
        if rels_path in z.namelist():
            root = ET.fromstring(z.read(rels_path))
            for rel in root:
                rid = rel.get('Id')
                target = rel.get('Target', '')
                if 'image' in target.lower():
                    rid_to_filename[rid] = os.path.basename(target)

    # 3. 按文档顺序提取并重命名图片
    extracted = []
    with zipfile.ZipFile(docx_path) as z:
        for idx, (name, rid, _) in enumerate(image_order, 1):
            if rid and rid in rid_to_filename:
                original_filename = rid_to_filename[rid]
                ext = os.path.splitext(original_filename)[1] or '.png'
                new_name = f"fig_{idx}{ext}"
                out_path = os.path.join(target_dir, new_name)

                source_path = f"word/media/{original_filename}"
                if source_path in z.namelist():
                    with z.open(source_path) as src, open(out_path, 'wb') as dst:
                        dst.write(src.read())
                    extracted.append((new_name, os.path.getsize(out_path)))

    print(f"[images] 提取 {len(extracted)} 张图片 → {target_dir}/")
    for name, size in extracted:
        print(f"         {name} ({size:,} bytes)")

    return len(image_order)


def get_image_order_in_document(docx_path: str) -> list:
    """解析 document.xml，返回文档中图片出现的顺序列表。

    返回:
        [(alt_text, embed_rId), ...] — 按文档出现顺序排列。
    """
    image_order = []
    with zipfile.ZipFile(docx_path) as z:
        content = z.read('word/document.xml').decode('utf-8')

    # 匹配 wp:inline 或 wp:anchor 中的图片引用
    for m in re.finditer(r'<wp:(?:inline|anchor).*?</wp:(?:inline|anchor)>', content, re.DOTALL):
        snippet = m.group()
        name_m = re.search(r'name="([^"]*)"', snippet)
        embed_m = re.search(r'r:embed="([^"]*)"', snippet)
        name = name_m.group(1) if name_m else "图片"
        rid = embed_m.group(1) if embed_m else None
        image_order.append((name, rid, m.start()))

    return image_order


def fix_image_references(md_path: str, docx_path: str, doc_stem: str):
    """将 markdown 中截断的 base64 图片引用替换为本地文件引用。

    图片引用格式：images/doc_stem/fig_N.ext（例如：images/论证报告/fig_1.png）

    策略：
        1. 如果 markdown 中已有完整 base64，跳过（不替换）
        2. 如果是截断的 base64（以 ... 结尾），从 docx 提取图片并替换为文件引用
        3. 按文档中图片出现的顺序，将 markdown 中的图片引用一一对应替换
    """
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    # 统计需要修复的图片
    truncated = TRUNCATED_BASE64_RE.findall(content)
    full_base64 = IMAGE_RE.findall(content)

    if not truncated and not full_base64:
        print("[images] 没有需要修复的图片引用")
        return

    if full_base64:
        print(f"[images] 检测到 {len(full_base64)} 个完整 base64 图片（已内嵌，跳过修复）")
        return

    print(f"[images] 检测到 {len(truncated)} 个截断的 base64 引用，开始修复...")

    # 获取文档中图片顺序
    doc_images = get_image_order_in_document(docx_path)

    # 获取 rId → 原始文件名映射
    rid_to_filename = {}
    rels_path = 'word/_rels/document.xml.rels'
    with zipfile.ZipFile(docx_path) as z:
        if rels_path in z.namelist():
            root = ET.fromstring(z.read(rels_path))
            for rel in root:
                rid = rel.get('Id')
                target = rel.get('Target', '')
                if 'image' in target.lower():
                    rid_to_filename[rid] = os.path.basename(target)

    # 按文档顺序构建图片引用路径列表：doc_stem/fig_N.ext
    image_refs = []
    for idx, (name, rid, _) in enumerate(doc_images, 1):
        if rid and rid in rid_to_filename:
            ext = os.path.splitext(rid_to_filename[rid])[1] or '.png'
            img_ref = f"images/{doc_stem}/fig_{idx}{ext}"
            image_refs.append((name, img_ref))

    # 替换：按出现顺序，将截断的 base64 引用替换为本地文件引用
    def replacer(match, idx=[0]):
        alt = match.group(1)
        if idx[0] < len(image_refs):
            doc_name, img_ref = image_refs[idx[0]]
            # 优先使用 markdown 中的 alt text，其次用文档中的名字
            img_alt = alt if alt and not alt.startswith("图片") else doc_name
            idx[0] += 1
            return f"![{img_alt}]({img_ref})"
        else:
            # 超出预期数量，保留原文
            return match.group(0)

    content = TRUNCATED_BASE64_RE.sub(replacer, content)

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"[images] 已修复 {min(len(truncated), len(image_refs))} 处图片引用")


def run_llm_analysis(md_path: str, model: str, api_base: str, api_key: str):
    """使用 OpenAI 兼容 API 对 markdown 内容进行结构化分析。

    需要安装 openai 包：uv pip install openai
    """
    try:
        from openai import OpenAI
    except ImportError:
        print("[llm] 需要 openai 包，正在安装...")
        os.system(f"{sys.executable} -m pip install -q openai")
        from openai import OpenAI

    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    client = OpenAI(api_key=api_key, base_url=api_base)

    prompt = """请分析以下低空飞行器工程技术专业人才培养方案，提取以下结构化信息：

1. 基本信息：专业名称、专业代码、学制、授予学位
2. 培养目标摘要（200字以内）
3. 学分学时统计：总学分、总学时、实践教学占比
4. 课程体系：专业基础课数量/学分、专业核心课数量/学分、专业拓展课数量/学分
5. 三个培养方向及其对应课程
6. 培养规格要点（列出第5-17条的技术能力关键词）
7. 实践教学条件：列出的实验室/实训室名称
8. 实习基地：列出的企业名称
9. 毕业要求：最低学分要求
10. 与旧版人才培养方案的主要差异（若能识别）

请用中文回答，输出格式为 JSON："""

    print(f"[llm] 正在使用 {model} 分析文档内容...")

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": "你是一位职业教育专业建设专家，擅长分析人才培养方案。请用中文回答，输出严格的 JSON 格式。"},
            {"role": "user", "content": f"{prompt}\n\n文档内容：\n\n{content[:30000]}"}  # 截断到 30K 字符
        ],
        temperature=0.1,
        max_tokens=4096,
    )

    result = response.choices[0].message.content
    print(f"[llm] 分析完成")
    print(result)
    return result


def main():
    parser = argparse.ArgumentParser(
        description="将 .docx 转换为 Markdown，提取图片，可选 LLM 分析",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="输入的 .docx 文件路径")
    parser.add_argument("--output-dir", "-o", default=None,
                        help="输出目录（默认：<input_dir>/../参考材料_MD/）")
    parser.add_argument("--no-images", action="store_true",
                        help="不提取和修复图片")
    parser.add_argument("--no-html-tables", action="store_true",
                        help="禁用 HTML 表格转换（使用 markitdown 原生 Markdown 表格）")
    parser.add_argument("--llm-model", default=None,
                        help="LLM 模型名称（如 deepseek-v4-flash），启用 LLM 分析")
    parser.add_argument("--llm-api-base", default="https://api.deepseek.com",
                        help="LLM API 端点（默认: https://api.deepseek.com）")
    parser.add_argument("--llm-api-key", default=None,
                        help="LLM API 密钥（默认读 DEEPSEEK_API_KEY 环境变量）")

    args = parser.parse_args()

    # 验证输入
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"[error] 文件不存在: {args.input}")
        sys.exit(1)
    if input_path.suffix.lower() != '.docx':
        print(f"[error] 仅支持 .docx 文件，收到: {input_path.suffix}")
        sys.exit(1)

    # 确定输出路径
    if args.output_dir:
        output_dir = Path(args.output_dir)
    else:
        # 默认：.docx 所在目录的兄弟目录 参考材料_MD/
        output_dir = input_path.parent

    output_dir.mkdir(parents=True, exist_ok=True)
    output_name = input_path.stem + ".md"
    output_path = output_dir / output_name

    # 图片输出目录（在 .md 同目录下，以文档名命名的子文件夹）
    doc_stem = input_path.stem

    # ── Step 1: 安装依赖 ──
    install_markitdown()

    # ── Step 2: markitdown 转换 ──
    print(f"\n{'='*60}")
    print(f"[convert] 输入: {input_path}")
    print(f"[convert] 输出: {output_path}")
    print(f"{'='*60}\n")

    # 不使用 --keep-data-uris，避免文件膨胀
    convert_with_markitdown(str(input_path), str(output_path), keep_data_uris=False)

    # ── Step 3: 提取图片 ──
    if not args.no_images:
        extract_images_from_docx(str(input_path), str(output_dir), doc_stem)
        # 图片引用修复：使用 doc_stem/fig_N.ext 格式
        fix_image_references(str(output_path), str(input_path), doc_stem)

    # ── Step 4: HTML 表格转换 ──
    if not args.no_html_tables:
        html_tables = extract_tables_from_docx(str(input_path))
        if html_tables:
            with open(output_path, "r", encoding="utf-8") as f:
                md_content = f.read()
            md_content = replace_markdown_tables_with_html(md_content, html_tables)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(md_content)
            print(f"[tables] 已转换 {len(html_tables)} 个表格为 HTML 格式")
        else:
            print("[tables] 文档中没有表格")

    # ── Step 5: LLM 分析（可选）──
    if args.llm_model:
        api_key = args.llm_api_key or os.environ.get("DEEPSEEK_API_KEY")
        if not api_key:
            print("[llm] 未设置 API 密钥。请设置 DEEPSEEK_API_KEY 环境变量或使用 --llm-api-key")
        else:
            run_llm_analysis(str(output_path), args.llm_model, args.llm_api_base, api_key)

    # ── 完成 ──
    print(f"\n{'='*60}")
    print(f"[done] 转换完成！")
    print(f"  Markdown: {output_path}")
    if not args.no_images:
        print(f"  图片:     {output_dir / 'images' / doc_stem}/")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
