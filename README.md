# x2md

将 `.docx` 文件转换为 Markdown，支持图片提取和 HTML 表格转换。提供 MCP server 和 Python API 两种使用方式。

## 安装

```bash
uv sync
```

## 使用方式

### 1. MCP Server

启动 MCP server（stdio 传输）：

```bash
uv run x2md-mcp
```

#### MCP 工具

**`convert_docx_to_markdown`**

| 参数 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `input_path` | `string` | 必填 | `.docx` 文件的绝对路径 |
| `output_dir` | `string` | 与文档同目录 | 输出目录，`.md` 和 `images/` 均放在此目录下 |
| `extract_images` | `bool` | `true` | 是否提取嵌入图片并修复引用 |
| `html_tables` | `bool` | `true` | 是否将表格转为 HTML（保留 colspan/rowspan） |

返回：转换摘要，包含输出路径、文件大小、提取的图片数量。

#### 配置 MCP 客户端

在 Claude Code 的 `~/.claude.json`（或 `settings.json`）中添加：

```json
{
  "mcpServers": {
    "x2md": {
      "type": "stdio",
      "command": "uv",
      "args": ["run", "--directory", "/home/oblivion/GitHub/x2md", "x2md-mcp"]
    }
  }
}
```

> **注意**：`"type": "stdio"` 是必填字段，缺少会导致 MCP server 被静默忽略。配置后需重启 Claude Code 才能生效。

### 2. Python API

```python
from x2md.converter import convert

# 基础转换（输出与原文件同目录）
result = convert("/path/to/document.docx")

# 指定输出目录
result = convert("/path/to/document.docx", output_dir="/path/to/output")

# 禁用图片提取和 HTML 表格
result = convert(
    "/path/to/document.docx",
    extract_images_flag=False,
    html_tables=False,
)

print(result)  # → 输出的 .md 文件路径
```

### 3. 原有 CLI

`docx2md.py` 仍可直接使用：

```bash
uv run python docx2md.py <input.docx> [options]
```

## 转换流程

```
.docx  ──[markitdown]──>  .md  ──[图片修复]──>  .md  ──[HTML表格]──>  .md
```

- **图片**：提取到 `images/<文档名>/fig_1.png, fig_2.png, ...`，修复 Markdown 中的引用
- **表格**：解析 DOCX 原始 XML，转为带 `colspan`/`rowspan` 的 HTML `<table>`，粗体/斜体格式保留

## 依赖

- Python >= 3.10
- [markitdown](https://github.com/microsoft/markitdown) — .docx → .md 转换
- [mcp](https://github.com/modelcontextprotocol/python-sdk) — MCP server 框架
- 通过 `uv` 管理所有依赖，使用阿里云 PyPI 镜像
