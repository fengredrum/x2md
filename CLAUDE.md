# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

x2md — Convert `.docx` files to Markdown with image extraction and HTML table conversion. Available as an MCP server and Python API. Python >= 3.10, managed with `uv`.

## Commands

```bash
uv sync                          # Install all dependencies
uv run x2md-mcp                  # Start MCP server (stdio transport)
uv run python scripts/docx2md.py <input.docx> [options]  # Legacy CLI
```

No test framework is configured.

## Architecture

**Conversion pipeline** (3-step, see `converter.py:convert`):
1. `markitdown` → raw .md
2. `images.py` → extract embedded images from the .docx ZIP, rename in document order (`fig_1.png`, `fig_2.png`, ...), fix truncated base64 `![](data:image/...;base64...)` references
3. `tables.py` → parse `w:tbl` XML from `word/document.xml` inside the .docx ZIP, convert to HTML `<table>` preserving `colspan`/`rowspan` and bold/italic, then replace markdown tables in the .md output

**Core modules** (`src/x2md/`):

| Module | Role |
|---|---|
| `converter.py` | Public `convert()` API — validates input, orchestrates the 3-step pipeline |
| `images.py` | `extract_images()` + `fix_references()` — reads `word/document.xml` and `word/_rels/document.xml.rels` from the .docx ZIP, maps `r:embed` rIds to `word/media/` files, renames by document order |
| `tables.py` | `extract_tables()` + `replace_markdown_tables()` + `parse_table_to_html()` — regex-based DOCX XML parser (not lxml), two-pass: first analyze `w:vMerge` for rowspan, then generate HTML with `<th>` detection (first row or all-bold cell) |
| `mcp_server.py` | FastMCP server exposing `convert_docx_to_markdown` tool — thin wrapper around `converter.convert()`, builds a summary string as the return value |

**Legacy script** (`scripts/docx2md.py`): Standalone CLI with duplicated internal logic from before the package split. Has an extra `--llm-model` flag for LLM analysis via OpenAI-compatible API (DeepSeek). Prefer the package API for new work.

**Entry point**: `x2md-mcp` console script → `x2md.mcp_server:main` (defined in `pyproject.toml`).

## Key details

- `.docx` files are ZIP archives — the code reads them via `zipfile.ZipFile`, not python-docx
- Image references in the output use relative paths: `images/<doc_stem>/fig_N.ext`
- Table parsing uses regex on raw XML; `w:gridSpan` → colspan, `w:vMerge` (restart/continue) → rowspan
- The MCP server uses stdio transport — configured in `~/.claude.json` (or `settings.json`) with `"type": "stdio"` (required field, silent failure without it)
- uv uses Aliyun PyPI mirror: `https://mirrors.aliyun.com/pypi/simple`
