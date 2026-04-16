"""Word 文档解析：提取标题、段落、表格为结构化 sections。"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document
from docx.document import Document as _DocType
from docx.table import Table
from docx.text.paragraph import Paragraph


def _iter_block_items(doc: _DocType):
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)


def _heading_level_from_style_name(style_name: str) -> int:
    raw = (style_name or "").strip().lower()
    if not raw:
        return 0
    # 常见英文/中文样式名：Heading 1、标题 1
    for token in raw.replace("heading", "").replace("标题", "").split():
        if token.isdigit():
            lv = int(token)
            return min(9, max(1, lv))
    digits = "".join(ch for ch in raw if ch.isdigit())
    if digits:
        lv = int(digits)
        return min(9, max(1, lv))
    return 0


def parse_docx(docx_path: str | Path) -> dict[str, Any]:
    path = Path(docx_path)
    if not path.is_file():
        raise RuntimeError("Word 文件不存在，无法解析。")
    doc = Document(str(path))
    sections: list[dict[str, Any]] = []
    para_index = 0
    table_index = 0

    for blk in _iter_block_items(doc):
        if isinstance(blk, Paragraph):
            text = (blk.text or "").strip()
            if not text:
                continue
            para_index += 1
            style_name = blk.style.name if blk.style is not None else ""
            level = _heading_level_from_style_name(style_name)
            row: dict[str, Any] = {
                "index": len(sections) + 1,
                "source": "paragraph",
                "paragraph_index": para_index,
                "text": text,
                "style": style_name or "",
            }
            if level > 0:
                row["type"] = "heading"
                row["level"] = level
            else:
                row["type"] = "paragraph"
            sections.append(row)
            continue

        if isinstance(blk, Table):
            table_index += 1
            rows: list[list[str]] = []
            max_cols = 0
            for r in blk.rows:
                vals = [(c.text or "").strip() for c in r.cells]
                max_cols = max(max_cols, len(vals))
                rows.append(vals)
            sections.append(
                {
                    "index": len(sections) + 1,
                    "type": "table",
                    "source": "table",
                    "table_index": table_index,
                    "row_count": len(rows),
                    "col_count": max_cols,
                    "rows": rows,
                },
            )

    return {
        "section_count": len(sections),
        "sections": sections,
    }

