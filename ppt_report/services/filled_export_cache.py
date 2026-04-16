"""生成历史对应的回填成品缓存（.pptx/.docx）。"""
from __future__ import annotations

import logging
import re
from pathlib import Path

from pptx import Presentation

from ppt_report import config
from ppt_report.services.presentation_cache import get_parsed_from_cache, resolve_template_path
from ppt_report.services.pptx_document import apply_generation_to_presentation

log = logging.getLogger(__name__)

_SAFE_ID = re.compile(r"^[0-9a-fA-F-]{36}$")


def _path_for_id(history_id: str, ext: str = "pptx") -> Path | None:
    hid = (history_id or "").strip()
    if not _SAFE_ID.match(hid):
        return None
    safe_ext = "docx" if str(ext).lower().strip() == "docx" else "pptx"
    return config.FILLED_EXPORT_DIR / f"{hid}.{safe_ext}"


def save_filled_export(
    history_id: str,
    task_id: str | None,
    generated: dict,
    chapter_ref: dict | None = None,
) -> bool:
    """将回填后的演示文稿写入 filled_exports/{history_id}.pptx。"""
    if not isinstance(generated, dict):
        return False
    out = _path_for_id(history_id, "pptx")
    if out is None:
        return False
    tpl = resolve_template_path(task_id)
    if not tpl or not tpl.is_file():
        return False
    try:
        prs = Presentation(str(tpl))
        parsed = get_parsed_from_cache(task_id) if task_id else None
        apply_generation_to_presentation(
            prs,
            generated,
            parsed=parsed,
            chapter_ref=chapter_ref if isinstance(chapter_ref, dict) else None,
            task_id=task_id,
        )
        config.FILLED_EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        prs.save(str(out))
        return True
    except Exception:  # noqa: BLE001
        log.warning("回填成品缓存写入失败 history_id=%s", history_id, exc_info=True)
        return False


def delete_filled_export(history_id: str) -> None:
    for ext in ("pptx", "docx"):
        p = _path_for_id(history_id, ext)
        if p is None or not p.is_file():
            continue
        try:
            if p.resolve().parent != config.FILLED_EXPORT_DIR.resolve():
                continue
        except OSError:
            continue
        try:
            p.unlink()
        except OSError:
            pass


def resolve_filled_export_path(history_id: str, ext: str | None = None) -> Path | None:
    """返回历史缓存成品路径（仅当文件存在时）。"""
    wanted = []
    if ext:
        wanted.append("docx" if str(ext).lower().strip() == "docx" else "pptx")
    else:
        wanted.extend(["pptx", "docx"])
    for one in wanted:
        p = _path_for_id(history_id, one)
        if p is None or not p.is_file():
            continue
        try:
            if p.resolve().parent != config.FILLED_EXPORT_DIR.resolve():
                continue
        except OSError:
            continue
        return p
    return None
