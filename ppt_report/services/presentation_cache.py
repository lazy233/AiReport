"""解析结果缓存与模板路径解析。"""
from __future__ import annotations

import uuid
from pathlib import Path
from typing import Any

from ppt_report import config
from ppt_report import state
from ppt_report.models import db
from ppt_report.services.chapter_ref_images import is_safe_task_id, purge_chapter_ref_task_dir


def is_word_stored_payload(parsed: dict[str, Any] | None) -> bool:
    if not isinstance(parsed, dict):
        return False
    kind = str(parsed.get("template_kind") or "").strip()
    return kind in ("word_stored", "word_parsed")


def save_parsed_to_cache(parsed: dict) -> str:
    task_id = uuid.uuid4().hex
    state.PARSE_CACHE[task_id] = parsed
    return task_id


def get_parsed_from_cache(task_id: str | None) -> dict | None:
    if not task_id:
        return None
    cached = state.PARSE_CACHE.get(task_id)
    if cached is not None:
        return cached
    if db.db_enabled():
        loaded = db.load_parsed_presentation(task_id)
        if loaded:
            state.PARSE_CACHE[task_id] = loaded
            tpl = config.UPLOAD_DIR / f"{task_id}.pptx"
            if tpl.is_file():
                state.TEMPLATE_PATHS[task_id] = tpl
            return loaded
    return None


def bump_parsed_file_name_in_cache(task_id: str, file_name: str) -> None:
    """内存缓存中的解析结果与库内显示名同步（若存在）。"""
    tid = (task_id or "").strip()
    if not tid:
        return
    cached = state.PARSE_CACHE.get(tid)
    if not isinstance(cached, dict):
        return
    updated = dict(cached)
    updated["file_name"] = (file_name or "").strip()[:512]
    state.PARSE_CACHE[tid] = updated


def resolve_template_path(task_id: str | None) -> Path | None:
    if not task_id:
        return None
    p = state.TEMPLATE_PATHS.get(task_id)
    if p and p.is_file():
        return p
    tpl = config.UPLOAD_DIR / f"{task_id}.pptx"
    if tpl.is_file():
        state.TEMPLATE_PATHS[task_id] = tpl
        return tpl
    return None


def purge_presentation(task_id: str) -> tuple[bool, str | None]:
    """
    删除库内记录、内存缓存、模板 .pptx、该任务的章节参考截图目录。
    若至少清理到一项则视为成功。
    """
    tid = (task_id or "").strip()
    if not is_safe_task_id(tid):
        return False, "无效的任务 ID。"

    removed = False
    if db.delete_parsed_presentation(tid):
        removed = True
    if state.PARSE_CACHE.pop(tid, None) is not None:
        removed = True
    state.TEMPLATE_PATHS.pop(tid, None)
    state.LAST_GENERATION.pop(tid, None)
    state.LAST_CHAPTER_REF.pop(tid, None)

    for ext in (".pptx", ".docx"):
        dest = config.UPLOAD_DIR / f"{tid}{ext}"
        if dest.is_file():
            try:
                dest.unlink()
                removed = True
            except OSError:
                pass

    if purge_chapter_ref_task_dir(tid):
        removed = True

    if not removed:
        return False, "未找到该模板记录。"
    return True, None
