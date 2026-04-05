"""生成历史对应的回填成品 .pptx 磁盘缓存（用于保留期届满时一并删除）。"""
from __future__ import annotations

import logging
import re
from pathlib import Path

from pptx import Presentation

from ppt_report import config
from ppt_report.services.presentation_cache import resolve_template_path
from ppt_report.services.pptx_document import apply_generation_to_presentation

log = logging.getLogger(__name__)

_SAFE_ID = re.compile(r"^[0-9a-fA-F-]{36}$")


def _path_for_id(history_id: str) -> Path | None:
    hid = (history_id or "").strip()
    if not _SAFE_ID.match(hid):
        return None
    return config.FILLED_EXPORT_DIR / f"{hid}.pptx"


def save_filled_export(history_id: str, task_id: str | None, generated: dict) -> bool:
    """将回填后的演示文稿写入 filled_exports/{history_id}.pptx。"""
    if not isinstance(generated, dict):
        return False
    out = _path_for_id(history_id)
    if out is None:
        return False
    tpl = resolve_template_path(task_id)
    if not tpl or not tpl.is_file():
        return False
    try:
        prs = Presentation(str(tpl))
        apply_generation_to_presentation(prs, generated)
        config.FILLED_EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        prs.save(str(out))
        return True
    except Exception:  # noqa: BLE001
        log.warning("回填成品缓存写入失败 history_id=%s", history_id, exc_info=True)
        return False


def delete_filled_export(history_id: str) -> None:
    p = _path_for_id(history_id)
    if p is None or not p.is_file():
        return
    try:
        if p.resolve().parent != config.FILLED_EXPORT_DIR.resolve():
            return
    except OSError:
        return
    try:
        p.unlink()
    except OSError:
        pass
