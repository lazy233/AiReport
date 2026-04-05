"""生成入口：合并上传材料、校验、调用模型。"""
from __future__ import annotations

import json
import logging

from ppt_report import state
from ppt_report.models import db
from ppt_report.services.filled_export_cache import save_filled_export
from ppt_report.services.presentation_cache import get_parsed_from_cache

log = logging.getLogger(__name__)
from ppt_report.services.text_generation import generate_text_orchestrated
from ppt_report.utils.files import allowed_content_file, decode_uploaded_text


def merge_extra_from_upload(merged_extra: str, content_file) -> tuple[str | None, str]:
    merged = (merged_extra or "").strip()
    if content_file and getattr(content_file, "filename", None):
        if not allowed_content_file(content_file.filename):
            return "数据文件仅支持 .txt/.md/.markdown/.json/.csv。", merged
        raw = content_file.read()
        if raw:
            file_text = decode_uploaded_text(raw).strip()
            if file_text:
                merged = (merged + "\n\n" + file_text).strip() if merged else file_text
    return None, merged


def validate_generate_prerequisites(
    task_id: str | None,
    topic: str,
    merged_extra: str,
    selected_slides: list[int],
) -> str | None:
    if not get_parsed_from_cache(task_id):
        return "解析结果已失效，请重新上传并解析。"
    if not selected_slides:
        return "请至少选择一页进行生成。"
    if not (topic or "").strip() and not (merged_extra or "").strip():
        return "请至少输入主题。"
    return None


def run_generate(
    task_id: str | None,
    topic: str,
    extra_content: str,
    selected_slides: list[int],
    content_file,
    *,
    chapter_ref: dict | None = None,
) -> tuple[str | None, dict | None, str]:
    topic = (topic or "").strip()
    err, merged_extra = merge_extra_from_upload((extra_content or "").strip(), content_file)
    if err:
        return err, None, merged_extra
    err = validate_generate_prerequisites(task_id, topic, merged_extra, selected_slides)
    if err:
        return err, None, merged_extra
    parsed = get_parsed_from_cache(task_id)
    try:
        generated = generate_text_orchestrated(
            parsed, selected_slides, topic, merged_extra, chapter_ref, progress=None
        )
    except Exception as exc:  # noqa: BLE001
        return f"生成失败：{exc}", None, merged_extra
    if task_id and generated:
        state.LAST_GENERATION[task_id] = generated
        persist_extra = merged_extra
        if chapter_ref:
            try:
                suffix = json.dumps(chapter_ref, ensure_ascii=False)
                persist_extra = (merged_extra + "\n\n【章节参考 JSON】\n" + suffix).strip()[:200000]
            except (TypeError, ValueError):
                persist_extra = merged_extra
        try:
            hid = db.persist_generation_history(
                task_id, topic, selected_slides, persist_extra, generated
            )
            if hid:
                save_filled_export(hid, task_id, generated)
        except Exception:  # noqa: BLE001
            log.warning("生成历史落库失败", exc_info=True)
    return None, generated, merged_extra
