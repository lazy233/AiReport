"""生成入口：合并上传材料、校验、调用模型。"""
from __future__ import annotations

import json
import logging

from ppt_report import state
from ppt_report.models import db
from ppt_report.services.filled_export_cache import save_filled_export
from ppt_report.services.chapter_reference_resolve import ppt_reference_slot_rows
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


def cover_home_title_validation_error(
    task_id: str | None,
    selected_slides: list[int],
    chapter_ref: dict | None,
) -> str | None:
    """已勾选首页且携带 chapter_ref 时，要求 slots[0].templateTitle 非空（报告主标题）。"""
    if not chapter_ref or not isinstance(chapter_ref, dict):
        return None
    slots = chapter_ref.get("slots")
    if not isinstance(slots, list) or not slots or not isinstance(slots[0], dict):
        return None
    parsed = get_parsed_from_cache(task_id)
    if not parsed:
        return None
    slot_rows = ppt_reference_slot_rows(parsed)
    if not slot_rows or str(slot_rows[0].get("kind") or "") != "cover":
        return None
    cover_pages: set[int] = set()
    for x in slot_rows[0].get("slides") or []:
        s = str(x).strip()
        if s.isdigit():
            try:
                cover_pages.add(int(s))
            except ValueError:
                pass
    selected_set = {int(x) for x in selected_slides}
    if not (cover_pages & selected_set):
        return None
    if str(slots[0].get("templateTitle") or "").strip():
        return None
    return (
        "已纳入首页幻灯片生成：请在「首页」Tab 的「模板章节名」填写报告主标题（生成后将强制写入封面标题文本框）。"
    )


def history_topic_for_record(
    task_id: str | None,
    chapter_ref: dict | None,
    form_topic: str,
) -> str:
    """生成历史「主题」列：优先首页 Tab 填写的报告主标题（templateTitle），否则用表单「额外条件」中的 topic。"""
    ft = (form_topic or "").strip()
    if not chapter_ref or not isinstance(chapter_ref, dict):
        return ft[:20000]
    slots = chapter_ref.get("slots")
    if not isinstance(slots, list):
        return ft[:20000]
    parsed = get_parsed_from_cache(task_id)
    if not parsed:
        return ft[:20000]
    slot_rows = ppt_reference_slot_rows(parsed)
    for i, row in enumerate(slot_rows):
        if str(row.get("kind") or "") != "cover":
            continue
        if i < len(slots) and isinstance(slots[i], dict):
            tt = str(slots[i].get("templateTitle") or "").strip()
            if tt:
                return tt[:20000]
        break
    return ft[:20000]


def validate_generate_prerequisites(
    task_id: str | None,
    topic: str,
    merged_extra: str,
    selected_slides: list[int],
    chapter_ref: dict | None = None,
) -> str | None:
    if not get_parsed_from_cache(task_id):
        return "解析结果已失效，请重新上传并解析。"
    if not selected_slides:
        return "请至少选择一页进行生成。"
    cerr = cover_home_title_validation_error(task_id, selected_slides, chapter_ref)
    if cerr:
        return cerr
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
    err = validate_generate_prerequisites(task_id, topic, merged_extra, selected_slides, chapter_ref)
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
        state.LAST_CHAPTER_REF[task_id] = chapter_ref if isinstance(chapter_ref, dict) else {}
        persist_extra = merged_extra
        if chapter_ref:
            try:
                suffix = json.dumps(chapter_ref, ensure_ascii=False)
                persist_extra = (merged_extra + "\n\n【章节参考 JSON】\n" + suffix).strip()[:200000]
            except (TypeError, ValueError):
                persist_extra = merged_extra
        try:
            hid = db.persist_generation_history(
                task_id,
                history_topic_for_record(task_id, chapter_ref, topic),
                selected_slides,
                persist_extra,
                generated,
            )
            if hid:
                save_filled_export(
                    hid,
                    task_id,
                    generated,
                    chapter_ref if isinstance(chapter_ref, dict) else None,
                )
        except Exception:  # noqa: BLE001
            log.warning("生成历史落库失败", exc_info=True)
    return None, generated, merged_extra
