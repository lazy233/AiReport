"""解析 / 生成：后台线程任务与内存队列。"""
from __future__ import annotations

import json
import logging
import shutil
import threading
import time
import uuid
from pathlib import Path

from ppt_report import config
from ppt_report import state
from ppt_report.constants import GENERATE_JOBS_MAX, PARSE_JOBS_MAX
from ppt_report.models import db
from ppt_report.services.generate_pipeline import validate_generate_prerequisites
from ppt_report.services.chapter_ref_images import is_safe_task_id
from ppt_report.services.page_types import classify_page_types_with_bailian
from ppt_report.services.presentation_cache import get_parsed_from_cache, save_parsed_to_cache
from ppt_report.services.pptx_document import parse_pptx
from ppt_report.services.filled_export_cache import save_filled_export
from ppt_report.services.text_generation import generate_text_orchestrated

log = logging.getLogger(__name__)


def _prune_parse_jobs_unlocked() -> None:
    while len(state.PARSE_JOBS) >= PARSE_JOBS_MAX:
        oldest_id = min(state.PARSE_JOBS, key=lambda jid: state.PARSE_JOBS[jid].get("created_at", 0.0))
        state.PARSE_JOBS.pop(oldest_id, None)


def _create_parse_job_record() -> str:
    job_id = str(uuid.uuid4())
    with state.PARSE_JOBS_LOCK:
        _prune_parse_jobs_unlocked()
        state.PARSE_JOBS[job_id] = {
            "status": "running",
            "created_at": time.time(),
            "phase": "queued",
            "message": "任务已排队…",
            "task_id": None,
            "error": None,
        }
    return job_id


def _update_parse_job(job_id: str, **kwargs: object) -> None:
    with state.PARSE_JOBS_LOCK:
        row = state.PARSE_JOBS.get(job_id)
        if not row:
            return
        for k, v in kwargs.items():
            row[k] = v


def snapshot_parse_job(job_id: str) -> dict | None:
    with state.PARSE_JOBS_LOCK:
        row = state.PARSE_JOBS.get(job_id)
        if not row:
            return None
        out: dict = {
            "ok": True,
            "status": row["status"],
            "phase": str(row.get("phase") or ""),
            "message": str(row.get("message") or ""),
        }
        if row["status"] == "done":
            out["task_id"] = row.get("task_id")
        if row["status"] == "error":
            out["error"] = str(row.get("error") or row.get("message") or "")
        return out


def _parse_job_worker(
    job_poll_id: str,
    temp_path: Path,
    orig_filename: str,
    replace_task_id: str | None = None,
) -> None:
    try:
        if not temp_path.is_file():
            raise RuntimeError("上传临时文件丢失，请重新上传。")
        rid = (replace_task_id or "").strip() or None
        if rid:
            if not is_safe_task_id(rid):
                raise RuntimeError("无效的替换目标 ID。")
            if not get_parsed_from_cache(rid):
                raise RuntimeError("未找到要替换的模板记录，可能已删除。")
        _update_parse_job(
            job_poll_id,
            phase="parsing",
            message="正在解析幻灯片结构、文本框与组合形状（本地）…",
        )
        parsed = parse_pptx(temp_path)
        parsed["file_name"] = orig_filename
        _update_parse_job(
            job_poll_id,
            phase="classifying",
            message="正在调用大模型识别每页类型（首页/目录/章节扉页/正文）…",
        )
        try:
            parsed = classify_page_types_with_bailian(parsed)
        except Exception:
            pass
        if rid:
            out_task_id = rid
            dest = config.UPLOAD_DIR / f"{out_task_id}.pptx"
            if dest.exists():
                dest.unlink()
            shutil.move(str(temp_path), str(dest))
            state.PARSE_CACHE[out_task_id] = parsed
            state.TEMPLATE_PATHS[out_task_id] = dest
        else:
            out_task_id = save_parsed_to_cache(parsed)
            dest = config.UPLOAD_DIR / f"{out_task_id}.pptx"
            if dest.exists():
                dest.unlink()
            shutil.move(str(temp_path), str(dest))
            state.TEMPLATE_PATHS[out_task_id] = dest
        try:
            db.persist_parsed_presentation(out_task_id, parsed)
        except Exception as exc:  # noqa: BLE001
            log.warning("解析结果写入数据库失败: %s", exc)
        _update_parse_job(
            job_poll_id,
            status="done",
            phase="done",
            message="解析完成，正在跳转…",
            task_id=out_task_id,
        )
    except Exception as exc:  # noqa: BLE001
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass
        _update_parse_job(
            job_poll_id,
            status="error",
            phase="error",
            error=str(exc),
            message=str(exc),
        )


def _prune_generate_jobs_unlocked() -> None:
    while len(state.GENERATE_JOBS) >= GENERATE_JOBS_MAX:
        oldest_id = min(state.GENERATE_JOBS, key=lambda jid: state.GENERATE_JOBS[jid].get("created_at", 0.0))
        state.GENERATE_JOBS.pop(oldest_id, None)


def _create_generate_job_record(task_id: str | None) -> str:
    job_id = str(uuid.uuid4())
    with state.GENERATE_JOBS_LOCK:
        _prune_generate_jobs_unlocked()
        state.GENERATE_JOBS[job_id] = {
            "status": "running",
            "created_at": time.time(),
            "task_id": task_id,
            "batch_index": 0,
            "batch_total": 0,
            "message": "任务已启动…",
            "result": None,
            "error": None,
        }
    return job_id


def _update_generate_job(job_id: str, **kwargs: object) -> None:
    with state.GENERATE_JOBS_LOCK:
        row = state.GENERATE_JOBS.get(job_id)
        if not row:
            return
        for k, v in kwargs.items():
            row[k] = v


def _generate_job_worker(
    job_id: str,
    task_id: str | None,
    topic: str,
    merged_extra: str,
    selected_slides: list[int],
    chapter_ref: dict | None = None,
) -> None:
    last_batch_total = 1

    def on_progress(batch_index: int, batch_total: int, batch_pages: list[int]) -> None:
        nonlocal last_batch_total
        last_batch_total = max(1, int(batch_total))
        n_pages = len(batch_pages)
        _update_generate_job(
            job_id,
            status="running",
            batch_index=batch_index + 1,
            batch_total=last_batch_total,
            message=(
                f"正在调用大模型（第 {batch_index + 1}/{last_batch_total} 批"
                f"{'' if n_pages <= 0 else f'，本批 {n_pages} 页'}）…"
            ),
        )

    try:
        err = validate_generate_prerequisites(task_id, topic, merged_extra, selected_slides)
        if err:
            raise RuntimeError(err)
        parsed = get_parsed_from_cache(task_id)
        _update_generate_job(
            job_id,
            batch_total=1,
            batch_index=0,
            message="准备按章节调用大模型…",
        )
        generated = generate_text_orchestrated(
            parsed,
            selected_slides,
            topic,
            merged_extra,
            chapter_ref,
            progress=on_progress,
        )
        if task_id:
            state.LAST_GENERATION[task_id] = generated
        history_id: str | None = None
        persist_extra = merged_extra
        if chapter_ref:
            try:
                suffix = json.dumps(chapter_ref, ensure_ascii=False)
                persist_extra = (merged_extra + "\n\n【章节参考 JSON】\n" + suffix).strip()[:200000]
            except (TypeError, ValueError):
                persist_extra = merged_extra
        try:
            history_id = db.persist_generation_history(
                task_id, topic, selected_slides, persist_extra, generated
            )
        except Exception:  # noqa: BLE001
            log.warning("生成历史落库失败（任务仍视为成功）", exc_info=True)
        done_kwargs: dict[str, object] = {
            "status": "done",
            "batch_index": last_batch_total,
            "batch_total": last_batch_total,
            "message": "生成完成",
            "result": generated,
            "task_id": task_id,
        }
        if history_id:
            done_kwargs["history_id"] = history_id
            save_filled_export(history_id, task_id, generated)
        _update_generate_job(job_id, **done_kwargs)
    except Exception as exc:  # noqa: BLE001
        _update_generate_job(
            job_id,
            status="error",
            error=str(exc),
            message=str(exc),
        )


def snapshot_generate_job(job_id: str) -> dict | None:
    with state.GENERATE_JOBS_LOCK:
        row = state.GENERATE_JOBS.get(job_id)
        if not row:
            return None
        out: dict = {
            "ok": True,
            "status": row["status"],
            "batch_index": int(row.get("batch_index") or 0),
            "batch_total": int(row.get("batch_total") or 0),
            "message": str(row.get("message") or ""),
        }
        if row["status"] == "done":
            out["result"] = row.get("result")
            out["task_id"] = row.get("task_id")
            hid = row.get("history_id")
            if hid:
                out["history_id"] = hid
        if row["status"] == "error":
            out["error"] = str(row.get("error") or row.get("message") or "")
        return out


__all__ = [
    "_create_generate_job_record",
    "_create_parse_job_record",
    "_generate_job_worker",
    "_parse_job_worker",
    "snapshot_generate_job",
    "snapshot_parse_job",
]
