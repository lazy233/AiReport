"""REST 风格 API：解析、生成、导出。"""
from __future__ import annotations

import io
import json
import mimetypes
import threading
import uuid
from typing import Any
from pathlib import Path

from flask import Blueprint, jsonify, request, send_file
from pptx import Presentation
from werkzeug.utils import secure_filename

from ppt_report import config
from ppt_report import state
from ppt_report.models import db as db_mod
from ppt_report.services.async_jobs import (
    _create_generate_job_record,
    _create_parse_job_record,
    _generate_job_worker,
    _parse_job_worker,
    snapshot_generate_job,
    snapshot_parse_job,
)
from ppt_report.services.generate_pipeline import merge_extra_from_upload, run_generate, validate_generate_prerequisites
from ppt_report.services.chapter_ref_images import (
    delete_chapter_ref_image,
    image_file_path,
    is_safe_task_id,
    save_chapter_ref_image,
)
from ppt_report.services.chapter_reference_resolve import resolve_chapter_reference
from ppt_report.services.page_types import compute_chapter_selection_groups
from ppt_report.services.presentation_cache import (
    bump_parsed_file_name_in_cache,
    get_parsed_from_cache,
    purge_presentation,
    resolve_template_path,
)
from ppt_report.services.pptx_document import apply_generation_to_presentation
from ppt_report.utils.files import allowed_file

api_bp = Blueprint("api", __name__, url_prefix="/api")


@api_bp.get("/presentations")
def api_presentations_list():
    return jsonify(
        {
            "ok": True,
            "items": db_mod.list_presentation_summaries(),
            "db_enabled": db_mod.db_enabled(),
        },
    )


@api_bp.get("/student-data")
def api_student_data_list():
    q = (request.args.get("q") or "").strip()
    items = db_mod.list_student_records(query=q)
    return jsonify({"ok": True, "items": items, "count": len(items)})


@api_bp.get("/student-data/<record_id>")
def api_student_data_get(record_id: str):
    item = db_mod.get_student_record(record_id)
    if not item:
        return jsonify({"ok": False, "error": "未找到该条学生数据。"}), 404
    return jsonify({"ok": True, "item": item})


@api_bp.post("/student-data")
def api_student_data_create():
    body: dict[str, Any] = request.get_json(silent=True) or {}
    try:
        record_id, item = db_mod.save_student_record(body)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception:
        return jsonify({"ok": False, "error": "保存失败，请稍后重试。"}), 500
    return jsonify({"ok": True, "id": record_id, "item": item})


@api_bp.put("/student-data/<record_id>")
def api_student_data_update(record_id: str):
    body: dict[str, Any] = request.get_json(silent=True) or {}
    body["id"] = (record_id or "").strip()
    try:
        saved_id, item = db_mod.save_student_record(body)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception:
        return jsonify({"ok": False, "error": "保存失败，请稍后重试。"}), 500
    return jsonify({"ok": True, "id": saved_id, "item": item})


@api_bp.delete("/student-data/<record_id>")
def api_student_data_delete(record_id: str):
    ok = db_mod.delete_student_record(record_id)
    if not ok:
        return jsonify({"ok": False, "error": "未找到该条学生数据。"}), 404
    return jsonify({"ok": True})


@api_bp.get("/chapter-templates")
def api_chapter_templates_list():
    q = (request.args.get("q") or "").strip()
    items = db_mod.list_chapter_templates(query=q)
    return jsonify({"ok": True, "items": items, "count": len(items)})


@api_bp.get("/chapter-templates/<template_id>")
def api_chapter_templates_get(template_id: str):
    item = db_mod.get_chapter_template(template_id)
    if not item:
        return jsonify({"ok": False, "error": "未找到该模板。"}), 404
    return jsonify({"ok": True, "item": item})


@api_bp.post("/chapter-templates")
def api_chapter_templates_create():
    body: dict[str, Any] = request.get_json(silent=True) or {}
    try:
        template_id, item = db_mod.save_chapter_template(body)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception:
        return jsonify({"ok": False, "error": "保存失败，请稍后重试。"}), 500
    return jsonify({"ok": True, "id": template_id, "item": item})


@api_bp.put("/chapter-templates/<template_id>")
def api_chapter_templates_update(template_id: str):
    body: dict[str, Any] = request.get_json(silent=True) or {}
    body["id"] = (template_id or "").strip()
    try:
        saved_id, item = db_mod.save_chapter_template(body)
    except KeyError:
        return jsonify({"ok": False, "error": "未找到该模板。"}), 404
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception:
        return jsonify({"ok": False, "error": "保存失败，请稍后重试。"}), 500
    return jsonify({"ok": True, "id": saved_id, "item": item})


@api_bp.delete("/chapter-templates/<template_id>")
def api_chapter_templates_delete(template_id: str):
    ok = db_mod.delete_chapter_template(template_id)
    if not ok:
        return jsonify({"ok": False, "error": "未找到该模板。"}), 404
    return jsonify({"ok": True})


@api_bp.post("/resolve-chapter-reference")
def api_resolve_chapter_reference():
    """
    按章节模板顺序写入各「章」块的标题，并由大模型将学生字段分配到各章。
    请求 JSON：task_id, chapter_template_id, student_data_id；可选 use_llm（默认 true）。
    """
    body = request.get_json(silent=True) or {}
    task_id = (body.get("task_id") or "").strip()
    chapter_template_id = (body.get("chapter_template_id") or "").strip()
    student_data_id = (body.get("student_data_id") or "").strip()
    use_llm = body.get("use_llm")
    if use_llm is None:
        use_llm = True
    elif isinstance(use_llm, str):
        use_llm = use_llm.lower() not in ("0", "false", "no")
    else:
        use_llm = bool(use_llm)
    try:
        data = resolve_chapter_reference(
            task_id,
            chapter_template_id,
            student_data_id,
            use_llm=use_llm,
        )
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except RuntimeError as e:
        return jsonify({"ok": False, "error": str(e)}), 503
    except Exception:
        return jsonify({"ok": False, "error": "解析失败，请稍后重试。"}), 500
    return jsonify({"ok": True, "data": data})


@api_bp.post("/chapter-ref-screenshot")
def api_chapter_ref_screenshot_upload():
    """上传某一章的参考截图（按 task_id 分目录存储）。"""
    task_id = (request.form.get("task_id") or "").strip()
    if not get_parsed_from_cache(task_id):
        return jsonify({"ok": False, "error": "未找到该 PPT 解析记录，无法上传。"}), 404
    file = request.files.get("file")
    item, err = save_chapter_ref_image(task_id, file)
    if err:
        code = 413 if "不能超过" in err else 400
        return jsonify({"ok": False, "error": err}), code
    return jsonify({"ok": True, "item": item})


@api_bp.delete("/chapter-ref-screenshot/<task_id>/<filename>")
def api_chapter_ref_screenshot_delete(task_id: str, filename: str):
    """删除已上传的参考截图。"""
    ok = delete_chapter_ref_image(task_id.strip(), filename.strip())
    if not ok:
        return jsonify({"ok": False, "error": "未找到该文件。"}), 404
    return jsonify({"ok": True})


@api_bp.get("/chapter-ref-images/<task_id>/<filename>")
def api_chapter_ref_image_get(task_id: str, filename: str):
    """读取章节参考截图（供预览与后续填入 PPT）。"""
    path = image_file_path(task_id.strip(), filename.strip())
    if not path:
        return jsonify({"ok": False, "error": "未找到文件。"}), 404
    mt = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    return send_file(path, mimetype=mt)


@api_bp.delete("/presentations/<task_id>")
def api_presentation_delete(task_id: str):
    """删除解析记录、模板文件、内存缓存及该任务的章节参考截图。"""
    ok, err = purge_presentation(task_id)
    if not ok:
        code = 404 if err and "未找到" in (err or "") else 400
        return jsonify({"ok": False, "error": err or "删除失败"}), code
    return jsonify({"ok": True})


@api_bp.patch("/presentations/<task_id>")
def api_presentation_rename(task_id: str):
    """修改已解析模板在数据库中的显示文件名（不改动磁盘 .pptx 文件名）。"""
    if not db_mod.db_enabled():
        return jsonify({"ok": False, "error": "数据库未启用，无法保存。"}), 503
    body = request.get_json(silent=True) or {}
    file_name = body.get("file_name")
    if file_name is None:
        file_name = body.get("name")
    if file_name is None:
        return jsonify({"ok": False, "error": "请提供 file_name。"}), 400
    fn = str(file_name).strip()
    if not fn:
        return jsonify({"ok": False, "error": "file_name 不能为空。"}), 400
    ok, err = db_mod.update_presentation_file_name(task_id, fn)
    if not ok:
        code = 404 if err and "未找到" in err else 400
        return jsonify({"ok": False, "error": err or "更新失败"}), code
    tid = (task_id or "").strip()
    row_name = fn[:512]
    bump_parsed_file_name_in_cache(tid, row_name)
    return jsonify({"ok": True, "task_id": tid, "file_name": row_name})


@api_bp.get("/presentations/<task_id>")
def api_presentation_for_generate(task_id: str):
    tid = (task_id or "").strip()
    parsed = get_parsed_from_cache(tid)
    if not parsed:
        return jsonify({"ok": False, "error": "未找到该解析记录。"}), 404
    tpl = resolve_template_path(tid)
    return jsonify(
        {
            "ok": True,
            "task_id": tid,
            "file_name": str(parsed.get("file_name") or ""),
            "slide_count": int(parsed.get("slide_count") or 0),
            "chapter_groups": compute_chapter_selection_groups(parsed),
            "has_template": bool(tpl and tpl.is_file()),
        },
    )


@api_bp.get("/generation_history")
def api_generation_history_list():
    return jsonify(
        {
            "ok": True,
            "db_enabled": db_mod.db_enabled(),
            "items": db_mod.list_generation_history_summaries(),
        },
    )


@api_bp.get("/generation_history/<record_id>")
def api_generation_history_get(record_id: str):
    rec = db_mod.get_generation_history(record_id)
    if not rec:
        return jsonify({"ok": False, "error": "记录不存在。"}), 404
    return jsonify({"ok": True, "record": rec})


@api_bp.delete("/generation_history/<record_id>")
def api_generation_history_delete(record_id: str):
    if not db_mod.db_enabled():
        return jsonify({"ok": False, "error": "数据库未启用。"}), 503
    if not db_mod.delete_generation_history(record_id):
        return jsonify({"ok": False, "error": "记录不存在。"}), 404
    return jsonify({"ok": True})


@api_bp.post("/parse_start")
def api_parse_start():
    file = request.files.get("ppt_file")
    if not file or not file.filename:
        return jsonify({"ok": False, "error": "请先选择一个 PPT 文件。"}), 400
    if not allowed_file(file.filename):
        return jsonify({"ok": False, "error": "只支持 .pptx 或 .ppt 文件。"}), 400
    lower = file.filename.lower()
    if lower.endswith(".ppt") and not lower.endswith(".pptx"):
        return jsonify({"ok": False, "error": "当前仅支持 .pptx 解析，请先将 .ppt 转为 .pptx。"}), 400

    orig_name = secure_filename(file.filename) or "presentation.pptx"
    tmp_path = config.UPLOAD_DIR / f"_parse_tmp_{uuid.uuid4().hex}.pptx"
    file.save(str(tmp_path))

    replace_raw = (request.form.get("replace_task_id") or "").strip()
    replace_tid = replace_raw if replace_raw and is_safe_task_id(replace_raw) else None

    job_poll_id = _create_parse_job_record()
    threading.Thread(
        target=_parse_job_worker,
        args=(job_poll_id, tmp_path, orig_name, replace_tid),
        daemon=True,
    ).start()
    return jsonify({"ok": True, "job_id": job_poll_id})


@api_bp.get("/parse_status/<job_id>")
def api_parse_status(job_id: str):
    snap = snapshot_parse_job(job_id)
    if not snap:
        return jsonify({"ok": False, "error": "任务不存在或已过期。"}), 404
    return jsonify(snap)


@api_bp.post("/generate_start")
def api_generate_start():
    task_id = request.form.get("task_id")
    topic = (request.form.get("topic") or "").strip()
    selected_slides = sorted(
        {
            int(x)
            for x in request.form.getlist("selected_slides")
            if str(x).isdigit()
        }
    )
    chapter_ref: dict | None = None
    raw_ref = (request.form.get("chapter_ref_json") or "").strip()
    if raw_ref:
        try:
            loaded = json.loads(raw_ref)
            if isinstance(loaded, dict):
                chapter_ref = loaded
        except (json.JSONDecodeError, TypeError, ValueError):
            chapter_ref = None
    err, merged_extra = merge_extra_from_upload("", None)
    if err:
        return jsonify({"ok": False, "error": err}), 400
    err = validate_generate_prerequisites(task_id, topic, merged_extra, selected_slides)
    if err:
        return jsonify({"ok": False, "error": err}), 400
    job_id = _create_generate_job_record(task_id)
    threading.Thread(
        target=_generate_job_worker,
        args=(job_id, task_id, topic, merged_extra, selected_slides, chapter_ref),
        daemon=True,
    ).start()
    return jsonify({"ok": True, "job_id": job_id})


@api_bp.get("/generate_status/<job_id>")
def api_generate_status(job_id: str):
    snap = snapshot_generate_job(job_id)
    if not snap:
        return jsonify({"ok": False, "error": "任务不存在或已过期。"}), 404
    return jsonify(snap)


@api_bp.post("/generate")
def api_generate():
    data = request.get_json(silent=True) or {}
    task_id = data.get("task_id")
    topic = data.get("topic") or ""
    extra_content = data.get("extra_content") or ""
    slides = data.get("selected_slides") or data.get("slides") or []
    selected_slides = sorted(
        {
            int(x)
            for x in slides
            if isinstance(x, int) or (isinstance(x, str) and x.strip().isdigit())
        }
    )

    chapter_ref = data.get("chapter_ref")
    if not isinstance(chapter_ref, dict):
        chapter_ref = None
    if chapter_ref is None:
        crj = data.get("chapter_ref_json")
        if isinstance(crj, str) and crj.strip():
            try:
                loaded = json.loads(crj)
                if isinstance(loaded, dict):
                    chapter_ref = loaded
            except (json.JSONDecodeError, TypeError, ValueError):
                chapter_ref = None

    err, result, _ = run_generate(
        task_id, topic, extra_content, selected_slides, None, chapter_ref=chapter_ref
    )
    if err:
        return jsonify({"ok": False, "error": err}), 400
    return jsonify({"ok": True, "data": result})


@api_bp.post("/export")
def api_export():
    body = request.get_json(silent=True) or {}
    task_id = body.get("task_id")
    generated = body.get("data") or (state.LAST_GENERATION.get(task_id) if task_id else None)
    template_path = resolve_template_path(task_id) if task_id else None
    if not task_id:
        return jsonify({"ok": False, "error": "缺少 task_id。"}), 400
    if not template_path or not template_path.is_file():
        return jsonify({"ok": False, "error": "未找到模板文件。"}), 400
    if not generated:
        return jsonify({"ok": False, "error": "请先生成文档内容，或在 body 中传入 data。"}), 400
    prs = Presentation(str(template_path))
    apply_generation_to_presentation(prs, generated)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    meta = get_parsed_from_cache(task_id) or {}
    raw_name = meta.get("file_name") or "presentation.pptx"
    stem = Path(raw_name).stem
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{stem}_filled.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


# 与历史路径一致：/export/<task_id> 挂在应用根上，不单挂载在 /api
export_bp = Blueprint("export", __name__)


@export_bp.get("/export/<task_id>")
def export_filled_pptx(task_id: str):
    template_path = resolve_template_path(task_id)
    generated = state.LAST_GENERATION.get(task_id)
    if not template_path or not template_path.is_file():
        return "未找到该任务的模板，请重新上传并解析。", 404
    if not generated:
        return "请先生成文档内容，再下载填充后的 PPT。", 400
    prs = Presentation(str(template_path))
    apply_generation_to_presentation(prs, generated)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    meta = get_parsed_from_cache(task_id) or {}
    raw_name = meta.get("file_name") or "presentation.pptx"
    stem = Path(raw_name).stem
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{stem}_filled.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
