"""页面模板路由。"""
from __future__ import annotations

import json

from flask import Blueprint, redirect, render_template, request, url_for

from ppt_report import config
from ppt_report.models import db as db_mod
from ppt_report.services.page_types import compute_chapter_selection_groups, demo_chapter_selection_groups
from ppt_report.services.presentation_cache import get_parsed_from_cache, is_word_stored_payload

web_bp = Blueprint("web", __name__)


@web_bp.get("/")
def index():
    return redirect(url_for("web.overview"))


@web_bp.get("/upload")
def assistant_upload():
    return render_template(
        "pages/assistant_upload.html",
        max_upload_mb=config.MAX_UPLOAD_MB,
        generation_hard_length_cap=config.GENERATION_HARD_LENGTH_CAP,
    )


@web_bp.get("/generate")
def assistant_generate():
    return render_template(
        "pages/assistant_generate.html",
        max_upload_mb=config.MAX_UPLOAD_MB,
        generation_hard_length_cap=config.GENERATION_HARD_LENGTH_CAP,
    )


@web_bp.get("/presentations/<task_id>")
def presentation_detail(task_id: str):
    tid = (task_id or "").strip()
    parsed = get_parsed_from_cache(tid)
    if not parsed:
        return (
            render_template(
                "pages/presentation_missing.html",
                task_id=tid,
                max_upload_mb=config.MAX_UPLOAD_MB,
            ),
            404,
        )
    is_word = is_word_stored_payload(parsed if isinstance(parsed, dict) else None)
    return render_template(
        "pages/presentation_detail.html",
        task_id=tid,
        parsed=parsed,
        parsed_json=json.dumps(parsed, ensure_ascii=False, indent=2),
        is_word_stored=is_word,
        max_upload_mb=config.MAX_UPLOAD_MB,
        generation_hard_length_cap=config.GENERATION_HARD_LENGTH_CAP,
    )


@web_bp.get("/partials/generate-form")
def partial_generate_form():
    tid = (request.args.get("task_id") or "").strip()
    gen_mode = (request.args.get("gen") or "ppt").strip().lower()
    if not tid:
        if gen_mode == "word":
            return render_template(
                "partials/ppt/_generate_panel.html",
                task_id="",
                chapter_groups=[],
                topic="",
                is_word_template=True,
            )
        return render_template(
            "partials/ppt/_generate_panel.html",
            task_id="",
            chapter_groups=demo_chapter_selection_groups(),
            topic="",
            is_word_template=False,
        )
    parsed = get_parsed_from_cache(tid)
    if not parsed:
        return (
            '<p class="error">未找到该解析记录，或会话已过期。'
            "请在文档模板管理页确认列表中仍有该条目，或重新解析。</p>"
        ), 404
    is_word_template = is_word_stored_payload(parsed if isinstance(parsed, dict) else None)
    chapter_groups = [] if is_word_template else compute_chapter_selection_groups(parsed)
    return render_template(
        "partials/ppt/_generate_panel.html",
        task_id=tid,
        chapter_groups=chapter_groups,
        topic="",
        is_word_template=is_word_template,
    )


@web_bp.get("/overview")
def overview():
    stats = db_mod.get_overview_stats()
    recent: list[dict[str, object]] = []
    chart_data: dict[str, object] = db_mod.get_overview_chart_data()
    if stats.get("db_enabled"):
        recent = db_mod.list_generation_history_summaries(limit=8)
    return render_template(
        "pages/overview.html",
        max_upload_mb=config.MAX_UPLOAD_MB,
        overview=stats,
        chart_data=chart_data,
        recent_generations=recent,
    )


@web_bp.get("/chapter-templates")
def chapter_templates_list():
    return render_template("pages/chapter_templates_list.html", max_upload_mb=config.MAX_UPLOAD_MB)


@web_bp.get("/chapter-templates/new")
def chapter_template_new():
    return render_template(
        "pages/chapter_template_editor.html",
        max_upload_mb=config.MAX_UPLOAD_MB,
        page_mode="new",
        template_id="",
    )


@web_bp.get("/chapter-templates/<template_id>")
def chapter_template_detail(template_id: str):
    return render_template(
        "pages/chapter_template_detail.html",
        max_upload_mb=config.MAX_UPLOAD_MB,
        template_id=(template_id or "").strip(),
    )


@web_bp.get("/chapter-templates/<template_id>/edit")
def chapter_template_edit(template_id: str):
    return render_template(
        "pages/chapter_template_editor.html",
        max_upload_mb=config.MAX_UPLOAD_MB,
        page_mode="edit",
        template_id=(template_id or "").strip(),
    )


@web_bp.get("/student-data")
def student_data_list():
    return render_template("pages/student_data_list.html", max_upload_mb=config.MAX_UPLOAD_MB)


@web_bp.get("/student-data/new")
def student_data_new():
    return render_template(
        "pages/student_data_editor.html",
        max_upload_mb=config.MAX_UPLOAD_MB,
        page_mode="new",
        record_id="",
    )


@web_bp.get("/student-data/<record_id>")
def student_data_detail(record_id: str):
    return render_template(
        "pages/student_data_detail.html",
        max_upload_mb=config.MAX_UPLOAD_MB,
        record_id=(record_id or "").strip(),
    )


@web_bp.get("/student-data/<record_id>/edit")
def student_data_edit(record_id: str):
    return render_template(
        "pages/student_data_editor.html",
        max_upload_mb=config.MAX_UPLOAD_MB,
        page_mode="edit",
        record_id=(record_id or "").strip(),
    )


@web_bp.get("/history")
def generation_history_list():
    return render_template("pages/generation_history_list.html", max_upload_mb=config.MAX_UPLOAD_MB)


@web_bp.get("/history/<record_id>")
def generation_history_detail(record_id: str):
    return render_template(
        "pages/generation_history_detail.html",
        max_upload_mb=config.MAX_UPLOAD_MB,
        record_id=(record_id or "").strip(),
    )
