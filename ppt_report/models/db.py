"""
PostgreSQL 持久化：解析结果全量 JSON（JSONB）。
环境变量 DATABASE_URL，默认连本地库 ppt_report_platform。
"""
from __future__ import annotations

import logging
import re
from datetime import datetime, timedelta, timezone
from typing import Any
from uuid import uuid4

from sqlalchemy import (
    Boolean,
    CheckConstraint,
    DateTime,
    ForeignKey,
    Integer,
    String,
    Text,
    UniqueConstraint,
    create_engine,
    delete,
    func,
    or_,
    select,
    text,
)
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, sessionmaker
from sqlalchemy.orm.attributes import flag_modified

log = logging.getLogger(__name__)

_engine = None
_SessionLocal = None


class Base(DeclarativeBase):
    pass


class ParsedPresentation(Base):
    __tablename__ = "parsed_presentations"

    task_id: Mapped[str] = mapped_column(String(32), primary_key=True)
    file_name: Mapped[str] = mapped_column(String(512), default="")
    slide_count: Mapped[int] = mapped_column(Integer, default=0)
    payload: Mapped[dict[str, Any]] = mapped_column(JSONB, nullable=False)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
    )


class GenerationHistory(Base):
    """文案生成历史（服务端持久化，供列表/详情/再导出）。"""

    __tablename__ = "generation_histories"

    id: Mapped[str] = mapped_column(
        String(36),
        primary_key=True,
        default=lambda: str(uuid4()),
    )
    task_id: Mapped[str | None] = mapped_column(String(32), nullable=True)
    topic: Mapped[str] = mapped_column(Text, nullable=False, default="")
    selected_slides: Mapped[list[Any]] = mapped_column(JSONB, nullable=False)
    merged_extra: Mapped[str] = mapped_column(Text, nullable=False, default="")
    slide_count: Mapped[int] = mapped_column(Integer, default=0)
    result: Mapped[dict[str, Any]] = mapped_column(JSONB, nullable=False)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
    )


class Student(Base):
    """学生主档（稳定信息）。"""

    __tablename__ = "students"

    id: Mapped[str] = mapped_column(
        String(36),
        primary_key=True,
        default=lambda: str(uuid4()),
    )
    student_code: Mapped[str] = mapped_column(String(64), unique=True, nullable=False)
    student_name: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    nickname_en: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    service_start_date: Mapped[str] = mapped_column(String(32), nullable=False, default="")
    planner_teacher: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
    )
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )


class StudentTermProfile(Base):
    """学生学期快照（四维数据）。"""

    __tablename__ = "student_term_profiles"
    __table_args__ = (
        UniqueConstraint("student_id", "term_code", name="uq_student_term_profiles_student_term"),
        CheckConstraint(
            "(total_hours IS NULL OR total_hours >= 0) "
            "AND (used_hours IS NULL OR used_hours >= 0) "
            "AND (remaining_hours IS NULL OR remaining_hours >= 0) "
            "AND (total_hours IS NULL OR used_hours IS NULL OR used_hours <= total_hours)",
            name="ck_student_term_profiles_hours_non_negative",
        ),
    )

    id: Mapped[str] = mapped_column(
        String(36),
        primary_key=True,
        default=lambda: str(uuid4()),
    )
    student_id: Mapped[str] = mapped_column(
        String(36),
        ForeignKey("students.id", ondelete="CASCADE"),
        nullable=False,
    )
    term_code: Mapped[str] = mapped_column(String(32), nullable=False, default="")

    # 基础信息（学期相关）
    grade_level: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    school: Mapped[str] = mapped_column(String(256), nullable=False, default="")
    major: Mapped[str] = mapped_column(String(256), nullable=False, default="")
    report_subtitle: Mapped[str] = mapped_column(String(256), nullable=False, default="")

    # 学习画像
    strong_subjects: Mapped[str] = mapped_column(Text, nullable=False, default="")
    intl_scores: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    study_intent: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    career_intent: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    interest_subjects: Mapped[str] = mapped_column(Text, nullable=False, default="")
    long_term_plan: Mapped[str] = mapped_column(Text, nullable=False, default="")
    learning_style: Mapped[str] = mapped_column(Text, nullable=False, default="")
    weak_areas: Mapped[str] = mapped_column(Text, nullable=False, default="")

    # 课时数据
    total_hours: Mapped[int | None] = mapped_column(Integer, nullable=True)
    used_hours: Mapped[int | None] = mapped_column(Integer, nullable=True)
    remaining_hours: Mapped[int | None] = mapped_column(Integer, nullable=True)
    tutor_subjects: Mapped[str] = mapped_column(Text, nullable=False, default="")
    preview_subjects: Mapped[str] = mapped_column(Text, nullable=False, default="")
    skill_direction: Mapped[str] = mapped_column(Text, nullable=False, default="")
    skill_description: Mapped[str] = mapped_column(Text, nullable=False, default="")

    # 成长指导
    term_summary: Mapped[str] = mapped_column(Text, nullable=False, default="")
    course_feedback: Mapped[str] = mapped_column(Text, nullable=False, default="")
    short_term_advice: Mapped[str] = mapped_column(Text, nullable=False, default="")
    long_term_development: Mapped[str] = mapped_column(Text, nullable=False, default="")

    # 预留扩展
    extra_json: Mapped[dict[str, Any]] = mapped_column(JSONB, nullable=False, default=dict)

    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
    )
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )


class ChapterTemplate(Base):
    """章节模板主表。"""

    __tablename__ = "chapter_templates"

    id: Mapped[str] = mapped_column(
        String(36),
        primary_key=True,
        default=lambda: str(uuid4()),
    )
    template_code: Mapped[str] = mapped_column(String(64), unique=True, nullable=False)
    name: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    description: Mapped[str] = mapped_column(Text, nullable=False, default="")
    is_active: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
    )
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )


class ChapterTemplateChapter(Base):
    """章节模板章节表（一个模板多章节）。"""

    __tablename__ = "chapter_template_chapters"
    __table_args__ = (
        UniqueConstraint("template_id", "sort_order", name="uq_ct_chapter_template_sort"),
        CheckConstraint("length(btrim(title)) > 0", name="ck_ct_chapter_title_non_empty"),
    )

    id: Mapped[str] = mapped_column(
        String(36),
        primary_key=True,
        default=lambda: str(uuid4()),
    )
    template_id: Mapped[str] = mapped_column(
        String(36),
        ForeignKey("chapter_templates.id", ondelete="CASCADE"),
        nullable=False,
    )
    title: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    hint: Mapped[str] = mapped_column(Text, nullable=False, default="")
    sort_order: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
    )
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        default=lambda: datetime.now(timezone.utc),
        onupdate=lambda: datetime.now(timezone.utc),
    )


def _norm_str(v: object, max_len: int | None = None) -> str:
    s = "" if v is None else str(v).strip()
    return s[:max_len] if max_len and max_len > 0 else s


def _pick_profile(profile: dict[str, Any], dim: str, key: str, max_len: int | None = None) -> str:
    src = profile.get(dim) if isinstance(profile, dict) else None
    if not isinstance(src, dict):
        return ""
    return _norm_str(src.get(key), max_len)


def _parse_optional_int(v: object) -> int | None:
    if v is None:
        return None
    if isinstance(v, int):
        return v
    s = _norm_str(v)
    if not s:
        return None
    m = re.search(r"-?\d+", s)
    if not m:
        return None
    try:
        return int(m.group(0))
    except ValueError:
        return None


def _build_profile_dict(student: Student, row: StudentTermProfile) -> dict[str, dict[str, str]]:
    extra = row.extra_json if isinstance(row.extra_json, dict) else {}
    merged = extra.get("profile") if isinstance(extra.get("profile"), dict) else {}
    basic = merged.get("basic") if isinstance(merged.get("basic"), dict) else {}
    learning = merged.get("learning") if isinstance(merged.get("learning"), dict) else {}
    hours = merged.get("hours") if isinstance(merged.get("hours"), dict) else {}
    guidance = merged.get("guidance") if isinstance(merged.get("guidance"), dict) else {}
    basic.update(
        {
            "studentName": _norm_str(student.student_name),
            "studentId": _norm_str(student.student_code),
            "nicknameEn": _norm_str(student.nickname_en),
            "serviceStart": _norm_str(student.service_start_date),
            "plannerTeacher": _norm_str(student.planner_teacher),
            "school": _norm_str(row.school),
            "major": _norm_str(row.major),
            "gradeLevel": _norm_str(row.grade_level),
            "currentTerm": _norm_str(row.term_code),
            "reportSubtitle": _norm_str(row.report_subtitle),
            "className": _norm_str(extra.get("className")),
            "email": _norm_str(extra.get("email")),
            "phone": _norm_str(extra.get("phone")),
            "remark": _norm_str(extra.get("remark")),
        },
    )
    learning.update(
        {
            "strongSubjects": _norm_str(row.strong_subjects),
            "intlScores": _norm_str(row.intl_scores),
            "studyIntent": _norm_str(row.study_intent),
            "careerIntent": _norm_str(row.career_intent),
            "interestSubjects": _norm_str(row.interest_subjects),
            "longTermPlan": _norm_str(row.long_term_plan),
            "learningStyle": _norm_str(row.learning_style),
            "weakAreas": _norm_str(row.weak_areas),
        },
    )
    hours.update(
        {
            "totalHours": "" if row.total_hours is None else str(row.total_hours),
            "usedHours": "" if row.used_hours is None else str(row.used_hours),
            "remainingHours": "" if row.remaining_hours is None else str(row.remaining_hours),
            "tutorSubjects": _norm_str(row.tutor_subjects),
            "previewSubjects": _norm_str(row.preview_subjects),
            "skillDirection": _norm_str(row.skill_direction),
            "skillDescription": _norm_str(row.skill_description),
        },
    )
    guidance.update(
        {
            "termSummary": _norm_str(row.term_summary),
            "courseFeedback": _norm_str(row.course_feedback),
            "shortTermAdvice": _norm_str(row.short_term_advice),
            "longTermDevelopment": _norm_str(row.long_term_development),
        },
    )
    return {"basic": basic, "learning": learning, "hours": hours, "guidance": guidance}


def _record_from_pair(student: Student, row: StudentTermProfile) -> dict[str, object]:
    extra = row.extra_json if isinstance(row.extra_json, dict) else {}
    profile = _build_profile_dict(student, row)
    return {
        "id": row.id,
        "name": _norm_str(student.student_name),
        "studentId": _norm_str(student.student_code),
        "className": _norm_str(extra.get("className")),
        "email": _norm_str(extra.get("email")),
        "phone": _norm_str(extra.get("phone")),
        "remark": _norm_str(extra.get("remark")),
        "content": _norm_str(extra.get("content")),
        "profile": profile,
        "createdAt": row.created_at.isoformat() if row.created_at else "",
        "updatedAt": row.updated_at.isoformat() if row.updated_at else "",
    }


def init_db(database_url: str | None) -> bool:
    """初始化引擎并建表。返回是否已启用数据库。"""
    global _engine, _SessionLocal
    if not database_url or not database_url.strip():
        log.warning("未设置 DATABASE_URL，解析结果不会写入数据库。")
        return False
    try:
        _engine = create_engine(
            database_url.strip(),
            pool_pre_ping=True,
            connect_args={"connect_timeout": 10},
        )
        _SessionLocal = sessionmaker(bind=_engine, autocommit=False, autoflush=False)
        Base.metadata.create_all(bind=_engine)
        with _engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        log.info(
            "数据库已连接，表 parsed_presentations/generation_histories/students/"
            "student_term_profiles/chapter_templates/chapter_template_chapters 就绪。",
        )
        return True
    except Exception:
        log.exception("数据库初始化失败，将仅使用内存缓存。")
        _engine = None
        _SessionLocal = None
        return False


def db_enabled() -> bool:
    return _SessionLocal is not None


def persist_parsed_presentation(task_id: str, parsed: dict) -> None:
    if not _SessionLocal or not task_id or not parsed:
        return
    row = ParsedPresentation(
        task_id=task_id,
        file_name=str(parsed.get("file_name") or "")[:512],
        slide_count=int(parsed.get("slide_count") or 0),
        payload=parsed,
    )
    with _SessionLocal() as session:
        session.merge(row)
        session.commit()


def load_parsed_presentation(task_id: str) -> dict | None:
    if not _SessionLocal or not task_id:
        return None
    with _SessionLocal() as session:
        row = session.get(ParsedPresentation, task_id)
        if row is None:
            return None
        return dict(row.payload) if isinstance(row.payload, dict) else row.payload


def delete_parsed_presentation(task_id: str) -> bool:
    """删除一条解析记录；无库或未找到时返回 False。"""
    tid = _norm_str(task_id, 32)
    if not _SessionLocal or not tid:
        return False
    with _SessionLocal() as session:
        row = session.get(ParsedPresentation, tid)
        if row is None:
            return False
        session.delete(row)
        session.commit()
    return True


def _parse_impl_label_from_payload(payload: Any) -> str:
    """列表展示用：区分本地解析与是否经大模型页类型标注。"""
    if not isinstance(payload, dict):
        return "—"
    slides = payload.get("slides")
    if not isinstance(slides, list) or not slides:
        return "本地解析"
    for s in slides:
        if not isinstance(s, dict):
            continue
        pt = str(s.get("page_type") or "").strip().lower()
        if pt and pt != "unknown":
            return "本地解析 + 大模型页类型"
    return "仅本地解析"


def update_presentation_file_name(task_id: str, display_name: str) -> tuple[bool, str | None]:
    """
    更新解析记录在库中的显示文件名，并同步 payload.file_name。
    不修改磁盘上的 uploads/{task_id}.pptx 文件名。
    """
    tid = _norm_str(task_id, 64)
    name = (display_name or "").strip()
    name = " ".join(name.split())
    if not tid:
        return False, "无效的任务 ID。"
    if not name:
        return False, "文件名不能为空。"
    name = name[:512]
    if not _SessionLocal:
        return False, "数据库未启用，无法保存。"
    with _SessionLocal() as session:
        row = session.get(ParsedPresentation, tid)
        if row is None:
            return False, "未找到该模板记录。"
        row.file_name = name
        pl = dict(row.payload) if isinstance(row.payload, dict) else {}
        pl["file_name"] = name
        row.payload = pl
        flag_modified(row, "payload")
        session.commit()
    return True, None


def persist_generation_history(
    task_id: str | None,
    topic: str,
    selected_slides: list[int],
    merged_extra: str,
    result: dict[str, Any],
) -> str | None:
    """写入一条生成历史，返回记录 id；数据库未启用或失败时返回 None。"""
    if not _SessionLocal or not isinstance(result, dict):
        return None
    slides_payload = result.get("slides")
    slide_count = len(slides_payload) if isinstance(slides_payload, list) else 0
    tid = _norm_str(task_id, 32) or None
    rec_id = str(uuid4())
    row = GenerationHistory(
        id=rec_id,
        task_id=tid,
        topic=(topic or "").strip()[:20000],
        selected_slides=list(selected_slides),
        merged_extra=(merged_extra or "")[:500000],
        slide_count=slide_count,
        result=result,
    )
    try:
        with _SessionLocal() as session:
            session.add(row)
            session.commit()
        return rec_id
    except Exception:
        log.exception("生成历史写入数据库失败")
        return None


def list_generation_history_summaries(limit: int = 200) -> list[dict[str, object]]:
    """生成历史列表摘要（不含 result JSON），按时间倒序。"""
    if not _SessionLocal:
        return []
    lim = max(1, min(int(limit), 2000))
    with _SessionLocal() as session:
        rows = session.scalars(
            select(GenerationHistory).order_by(GenerationHistory.created_at.desc()).limit(lim),
        ).all()
    out: list[dict[str, object]] = []
    for r in rows:
        created = r.created_at
        out.append(
            {
                "id": r.id,
                "taskId": r.task_id or "",
                "topic": r.topic or "",
                "createdAt": created.isoformat() if created else "",
                "slideCount": int(r.slide_count or 0),
            },
        )
    return out


def get_generation_history(record_id: str) -> dict[str, object] | None:
    """单条历史（含 result，供详情与导出）。"""
    rid = _norm_str(record_id, 36)
    if not _SessionLocal or not rid:
        return None
    with _SessionLocal() as session:
        row = session.get(GenerationHistory, rid)
        if row is None:
            return None
        created = row.created_at
        res = row.result if isinstance(row.result, dict) else {}
        return {
            "id": row.id,
            "taskId": row.task_id or "",
            "topic": row.topic or "",
            "createdAt": created.isoformat() if created else "",
            "selectedSlides": list(row.selected_slides) if isinstance(row.selected_slides, list) else [],
            "result": res,
        }


def delete_generation_history(record_id: str) -> bool:
    from ppt_report.services.filled_export_cache import delete_filled_export

    rid = _norm_str(record_id, 36)
    if not _SessionLocal or not rid:
        return False
    with _SessionLocal() as session:
        row = session.get(GenerationHistory, rid)
        if row is None:
            return False
        session.delete(row)
        session.commit()
    delete_filled_export(rid)
    return True


def cleanup_expired_generation_history(retention_days: int | None = None) -> int:
    """删除创建时间早于「现在 − 保留天数」的历史记录，并删除 filled_exports 下对应成品 .pptx。"""
    from ppt_report import config as app_config
    from ppt_report.services.filled_export_cache import delete_filled_export

    if not _SessionLocal:
        return 0
    days = retention_days if retention_days is not None else app_config.GENERATION_HISTORY_RETENTION_DAYS
    days = max(1, int(days))
    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    with _SessionLocal() as session:
        ids = list(
            session.scalars(
                select(GenerationHistory.id).where(GenerationHistory.created_at < cutoff),
            ).all(),
        )
        if not ids:
            return 0
        session.execute(delete(GenerationHistory).where(GenerationHistory.created_at < cutoff))
        session.commit()
    for hid in ids:
        delete_filled_export(str(hid))
    return len(ids)


def get_overview_stats() -> dict[str, Any]:
    """首页概览：各模块数量（库未启用或查询失败时返回 0）。"""
    base: dict[str, Any] = {
        "db_enabled": False,
        "parsed_templates": 0,
        "generation_total": 0,
        "generation_this_month": 0,
        "student_profiles": 0,
        "chapter_templates": 0,
    }
    if not _SessionLocal:
        return base
    now = datetime.now(timezone.utc)
    month_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    try:
        with _SessionLocal() as session:
            n_parsed = session.scalar(select(func.count()).select_from(ParsedPresentation)) or 0
            n_gen = session.scalar(select(func.count()).select_from(GenerationHistory)) or 0
            n_gen_month = session.scalar(
                select(func.count())
                .select_from(GenerationHistory)
                .where(GenerationHistory.created_at >= month_start),
            ) or 0
            n_profiles = session.scalar(select(func.count()).select_from(StudentTermProfile)) or 0
            n_ch_tpl = session.scalar(select(func.count()).select_from(ChapterTemplate)) or 0
        return {
            "db_enabled": True,
            "parsed_templates": int(n_parsed),
            "generation_total": int(n_gen),
            "generation_this_month": int(n_gen_month),
            "student_profiles": int(n_profiles),
            "chapter_templates": int(n_ch_tpl),
        }
    except Exception:
        log.exception("读取概览统计失败")
        return {**base, "db_enabled": True}


def list_presentation_summaries(limit: int = 500) -> list[dict[str, object]]:
    """列表用摘要（不含 payload）；按创建时间倒序。"""
    if not _SessionLocal:
        return []
    lim = max(1, min(int(limit), 2000))
    with _SessionLocal() as session:
        rows = session.scalars(
            select(ParsedPresentation).order_by(ParsedPresentation.created_at.desc()).limit(lim),
        ).all()
    out: list[dict[str, object]] = []
    for r in rows:
        created = r.created_at
        pl = r.payload if isinstance(r.payload, dict) else {}
        out.append(
            {
                "task_id": r.task_id,
                "file_name": r.file_name or "",
                "slide_count": int(r.slide_count or 0),
                "created_at": created.isoformat() if created else None,
                "parse_impl": _parse_impl_label_from_payload(pl),
            },
        )
    return out


def list_student_records(query: str = "", limit: int = 500) -> list[dict[str, object]]:
    if not _SessionLocal:
        return []
    lim = max(1, min(int(limit), 2000))
    q = _norm_str(query)
    with _SessionLocal() as session:
        stmt = (
            select(StudentTermProfile, Student)
            .join(Student, Student.id == StudentTermProfile.student_id)
            .order_by(StudentTermProfile.updated_at.desc())
            .limit(lim)
        )
        if q:
            like = f"%{q}%"
            stmt = stmt.where(
                or_(
                    Student.student_name.ilike(like),
                    Student.student_code.ilike(like),
                    StudentTermProfile.school.ilike(like),
                    StudentTermProfile.major.ilike(like),
                    text("(student_term_profiles.extra_json->>'className') ILIKE :like"),
                    text("(student_term_profiles.extra_json->>'content') ILIKE :like"),
                ),
            ).params(like=like)
        rows = session.execute(stmt).all()
    return [_record_from_pair(stu, rec) for rec, stu in rows]


def get_student_record(record_id: str) -> dict[str, object] | None:
    rid = _norm_str(record_id, 64)
    if not _SessionLocal or not rid:
        return None
    with _SessionLocal() as session:
        pair = session.execute(
            select(StudentTermProfile, Student)
            .join(Student, Student.id == StudentTermProfile.student_id)
            .where(StudentTermProfile.id == rid),
        ).first()
        if not pair:
            return None
        rec, stu = pair
        return _record_from_pair(stu, rec)


def save_student_record(payload: dict[str, Any]) -> tuple[str, dict[str, object]]:
    if not _SessionLocal:
        raise RuntimeError("database disabled")
    data = payload if isinstance(payload, dict) else {}
    profile = data.get("profile") if isinstance(data.get("profile"), dict) else {}
    basic = profile.get("basic") if isinstance(profile.get("basic"), dict) else {}
    learning = profile.get("learning") if isinstance(profile.get("learning"), dict) else {}
    hours = profile.get("hours") if isinstance(profile.get("hours"), dict) else {}
    guidance = profile.get("guidance") if isinstance(profile.get("guidance"), dict) else {}

    content = _norm_str(data.get("content"))
    student_name = _norm_str(basic.get("studentName") or data.get("name"), 128)
    student_code = _norm_str(basic.get("studentId") or data.get("studentId"), 64)
    if not student_name and not student_code and not content:
        raise ValueError("请至少填写姓名、学号或数据内容中的一项。")

    record_id = _norm_str(data.get("id"), 64)
    with _SessionLocal() as session:
        existing = session.get(StudentTermProfile, record_id) if record_id else None
        student: Student | None = None
        if student_code:
            student = session.execute(
                select(Student).where(Student.student_code == student_code),
            ).scalar_one_or_none()
        if not student and existing:
            student = session.get(Student, existing.student_id)
        if not student:
            student = Student()
            student.id = str(uuid4())
            session.add(student)

        if not student_code:
            if student.student_code:
                student_code = student.student_code
            else:
                student_code = f"AUTO_{uuid4().hex[:10]}"

        student.student_code = student_code
        student.student_name = student_name or student.student_name or ""
        student.nickname_en = _pick_profile(profile, "basic", "nicknameEn", 128)
        student.service_start_date = _pick_profile(profile, "basic", "serviceStart", 32)
        student.planner_teacher = _pick_profile(profile, "basic", "plannerTeacher", 64)

        rec = existing or StudentTermProfile()
        rec.student_id = student.id
        rec.term_code = _pick_profile(profile, "basic", "currentTerm", 32)
        rec.grade_level = _pick_profile(profile, "basic", "gradeLevel", 64)
        rec.school = _pick_profile(profile, "basic", "school", 256)
        rec.major = _pick_profile(profile, "basic", "major", 256)
        rec.report_subtitle = _pick_profile(profile, "basic", "reportSubtitle", 256)

        rec.strong_subjects = _pick_profile(profile, "learning", "strongSubjects")
        rec.intl_scores = _pick_profile(profile, "learning", "intlScores", 128)
        rec.study_intent = _pick_profile(profile, "learning", "studyIntent", 128)
        rec.career_intent = _pick_profile(profile, "learning", "careerIntent", 128)
        rec.interest_subjects = _pick_profile(profile, "learning", "interestSubjects")
        rec.long_term_plan = _pick_profile(profile, "learning", "longTermPlan")
        rec.learning_style = _pick_profile(profile, "learning", "learningStyle")
        rec.weak_areas = _pick_profile(profile, "learning", "weakAreas")

        rec.total_hours = _parse_optional_int(hours.get("totalHours"))
        rec.used_hours = _parse_optional_int(hours.get("usedHours"))
        rec.remaining_hours = _parse_optional_int(hours.get("remainingHours"))
        rec.tutor_subjects = _pick_profile(profile, "hours", "tutorSubjects")
        rec.preview_subjects = _pick_profile(profile, "hours", "previewSubjects")
        rec.skill_direction = _pick_profile(profile, "hours", "skillDirection")
        rec.skill_description = _pick_profile(profile, "hours", "skillDescription")

        rec.term_summary = _pick_profile(profile, "guidance", "termSummary")
        rec.course_feedback = _pick_profile(profile, "guidance", "courseFeedback")
        rec.short_term_advice = _pick_profile(profile, "guidance", "shortTermAdvice")
        rec.long_term_development = _pick_profile(profile, "guidance", "longTermDevelopment")

        # JSONB 字段需重新赋值新 dict 才能稳定触发 SQLAlchemy 脏检查
        extra = dict(rec.extra_json) if isinstance(rec.extra_json, dict) else {}
        extra.update(
            {
                "className": _pick_profile(profile, "basic", "className", 128),
                "email": _pick_profile(profile, "basic", "email", 256),
                "phone": _pick_profile(profile, "basic", "phone", 64),
                "remark": _pick_profile(profile, "basic", "remark", 512),
                "content": content,
                "profile": {
                    "basic": basic if isinstance(basic, dict) else {},
                    "learning": learning if isinstance(learning, dict) else {},
                    "hours": hours if isinstance(hours, dict) else {},
                    "guidance": guidance if isinstance(guidance, dict) else {},
                },
            },
        )
        rec.extra_json = extra

        if not existing:
            session.add(rec)
        session.commit()
        session.refresh(rec)
        session.refresh(student)
        return rec.id, _record_from_pair(student, rec)


def delete_student_record(record_id: str) -> bool:
    rid = _norm_str(record_id, 64)
    if not _SessionLocal or not rid:
        return False
    with _SessionLocal() as session:
        rec = session.get(StudentTermProfile, rid)
        if not rec:
            return False
        session.delete(rec)
        session.commit()
        return True


def _normalize_template_chapters(raw_chapters: object) -> list[dict[str, object]]:
    rows = raw_chapters if isinstance(raw_chapters, list) else []
    out: list[dict[str, object]] = []
    for i, ch in enumerate(rows):
        if not isinstance(ch, dict):
            continue
        title = _norm_str(ch.get("title"), 128) or "未命名章节"
        hint = _norm_str(ch.get("hint"))
        sort_raw = ch.get("sort")
        sort = int(sort_raw) if isinstance(sort_raw, int) else i
        out.append(
            {
                "id": _norm_str(ch.get("id"), 64),
                "title": title,
                "hint": hint,
                "sort": sort,
            },
        )
    out.sort(key=lambda x: int(x.get("sort") or 0))
    for idx, row in enumerate(out):
        row["sort"] = idx
    return out


def _template_record(tpl: ChapterTemplate, chapters: list[ChapterTemplateChapter]) -> dict[str, object]:
    ch_rows = [
        {
            "id": c.id,
            "title": _norm_str(c.title),
            "hint": _norm_str(c.hint),
            "sort": int(c.sort_order or 0),
        }
        for c in sorted(chapters, key=lambda r: int(r.sort_order or 0))
    ]
    return {
        "id": tpl.id,
        "name": _norm_str(tpl.name),
        "description": _norm_str(tpl.description),
        "chapters": ch_rows,
        "chapterCount": len(ch_rows),
        "createdAt": tpl.created_at.isoformat() if tpl.created_at else "",
        "updatedAt": tpl.updated_at.isoformat() if tpl.updated_at else "",
    }


def list_chapter_templates(query: str = "", limit: int = 500) -> list[dict[str, object]]:
    if not _SessionLocal:
        return []
    q = _norm_str(query)
    lim = max(1, min(int(limit), 2000))
    with _SessionLocal() as session:
        stmt = select(ChapterTemplate).order_by(ChapterTemplate.updated_at.desc()).limit(lim)
        if q:
            like = f"%{q}%"
            stmt = stmt.where(
                or_(
                    ChapterTemplate.name.ilike(like),
                    ChapterTemplate.description.ilike(like),
                    ChapterTemplate.template_code.ilike(like),
                ),
            )
        rows = session.scalars(stmt).all()
        out: list[dict[str, object]] = []
        for tpl in rows:
            ch_count = session.scalar(
                select(func.count()).select_from(ChapterTemplateChapter).where(
                    ChapterTemplateChapter.template_id == tpl.id,
                ),
            ) or 0
            raw_desc = _norm_str(tpl.description)
            desc_preview = raw_desc if len(raw_desc) <= 160 else raw_desc[:157] + "…"
            out.append(
                {
                    "id": tpl.id,
                    "name": _norm_str(tpl.name),
                    "description": desc_preview,
                    "chapterCount": int(ch_count),
                    "updatedAt": tpl.updated_at.isoformat() if tpl.updated_at else "",
                },
            )
        return out


def get_chapter_template(template_id: str) -> dict[str, object] | None:
    tid = _norm_str(template_id, 64)
    if not _SessionLocal or not tid:
        return None
    with _SessionLocal() as session:
        tpl = session.get(ChapterTemplate, tid)
        if not tpl:
            return None
        chs = session.scalars(
            select(ChapterTemplateChapter)
            .where(ChapterTemplateChapter.template_id == tid)
            .order_by(ChapterTemplateChapter.sort_order.asc()),
        ).all()
        return _template_record(tpl, list(chs))


def save_chapter_template(payload: dict[str, Any]) -> tuple[str, dict[str, object]]:
    if not _SessionLocal:
        raise RuntimeError("database disabled")
    data = payload if isinstance(payload, dict) else {}
    template_id = _norm_str(data.get("id"), 64)
    name = _norm_str(data.get("name"), 128)
    description = _norm_str(data.get("description"))
    chapters = _normalize_template_chapters(data.get("chapters"))
    if not name:
        raise ValueError("请填写模板名称。")
    if not chapters:
        raise ValueError("至少保留一个章节。")

    with _SessionLocal() as session:
        tpl = session.get(ChapterTemplate, template_id) if template_id else None
        if template_id and not tpl:
            raise KeyError("模板不存在。")
        if not tpl:
            tpl = ChapterTemplate(
                id=str(uuid4()),
                template_code=f"TPL_{uuid4().hex[:12].upper()}",
                is_active=True,
            )
            session.add(tpl)

        tpl.name = name
        tpl.description = description
        session.flush()

        old_rows = session.scalars(
            select(ChapterTemplateChapter).where(ChapterTemplateChapter.template_id == tpl.id),
        ).all()
        old_map = {r.id: r for r in old_rows}
        keep_ids: set[str] = set()
        for row in chapters:
            rid = _norm_str(row.get("id"), 64)
            rec = old_map.get(rid) if rid else None
            if not rec:
                rec = ChapterTemplateChapter(id=str(uuid4()), template_id=tpl.id)
                session.add(rec)
            rec.title = _norm_str(row.get("title"), 128) or "未命名章节"
            rec.hint = _norm_str(row.get("hint"))
            rec.sort_order = int(row.get("sort") or 0)
            keep_ids.add(rec.id)

        for old in old_rows:
            if old.id not in keep_ids:
                session.delete(old)

        session.commit()
        session.refresh(tpl)
        chs = session.scalars(
            select(ChapterTemplateChapter)
            .where(ChapterTemplateChapter.template_id == tpl.id)
            .order_by(ChapterTemplateChapter.sort_order.asc()),
        ).all()
        return tpl.id, _template_record(tpl, list(chs))


def delete_chapter_template(template_id: str) -> bool:
    tid = _norm_str(template_id, 64)
    if not _SessionLocal or not tid:
        return False
    with _SessionLocal() as session:
        tpl = session.get(ChapterTemplate, tid)
        if not tpl:
            return False
        session.delete(tpl)
        session.commit()
        return True

