"""
PostgreSQL 持久化：解析结果全量 JSON（JSONB）。
环境变量 DATABASE_URL，默认连本地库 ppt_report_platform。
"""
from __future__ import annotations

import logging
import re
from datetime import date, datetime, timedelta, timezone
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
    advisor_teacher: Mapped[str] = mapped_column(String(64), nullable=False, default="")
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
    grade_intake: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    school: Mapped[str] = mapped_column(String(256), nullable=False, default="")
    major: Mapped[str] = mapped_column(String(256), nullable=False, default="")
    report_subtitle: Mapped[str] = mapped_column(String(256), nullable=False, default="")
    service_product: Mapped[str] = mapped_column(String(256), nullable=False, default="")

    # 学习画像（字段需求：选课/规划；列名与 profile.learning 键均为 snake_case）
    strength_subjects: Mapped[str] = mapped_column(Text, nullable=False, default="")
    scores: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    learning_good: Mapped[str] = mapped_column(Text, nullable=False, default="")
    learning_weak: Mapped[str] = mapped_column(Text, nullable=False, default="")
    interests: Mapped[str] = mapped_column(Text, nullable=False, default="")
    study_goal: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    career_goal: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    long_goal: Mapped[str] = mapped_column(Text, nullable=False, default="")
    degree: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    duration: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    credits: Mapped[str] = mapped_column(Text, nullable=False, default="")
    course_rule: Mapped[str] = mapped_column(Text, nullable=False, default="")
    gpa_rule: Mapped[str] = mapped_column(Text, nullable=False, default="")
    selection_rule: Mapped[str] = mapped_column(Text, nullable=False, default="")
    recommended_courses: Mapped[str] = mapped_column(Text, nullable=False, default="")
    course_notes: Mapped[str] = mapped_column(Text, nullable=False, default="")
    term_plan: Mapped[str] = mapped_column(Text, nullable=False, default="")
    future_plan: Mapped[str] = mapped_column(Text, nullable=False, default="")

    # 课时数据
    total_hours: Mapped[int | None] = mapped_column(Integer, nullable=True)
    used_hours: Mapped[int | None] = mapped_column(Integer, nullable=True)
    remaining_hours: Mapped[int | None] = mapped_column(Integer, nullable=True)
    prep_courses: Mapped[str] = mapped_column(Text, nullable=False, default="")
    tutoring_courses: Mapped[str] = mapped_column(Text, nullable=False, default="")
    skill_direction: Mapped[str] = mapped_column(Text, nullable=False, default="")
    skill_description: Mapped[str] = mapped_column(Text, nullable=False, default="")

    # 学期总结/结单（字段需求；profile.term_summary；DB 列 snake_case，「备注」列名 summary_remarks）
    student_summary: Mapped[str] = mapped_column(Text, nullable=False, default="")
    school_ddl: Mapped[str] = mapped_column(Text, nullable=False, default="")
    first_class_time: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    first_class_note: Mapped[str] = mapped_column(Text, nullable=False, default="")
    summer_work: Mapped[str] = mapped_column(Text, nullable=False, default="")
    term_work: Mapped[str] = mapped_column(Text, nullable=False, default="")
    recorded_courses: Mapped[str] = mapped_column(Text, nullable=False, default="")
    grades: Mapped[str] = mapped_column(Text, nullable=False, default="")
    gpa: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    target_gpa: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    final_score: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    services: Mapped[str] = mapped_column(Text, nullable=False, default="")
    service_count: Mapped[str] = mapped_column(String(32), nullable=False, default="")
    class_count: Mapped[str] = mapped_column(String(32), nullable=False, default="")
    total_duration: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    avg_duration: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    communication: Mapped[str] = mapped_column(Text, nullable=False, default="")
    next_goal: Mapped[str] = mapped_column(Text, nullable=False, default="")
    risk_courses: Mapped[str] = mapped_column(Text, nullable=False, default="")
    suggestions: Mapped[str] = mapped_column(Text, nullable=False, default="")
    summary_remarks: Mapped[str] = mapped_column(Text, nullable=False, default="")

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


WORD_TABLE_FILL_TEMPLATE_CODE = "word_table_fill"


def _norm_str(v: object, max_len: int | None = None) -> str:
    s = "" if v is None else str(v).strip()
    return s[:max_len] if max_len and max_len > 0 else s


def _pick_profile(profile: dict[str, Any], dim: str, key: str, max_len: int | None = None) -> str:
    src = profile.get(dim) if isinstance(profile, dict) else None
    if not isinstance(src, dict):
        return ""
    return _norm_str(src.get(key), max_len)


def _pick_learning_field(
    profile: dict[str, Any],
    new_key: str,
    max_len: int | None,
    *legacy_keys: str,
) -> str:
    """读取 learning 维度；优先字段需求 snake_case，兼容旧 camelCase。"""
    v = _pick_profile(profile, "learning", new_key, max_len)
    if v:
        return v
    for lk in legacy_keys:
        v = _pick_profile(profile, "learning", lk, max_len)
        if v:
            return v
    return ""


def _pick_hours_field(
    profile: dict[str, Any],
    new_key: str,
    max_len: int | None,
    *legacy_keys: str,
) -> str:
    v = _pick_profile(profile, "hours", new_key, max_len)
    if v:
        return v
    for lk in legacy_keys:
        v = _pick_profile(profile, "hours", lk, max_len)
        if v:
            return v
    return ""


def _pick_term_summary_field(profile: dict[str, Any], key: str, max_len: int | None = None) -> str:
    return _pick_profile(profile, "term_summary", key, max_len)


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
    ts_block = merged.get("term_summary") if isinstance(merged.get("term_summary"), dict) else {}
    basic.update(
        {
            "studentName": _norm_str(student.student_name),
            "studentId": _norm_str(student.student_code),
            "nicknameEn": _norm_str(student.nickname_en),
            "serviceStart": _norm_str(student.service_start_date),
            "plannerTeacher": _norm_str(student.planner_teacher),
            "advisorTeacher": _norm_str(student.advisor_teacher),
            "school": _norm_str(row.school),
            "major": _norm_str(row.major),
            "gradeLevel": _norm_str(row.grade_level),
            "gradeIntake": _norm_str(row.grade_intake),
            "currentTerm": _norm_str(row.term_code),
            "product": _norm_str(row.service_product),
            "reportSubtitle": _norm_str(row.report_subtitle),
            "className": _norm_str(extra.get("className")),
            "email": _norm_str(extra.get("email")),
            "phone": _norm_str(extra.get("phone")),
            "remark": _norm_str(extra.get("remark")),
        },
    )
    learning.update(
        {
            "strength_subjects": _norm_str(row.strength_subjects),
            "scores": _norm_str(row.scores),
            "learning_good": _norm_str(row.learning_good),
            "learning_weak": _norm_str(row.learning_weak),
            "interests": _norm_str(row.interests),
            "study_goal": _norm_str(row.study_goal),
            "career_goal": _norm_str(row.career_goal),
            "long_goal": _norm_str(row.long_goal),
            "degree": _norm_str(row.degree),
            "duration": _norm_str(row.duration),
            "credits": _norm_str(row.credits),
            "course_rule": _norm_str(row.course_rule),
            "gpa_rule": _norm_str(row.gpa_rule),
            "selection_rule": _norm_str(row.selection_rule),
            "recommended_courses": _norm_str(row.recommended_courses),
            "course_notes": _norm_str(row.course_notes),
            "term_plan": _norm_str(row.term_plan),
            "future_plan": _norm_str(row.future_plan),
        },
    )
    hours.update(
        {
            "totalHours": "" if row.total_hours is None else str(row.total_hours),
            "usedHours": "" if row.used_hours is None else str(row.used_hours),
            "remainingHours": "" if row.remaining_hours is None else str(row.remaining_hours),
            "prep_courses": _norm_str(row.prep_courses),
            "tutoring_courses": _norm_str(row.tutoring_courses),
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
    ts_block.update(
        {
            "student_summary": _norm_str(row.student_summary),
            "school_ddl": _norm_str(row.school_ddl),
            "first_class_time": _norm_str(row.first_class_time),
            "first_class_note": _norm_str(row.first_class_note),
            "summer_work": _norm_str(row.summer_work),
            "term_work": _norm_str(row.term_work),
            "recorded_courses": _norm_str(row.recorded_courses),
            "total_hours": "" if row.total_hours is None else str(row.total_hours),
            "used_hours": "" if row.used_hours is None else str(row.used_hours),
            "left_hours": "" if row.remaining_hours is None else str(row.remaining_hours),
            "grades": _norm_str(row.grades),
            "gpa": _norm_str(row.gpa),
            "target_gpa": _norm_str(row.target_gpa),
            "final_score": _norm_str(row.final_score),
            "services": _norm_str(row.services),
            "service_count": _norm_str(row.service_count),
            "class_count": _norm_str(row.class_count),
            "total_duration": _norm_str(row.total_duration),
            "avg_duration": _norm_str(row.avg_duration),
            "communication": _norm_str(row.communication),
            "next_goal": _norm_str(row.next_goal),
            "risk_courses": _norm_str(row.risk_courses),
            "suggestions": _norm_str(row.suggestions),
            "remarks": _norm_str(row.summary_remarks),
        },
    )
    return {
        "basic": basic,
        "learning": learning,
        "hours": hours,
        "guidance": guidance,
        "term_summary": ts_block,
    }


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
        with _engine.begin() as conn:
            conn.execute(text("SELECT 1"))
            _ensure_student_field_columns(conn)
            _migrate_student_term_profile_rename_and_add(conn)
            _migrate_term_summary_closing_columns(conn)
        try:
            with _SessionLocal() as s:
                _ensure_word_table_fill_template(s)
        except Exception:
            log.exception("Word 表格回填报告类型种子失败（可在「报告类型管理」中手动新建并勾选 Word 特殊类型）。")
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


def _ensure_student_field_columns(conn) -> None:
    """已有库在模型增列后需 ALTER；create_all 不会修改已存在表结构。"""
    for stmt in (
        "ALTER TABLE students ADD COLUMN IF NOT EXISTS advisor_teacher VARCHAR(64) NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS grade_intake VARCHAR(128) NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS service_product VARCHAR(256) NOT NULL DEFAULT ''",
    ):
        conn.execute(text(stmt))


def _pg_table_columns(conn, table: str) -> set[str]:
    r = conn.execute(
        text(
            "SELECT column_name FROM information_schema.columns "
            "WHERE table_schema = 'public' AND table_name = :t",
        ),
        {"t": table},
    )
    return {row[0] for row in r.fetchall()}


def _migrate_student_term_profile_rename_and_add(conn) -> None:
    """旧列名重命名为字段需求名，并补充选课/规划新增列。"""
    cols = _pg_table_columns(conn, "student_term_profiles")
    if not cols:
        return
    renames = [
        ("strong_subjects", "strength_subjects"),
        ("intl_scores", "scores"),
        ("study_intent", "study_goal"),
        ("career_intent", "career_goal"),
        ("interest_subjects", "interests"),
        ("long_term_plan", "long_goal"),
        ("learning_style", "learning_good"),
        ("weak_areas", "learning_weak"),
        ("tutor_subjects", "tutoring_courses"),
        ("preview_subjects", "prep_courses"),
    ]
    for old, new in renames:
        if old in cols and new not in cols:
            conn.execute(text(f'ALTER TABLE student_term_profiles RENAME COLUMN "{old}" TO "{new}"'))
            cols.discard(old)
            cols.add(new)
    cols = _pg_table_columns(conn, "student_term_profiles")
    for stmt in (
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS degree VARCHAR(128) NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS duration VARCHAR(64) NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS credits TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS course_rule TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS gpa_rule TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS selection_rule TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS recommended_courses TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS course_notes TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS term_plan TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS future_plan TEXT NOT NULL DEFAULT ''",
    ):
        conn.execute(text(stmt))


def _migrate_term_summary_closing_columns(conn) -> None:
    """学期总结/结单模板字段（字段需求）。"""
    for stmt in (
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS student_summary TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS school_ddl TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS first_class_time VARCHAR(64) NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS first_class_note TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS summer_work TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS term_work TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS recorded_courses TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS grades TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS gpa VARCHAR(64) NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS target_gpa VARCHAR(64) NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS final_score VARCHAR(128) NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS services TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS service_count VARCHAR(32) NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS class_count VARCHAR(32) NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS total_duration VARCHAR(64) NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS avg_duration VARCHAR(64) NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS communication TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS next_goal TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS risk_courses TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS suggestions TEXT NOT NULL DEFAULT ''",
        "ALTER TABLE student_term_profiles ADD COLUMN IF NOT EXISTS summary_remarks TEXT NOT NULL DEFAULT ''",
    ):
        conn.execute(text(stmt))


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
    tk = str(payload.get("template_kind") or "").strip()
    if tk == "word_stored":
        return "Word 文档（仅存储）"
    if tk == "word_parsed":
        return "Word 文档（已解析）"
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


def list_generation_history_summaries(query: str = "", limit: int = 200) -> list[dict[str, object]]:
    """生成历史列表摘要（不含 result JSON），支持关键词检索，按时间倒序。"""
    if not _SessionLocal:
        return []
    lim = max(1, min(int(limit), 2000))
    q = _norm_str(query)
    with _SessionLocal() as session:
        stmt = select(GenerationHistory).order_by(GenerationHistory.created_at.desc()).limit(lim)
        if q:
            like = f"%{q}%"
            stmt = stmt.where(
                or_(
                    GenerationHistory.topic.ilike(like),
                    GenerationHistory.task_id.ilike(like),
                ),
            )
        rows = session.scalars(stmt).all()
    out: list[dict[str, object]] = []
    for r in rows:
        created = r.created_at
        res = r.result if isinstance(r.result, dict) else {}
        output_kind = "docx" if str(res.get("output_kind") or "").strip() == "docx" else "pptx"
        out.append(
            {
                "id": r.id,
                "taskId": r.task_id or "",
                "topic": r.topic or "",
                "createdAt": created.isoformat() if created else "",
                "slideCount": int(r.slide_count or 0),
                "outputKind": output_kind,
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
            "outputKind": "docx" if str(res.get("output_kind") or "").strip() == "docx" else "pptx",
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


def _empty_gen_by_day_labels() -> list[dict[str, Any]]:
    now = datetime.now(timezone.utc)
    today = now.date()
    out: list[dict[str, Any]] = []
    for i in range(7):
        d = today - timedelta(days=6 - i)
        out.append({"label": f"{d.month}/{d.day}", "count": 0})
    return out


def get_overview_chart_data() -> dict[str, Any]:
    """首页图表：近 7 日生成条数、全部历史上 PPT/Word 产物数量。"""
    empty: dict[str, Any] = {
        "gen_by_day": _empty_gen_by_day_labels(),
        "output_pptx": 0,
        "output_docx": 0,
        "bar_max": 1,
    }
    if not _SessionLocal:
        return empty
    now = datetime.now(timezone.utc)
    today = now.date()
    start_dt = datetime.combine(today - timedelta(days=6), datetime.min.time(), tzinfo=timezone.utc)
    try:
        with _SessionLocal() as session:
            created_rows = session.scalars(
                select(GenerationHistory.created_at).where(GenerationHistory.created_at >= start_dt),
            ).all()
        by_day: dict[date, int] = {}
        for ca in created_rows:
            if ca is None:
                continue
            d = ca.date()
            by_day[d] = by_day.get(d, 0) + 1
        gen_by_day: list[dict[str, Any]] = []
        for i in range(7):
            d = today - timedelta(days=6 - i)
            gen_by_day.append(
                {
                    "label": f"{d.month}/{d.day}",
                    "count": int(by_day.get(d, 0)),
                },
            )
        max_daily = max((x["count"] for x in gen_by_day), default=0)
        bar_max = max(max_daily, 1)
        with _SessionLocal() as session:
            n_docx = int(
                session.scalar(
                    select(func.count())
                    .select_from(GenerationHistory)
                    .where(GenerationHistory.result["output_kind"].astext == "docx"),
                )
                or 0,
            )
            n_total = int(session.scalar(select(func.count()).select_from(GenerationHistory)) or 0)
        n_pptx = max(0, n_total - n_docx)
        return {
            "gen_by_day": gen_by_day,
            "output_pptx": n_pptx,
            "output_docx": n_docx,
            "bar_max": bar_max,
        }
    except Exception:
        log.exception("读取概览图表数据失败")
        return empty


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
        tk = str(pl.get("template_kind") or "").strip() or "ppt_parsed"
        out.append(
            {
                "task_id": r.task_id,
                "file_name": r.file_name or "",
                "slide_count": int(r.slide_count or 0),
                "created_at": created.isoformat() if created else None,
                "parse_impl": _parse_impl_label_from_payload(pl),
                "template_kind": tk,
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
    term_summary = profile.get("term_summary") if isinstance(profile.get("term_summary"), dict) else {}

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
        student.advisor_teacher = _pick_profile(profile, "basic", "advisorTeacher", 64)

        term_code_val = _pick_profile(profile, "basic", "currentTerm", 32)
        rec: StudentTermProfile
        if existing:
            rec = existing
        else:
            dup = session.execute(
                select(StudentTermProfile).where(
                    StudentTermProfile.student_id == student.id,
                    StudentTermProfile.term_code == term_code_val,
                ),
            ).scalar_one_or_none()
            if dup is not None:
                # 同一学生、同一学期已有一条快照：前端「新建」未带 id 时仍应更新该行，避免违反 uq_student_term_profiles_student_term
                rec = dup
                existing = dup
            else:
                rec = StudentTermProfile()
        rec.student_id = student.id
        rec.term_code = term_code_val
        rec.grade_level = _pick_profile(profile, "basic", "gradeLevel", 64)
        rec.grade_intake = _pick_profile(profile, "basic", "gradeIntake", 128)
        rec.school = _pick_profile(profile, "basic", "school", 256)
        rec.major = _pick_profile(profile, "basic", "major", 256)
        rec.report_subtitle = _pick_profile(profile, "basic", "reportSubtitle", 256)
        rec.service_product = _pick_profile(profile, "basic", "product", 256)

        rec.strength_subjects = _pick_learning_field(
            profile, "strength_subjects", None, "strongSubjects",
        )
        rec.scores = _pick_learning_field(profile, "scores", 128, "intlScores")
        rec.learning_good = _pick_learning_field(profile, "learning_good", None, "learningStyle")
        rec.learning_weak = _pick_learning_field(profile, "learning_weak", None, "weakAreas")
        rec.interests = _pick_learning_field(profile, "interests", None, "interestSubjects")
        rec.study_goal = _pick_learning_field(profile, "study_goal", 128, "studyIntent")
        rec.career_goal = _pick_learning_field(profile, "career_goal", 128, "careerIntent")
        rec.long_goal = _pick_learning_field(profile, "long_goal", None, "longTermPlan")
        rec.degree = _pick_learning_field(profile, "degree", 128)
        rec.duration = _pick_learning_field(profile, "duration", 64)
        rec.credits = _pick_learning_field(profile, "credits", None)
        rec.course_rule = _pick_learning_field(profile, "course_rule", None)
        rec.gpa_rule = _pick_learning_field(profile, "gpa_rule", None)
        rec.selection_rule = _pick_learning_field(profile, "selection_rule", None)
        rec.recommended_courses = _pick_learning_field(profile, "recommended_courses", None)
        rec.course_notes = _pick_learning_field(profile, "course_notes", None)
        rec.term_plan = _pick_learning_field(profile, "term_plan", None)
        rec.future_plan = _pick_learning_field(profile, "future_plan", None)

        def _hours_int(camel: str, snake: str) -> int | None:
            v = _parse_optional_int(hours.get(camel))
            if v is not None:
                return v
            return _parse_optional_int(term_summary.get(snake))

        rec.total_hours = _hours_int("totalHours", "total_hours")
        rec.used_hours = _hours_int("usedHours", "used_hours")
        rec.remaining_hours = _hours_int("remainingHours", "left_hours")
        rec.prep_courses = _pick_hours_field(profile, "prep_courses", None, "previewSubjects")
        rec.tutoring_courses = _pick_hours_field(profile, "tutoring_courses", None, "tutorSubjects")
        rec.skill_direction = _pick_profile(profile, "hours", "skillDirection")
        rec.skill_description = _pick_profile(profile, "hours", "skillDescription")

        rec.student_summary = _pick_term_summary_field(profile, "student_summary")
        rec.school_ddl = _pick_term_summary_field(profile, "school_ddl")
        rec.first_class_time = _pick_term_summary_field(profile, "first_class_time", 64)
        rec.first_class_note = _pick_term_summary_field(profile, "first_class_note")
        rec.summer_work = _pick_term_summary_field(profile, "summer_work")
        rec.term_work = _pick_term_summary_field(profile, "term_work")
        rec.recorded_courses = _pick_term_summary_field(profile, "recorded_courses")
        rec.grades = _pick_term_summary_field(profile, "grades")
        rec.gpa = _pick_term_summary_field(profile, "gpa", 64)
        rec.target_gpa = _pick_term_summary_field(profile, "target_gpa", 64)
        rec.final_score = _pick_term_summary_field(profile, "final_score", 128)
        rec.services = _pick_term_summary_field(profile, "services")
        rec.service_count = _pick_term_summary_field(profile, "service_count", 32)
        rec.class_count = _pick_term_summary_field(profile, "class_count", 32)
        rec.total_duration = _pick_term_summary_field(profile, "total_duration", 64)
        rec.avg_duration = _pick_term_summary_field(profile, "avg_duration", 64)
        rec.communication = _pick_term_summary_field(profile, "communication")
        rec.next_goal = _pick_term_summary_field(profile, "next_goal")
        rec.risk_courses = _pick_term_summary_field(profile, "risk_courses")
        rec.suggestions = _pick_term_summary_field(profile, "suggestions")
        rec.summary_remarks = _pick_term_summary_field(profile, "remarks")

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
                    "term_summary": term_summary if isinstance(term_summary, dict) else {},
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
        "templateCode": _norm_str(tpl.template_code, 64),
        "name": _norm_str(tpl.name),
        "description": _norm_str(tpl.description),
        "chapters": ch_rows,
        "chapterCount": len(ch_rows),
        "createdAt": tpl.created_at.isoformat() if tpl.created_at else "",
        "updatedAt": tpl.updated_at.isoformat() if tpl.updated_at else "",
    }


def _ensure_word_table_fill_template(session) -> None:
    existing = session.scalar(
        select(ChapterTemplate).where(ChapterTemplate.template_code == WORD_TABLE_FILL_TEMPLATE_CODE),
    )
    if existing:
        return
    tpl = ChapterTemplate(
        id=str(uuid4()),
        template_code=WORD_TABLE_FILL_TEMPLATE_CODE,
        name="Word 表格回填",
        description="特殊报告类型：仅用于 Word 模板表格回填，不参与 PPT 章节解析。",
        is_active=True,
    )
    session.add(tpl)
    session.flush()
    session.add(
        ChapterTemplateChapter(
            id=str(uuid4()),
            template_id=tpl.id,
            title="Word 表格",
            hint="仅用于触发 Word 生成模式。",
            sort_order=1,
        ),
    )
    session.commit()


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
                    "templateCode": _norm_str(tpl.template_code, 64),
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
    word_fill_requested = bool(data.get("wordTableFill") or data.get("word_table_fill"))
    if not name:
        raise ValueError("请填写模板名称。")
    if not chapters:
        raise ValueError("至少保留一个章节。")

    with _SessionLocal() as session:
        tpl = session.get(ChapterTemplate, template_id) if template_id else None
        if template_id and not tpl:
            raise KeyError("模板不存在。")
        if tpl and word_fill_requested and tpl.template_code != WORD_TABLE_FILL_TEMPLATE_CODE:
            raise ValueError("无法将已有报告类型切换为 Word 特殊类型；请新建报告类型并勾选「Word 表格回填」。")
        if not tpl:
            if word_fill_requested:
                dup = session.scalar(
                    select(ChapterTemplate).where(
                        ChapterTemplate.template_code == WORD_TABLE_FILL_TEMPLATE_CODE,
                    ),
                )
                if dup:
                    raise ValueError(
                        "已存在编码为 word_table_fill 的 Word 表格回填类型，请在列表中编辑该条，勿重复创建。",
                    )
                tpl = ChapterTemplate(
                    id=str(uuid4()),
                    template_code=WORD_TABLE_FILL_TEMPLATE_CODE,
                    is_active=True,
                )
            else:
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

