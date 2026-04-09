"""章节模板 + 学生数据解析：按顺序对齐模板章节名，由大模型结合成品 PPT 各章正文分配学生字段。"""
from __future__ import annotations

import json
import logging
import os
from typing import Any

import requests

from ppt_report.models import db as db_mod
from ppt_report.services.llm_json import extract_json_from_text
from ppt_report.services.page_types import compute_chapter_selection_groups
from ppt_report.services.presentation_cache import get_parsed_from_cache

log = logging.getLogger(__name__)

_LABEL_HINTS: dict[str, str] = {
    "name": "姓名",
    "studentId": "学号",
    "className": "班级",
    "email": "邮箱",
    "phone": "手机",
    "remark": "备注",
    "content": "数据内容",
    "basic.studentName": "姓名",
    "basic.studentId": "学号",
    "basic.nicknameEn": "英文名",
    "basic.school": "学校",
    "basic.major": "专业",
    "basic.gradeLevel": "年级",
    "basic.currentTerm": "学期",
    "basic.className": "班级",
    "basic.plannerTeacher": "规划老师",
    "learning.strongSubjects": "优势学科",
    "learning.studyIntent": "升学意向",
    "learning.careerIntent": "职业意向",
    "learning.longTermPlan": "长期规划",
    "learning.learningStyle": "学习风格",
    "hours.totalHours": "总课时",
    "hours.usedHours": "已用课时",
    "guidance.termSummary": "学期总结",
    "guidance.courseFeedback": "课程反馈",
}


def _label_for_key(key: str) -> str:
    return _LABEL_HINTS.get(key) or key


def _push_field(out: list[dict[str, str]], key: str, label: str, v: object) -> None:
    if v is None:
        return
    t = str(v).strip()
    if not t:
        return
    out.append({"key": key, "label": label or _label_for_key(key), "value": t[:800]})


def flatten_student_record(item: dict[str, Any]) -> list[dict[str, str]]:
    """与前端 generate-chapter-ref.js 的扁平化规则对齐。"""
    out: list[dict[str, str]] = []
    if not isinstance(item, dict):
        return out
    _push_field(out, "name", "姓名", item.get("name"))
    _push_field(out, "studentId", "学号", item.get("studentId"))
    _push_field(out, "className", "班级", item.get("className"))
    _push_field(out, "email", "邮箱", item.get("email"))
    _push_field(out, "phone", "手机", item.get("phone"))
    _push_field(out, "remark", "备注", item.get("remark"))
    _push_field(out, "content", "数据内容", item.get("content"))
    prof = item.get("profile")
    if not isinstance(prof, dict):
        return out
    for dim in ("basic", "learning", "hours", "guidance"):
        block = prof.get(dim)
        if not isinstance(block, dict):
            continue
        for k, v in block.items():
            if v is None or v == "":
                continue
            if isinstance(v, (dict, list)):
                continue
            key = f"{dim}.{k}"
            _push_field(out, key, _label_for_key(key), v)
    return out


def ppt_chapter_slot_rows(parsed: dict[str, Any]) -> list[dict[str, Any]]:
    """仅 kind=chapter 的块，顺序与生成页 tab 一致。"""
    groups = compute_chapter_selection_groups(parsed)
    return [g for g in groups if str(g.get("kind") or "") == "chapter"]


def ppt_reference_slot_rows(parsed: dict[str, Any]) -> list[dict[str, Any]]:
    """首页（若存在，仅取第一个 cover）+ 各「章」块，顺序与生成页 Tab 一致，用于解析与 chapter_ref_json。"""
    groups = compute_chapter_selection_groups(parsed)
    rows: list[dict[str, Any]] = []
    seen_cover = False
    for g in groups:
        k = str(g.get("kind") or "")
        if k == "cover" and not seen_cover:
            rows.append(g)
            seen_cover = True
        elif k == "chapter":
            rows.append(g)
    return rows


def _slides_by_index(parsed: dict[str, Any]) -> dict[int, dict[str, Any]]:
    out: dict[int, dict[str, Any]] = {}
    for s in parsed.get("slides") or []:
        if not isinstance(s, dict):
            continue
        try:
            idx = int(s.get("slide_index"))
        except (TypeError, ValueError):
            continue
        out[idx] = s
    return out


def _norm_slide_indices(slide_list: object) -> list[int]:
    seen: set[int] = set()
    ordered: list[int] = []
    if not isinstance(slide_list, list):
        return ordered
    for x in slide_list:
        try:
            n = int(x)
        except (TypeError, ValueError):
            continue
        if n not in seen:
            seen.add(n)
            ordered.append(n)
    return ordered


_SKIP_TYPES_FOR_EXCERPT = frozenset({"image", "chart", "line_or_arrow", "group"})


def build_chapter_ppt_report_excerpt(
    parsed: dict[str, Any],
    slide_indices: list[int] | object,
    *,
    max_total_chars: int | None = None,
    max_per_component: int | None = None,
) -> str:
    """
    从解析结果中抽取「该章涉及页」的正文与组件类型线索，供大模型对照成品报告划分学生字段。
    slide_indices 为 1-based 幻灯片页码，顺序与章节块内 pages 一致。
    """
    cap_total = max_total_chars
    if cap_total is None:
        cap_total = max(2000, int(os.getenv("CHAPTER_RESOLVE_MAX_PPT_CHARS_PER_SLOT", "8000")))
    cap_comp = max_per_component
    if cap_comp is None:
        cap_comp = max(200, int(os.getenv("CHAPTER_RESOLVE_MAX_COMPONENT_CHARS", "600")))

    indices = _norm_slide_indices(slide_indices)
    if not indices:
        return "(该章未关联幻灯片页码，无法提取成品报告正文。)"

    by_idx = _slides_by_index(parsed)
    parts: list[str] = []
    used = 0

    for si in indices:
        slide = by_idx.get(si)
        if not slide:
            frag = f"--- 第 {si} 页 ---\n(解析结果中无此页数据)"
            if used + len(frag) + 2 > cap_total:
                parts.append("…(后续页因长度上限省略)")
                break
            parts.append(frag)
            used += len(frag) + 2
            continue

        lines: list[str] = [f"--- 第 {si} 页 ---"]
        plab = str(slide.get("page_type_label") or "").strip()
        ptype = str(slide.get("page_type") or "").strip()
        if plab or ptype:
            lines.append(f"(页类型: {plab or ptype})")
        cc = slide.get("component_count")
        if isinstance(cc, int) and cc >= 0:
            lines.append(f"(本页解析到约 {cc} 个内容组件)")

        comps = slide.get("components") or []
        if isinstance(comps, list):
            for c in comps:
                if not isinstance(c, dict):
                    continue
                ct = str(c.get("type") or "").strip().lower()
                if ct in _SKIP_TYPES_FOR_EXCERPT:
                    continue
                txt = str(c.get("text") or "").strip()
                if not txt:
                    continue
                if len(txt) > cap_comp:
                    txt = txt[: cap_comp - 1] + "…"
                nm = str(c.get("name") or "").strip()
                label = ct or "text"
                if nm and len(nm) <= 48 and "\n" not in nm:
                    lines.append(f"[{label}|{nm}] {txt}")
                else:
                    lines.append(f"[{label}] {txt}")

        block = "\n".join(lines)
        if used + len(block) + 2 > cap_total:
            remain = cap_total - used - 80
            if remain > 120:
                parts.append(block[:remain] + "\n…(该章成品报告摘录已截断)")
            else:
                parts.append("…(该章剩余页因长度上限省略)")
            break
        parts.append(block)
        used += len(block) + 2

    out = "\n\n".join(parts).strip()
    return out if out else "(该章幻灯片中未解析到可见正文文本。)"


def template_chapter_titles(template: dict[str, Any]) -> tuple[list[str], list[dict[str, str]]]:
    """按 sort 排序后的标题与每章 hint，供模型参考。"""
    chs = template.get("chapters")
    if not isinstance(chs, list):
        return [], []
    rows = sorted(
        (c for c in chs if isinstance(c, dict)),
        key=lambda c: int(c.get("sort") or 0),
    )
    titles: list[str] = []
    meta: list[dict[str, str]] = []
    for c in rows:
        title = str(c.get("title") or "").strip() or "未命名章节"
        hint = str(c.get("hint") or "").strip()
        titles.append(title)
        meta.append({"title": title, "hint": hint})
    return titles, meta


def _field_catalog_for_llm(flat: list[dict[str, str]], preview_len: int = 120) -> list[dict[str, str]]:
    cat: list[dict[str, str]] = []
    for f in flat:
        v = f.get("value") or ""
        prev = v if len(v) <= preview_len else v[:preview_len] + "…"
        cat.append(
            {
                "key": f["key"],
                "label": f.get("label") or _label_for_key(f["key"]),
                "value_preview": prev,
            },
        )
    return cat


def _dashscope_assign_fields(
    slots_for_model: list[dict[str, Any]],
    field_catalog: list[dict[str, str]],
) -> dict[str, Any]:
    api_key = os.getenv("DASHSCOPE_API_KEY") or os.getenv("OPENAI_API_KEY")
    model = (os.getenv("DASHSCOPE_MODEL") or "qwen3-max").strip()
    if not api_key:
        raise RuntimeError("未检测到环境变量 DASHSCOPE_API_KEY（或 OPENAI_API_KEY）。")

    system_prompt = (
        "你是学业报告助手。用户上传的 PPT 是已填写好的成品报告（非空模板）。"
        "每个 slot 的 pptChapterExcerpt 来自解析结果：包含该块对应各页的正文、标题、表格单元格等文本，以及组件类型线索。"
        "首个 slot 可能为「首页/封面」：请根据成品首页中的标题、副标题、人名、日期等，分配姓名、学期、报告副标题等基础字段。"
        "后续 slot 为各「章」：请对照成品里该章实际在写什么，把学生数据字段（仅用 field_catalog 里的 key）"
        "分配到各块：每块只挂载与正文语义相关的字段；同一 key 可出现在多块若合理。"
        "templateTitle 与 templateHint 来自章节模板或首页说明，可作辅助；以 pptChapterExcerpt 为准判断内容需求。"
        "不要编造 key；只使用 field_catalog 中出现的 key。"
        "只返回 JSON，不要 Markdown。"
    )
    user_obj = {
        "task": "为每个 slotIndex 选择若干 field key（须结合 pptChapterExcerpt；含首页与各章）",
        "slots": slots_for_model,
        "field_catalog": field_catalog,
        "output_schema": {
            "assignments": [
                {"slotIndex": 0, "fieldKeys": ["basic.studentName", "learning.studyIntent"]},
            ],
        },
    }
    user_prompt = (
        "请输出 JSON，格式严格为：\n"
        '{"assignments":[{"slotIndex":整数,"fieldKeys":["key1",...]},...]}\n'
        "要求：\n"
        "1) 每个 slotIndex 必须出现且仅对应一章，从 0 递增到 N-1。\n"
        "2) fieldKeys 中的字符串必须完全等于 field_catalog 中某项的 key。\n"
        "3) 若某章无合适字段，fieldKeys 可为空数组。\n"
        "4) 务必阅读每个 slot 的 pptChapterExcerpt，按成品报告该章实际内容匹配学生字段。\n"
        "输入：\n"
        f"{json.dumps(user_obj, ensure_ascii=False)}"
    )

    response = requests.post(
        "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "temperature": 0.25,
            "max_tokens": 4096,
        },
        timeout=int(os.getenv("DASHSCOPE_TIMEOUT_SEC", "120")),
    )
    if response.status_code >= 400:
        raise RuntimeError(f"模型调用失败：HTTP {response.status_code} - {response.text}")
    data = response.json()
    content = data["choices"][0]["message"]["content"]
    return extract_json_from_text(content)


def _normalize_assignments(
    raw: dict[str, Any],
    num_slots: int,
    valid_keys: set[str],
) -> dict[int, list[str]]:
    """slotIndex -> 去重后的 key 列表。"""
    by_slot: dict[int, list[str]] = {i: [] for i in range(num_slots)}
    items = raw.get("assignments")
    if not isinstance(items, list):
        return by_slot
    for it in items:
        if not isinstance(it, dict):
            continue
        try:
            si = int(it.get("slotIndex"))
        except (TypeError, ValueError):
            continue
        if si < 0 or si >= num_slots:
            continue
        keys = it.get("fieldKeys")
        if not isinstance(keys, list):
            continue
        seen: set[str] = set()
        for k in keys:
            if not isinstance(k, str):
                continue
            k = k.strip()
            if not k or k not in valid_keys or k in seen:
                continue
            seen.add(k)
            by_slot[si].append(k)
    return by_slot


def _fallback_round_robin(flat: list[dict[str, str]], num_slots: int) -> dict[int, list[str]]:
    if num_slots <= 0 or not flat:
        return {i: [] for i in range(max(0, num_slots))}
    buckets: dict[int, list[str]] = {i: [] for i in range(num_slots)}
    for idx, f in enumerate(flat):
        buckets[idx % num_slots].append(f["key"])
    return buckets


def resolve_chapter_reference(
    task_id: str,
    chapter_template_id: str,
    student_data_id: str,
    *,
    use_llm: bool = True,
) -> dict[str, Any]:
    """
    返回与前端 chapter_ref_json 一致的 structure: { version, slots }.
    若 use_llm 为 False 或未配置 API Key，则退回轮询分配（便于测试）。
    """
    tid = (task_id or "").strip()
    ctid = (chapter_template_id or "").strip()
    sid = (student_data_id or "").strip()
    if not tid or not ctid or not sid:
        raise ValueError("请提供 task_id、chapter_template_id、student_data_id。")

    parsed = get_parsed_from_cache(tid)
    if not parsed:
        raise ValueError("未找到该 PPT 解析记录，请重新选择或解析。")

    tpl = db_mod.get_chapter_template(ctid)
    if not tpl:
        raise ValueError("未找到该章节模板。")

    stu = db_mod.get_student_record(sid)
    if not stu:
        raise ValueError("未找到该学生数据。")

    slot_rows = ppt_reference_slot_rows(parsed)
    chapter_only = ppt_chapter_slot_rows(parsed)
    n_ch = len(chapter_only)
    if n_ch <= 0:
        raise ValueError("当前 PPT 未解析出「章」块，无法对齐章节模板。")

    tpl_name = str(tpl.get("name") or "").strip()

    titles, meta = template_chapter_titles(tpl)
    n_tpl = len(titles)
    if n_tpl > n_ch:
        raise ValueError(
            f"章节模板含 {n_tpl} 个章节，当前 PPT 仅 {n_ch} 个「章」块，数量不匹配。",
        )

    flat = flatten_student_record(stu)
    key_to_field = {f["key"]: f for f in flat}
    valid_keys = set(key_to_field.keys())

    num_slots = len(slot_rows)
    slots_for_model: list[dict[str, Any]] = []
    ch_i = 0
    for i in range(num_slots):
        row = slot_rows[i] if i < len(slot_rows) else {}
        kind = str(row.get("kind") or "")
        slide_list = row.get("slides") if isinstance(row, dict) else []
        excerpt = build_chapter_ppt_report_excerpt(parsed, slide_list or [])
        if kind == "cover":
            slots_for_model.append(
                {
                    "slotIndex": i,
                    "templateTitle": tpl_name,
                    "templateHint": "封面/首页：templateTitle 默认为当前章节模板名称，用作报告主标题写入封面；可按需修改。"
                    "结合成品副标题、姓名、日期等位置分配学生基础信息字段。",
                    "pptChapterExcerpt": excerpt,
                },
            )
        else:
            title = titles[ch_i] if ch_i < len(titles) else ""
            hint = ""
            if ch_i < len(meta):
                hint = meta[ch_i].get("hint") or ""
            ch_i += 1
            slots_for_model.append(
                {
                    "slotIndex": i,
                    "templateTitle": title,
                    "templateHint": hint,
                    "pptChapterExcerpt": excerpt,
                },
            )

    by_slot: dict[int, list[str]]
    if use_llm and valid_keys:
        try:
            catalog = _field_catalog_for_llm(flat)
            raw = _dashscope_assign_fields(slots_for_model, catalog)
            by_slot = _normalize_assignments(raw, num_slots, valid_keys)
            # 若模型漏掉某些 slot，保持空；若全部为空且应有字段，用轮询兜底
            total_assigned = sum(len(v) for v in by_slot.values())
            if total_assigned == 0 and flat:
                log.warning("模型未返回有效字段分配，使用轮询兜底")
                by_slot = _fallback_round_robin(flat, num_slots)
        except Exception:
            log.exception("大模型分配学生字段失败，使用轮询兜底")
            by_slot = _fallback_round_robin(flat, num_slots)
    elif flat and num_slots:
        by_slot = _fallback_round_robin(flat, num_slots)
    else:
        by_slot = {i: [] for i in range(num_slots)}

    slots_out: list[dict[str, Any]] = []
    ch_out = 0
    for i, row in enumerate(slot_rows):
        slide_list = row.get("slides") or []
        slides_str = ",".join(str(int(s)) for s in slide_list if str(s).strip().isdigit())
        kind = str(row.get("kind") or "")
        if kind == "cover":
            title = tpl_name
        else:
            title = titles[ch_out] if ch_out < len(titles) else ""
            ch_out += 1
        keys = by_slot.get(i, [])
        fields = []
        for k in keys:
            f = key_to_field.get(k)
            if f:
                fields.append({"key": f["key"], "label": f["label"], "value": f["value"]})
        slots_out.append(
            {
                "slotIndex": i,
                "slides": slides_str,
                "templateTitle": title,
                "fields": fields,
                "screenshots": [],
            },
        )

    return {
        "version": 2,
        "slots": slots_out,
        "allStudentFields": [{"key": f["key"], "label": f["label"], "value": f["value"]} for f in flat],
    }
