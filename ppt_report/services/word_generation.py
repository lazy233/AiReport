"""Word 生成：根据学生数据调用大模型，决策并回填表格单元格内容。"""
from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any

from docx import Document
import requests

from ppt_report import config
from ppt_report.models import db
from ppt_report.services.chapter_ref_images import is_safe_task_id
from ppt_report.services.llm_json import extract_json_from_text

_PLACEHOLDER_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_.\-]+)\s*\}\}")
_KV_LINE_RE = re.compile(r"^\s*([^:：]{1,24})\s*[:：]\s*(.*?)\s*$")


def _normalize_key(s: str) -> str:
    return re.sub(r"[\s_:\-：]", "", str(s or "").strip().lower())


def _build_student_value_map(record: dict[str, Any]) -> tuple[dict[str, str], dict[str, str]]:
    profile = record.get("profile") if isinstance(record.get("profile"), dict) else {}
    basic = profile.get("basic") if isinstance(profile.get("basic"), dict) else {}
    learning = profile.get("learning") if isinstance(profile.get("learning"), dict) else {}
    hours = profile.get("hours") if isinstance(profile.get("hours"), dict) else {}
    guidance = profile.get("guidance") if isinstance(profile.get("guidance"), dict) else {}
    raw: dict[str, str] = {
        "studentName": str(record.get("name") or basic.get("studentName") or "").strip(),
        "studentId": str(record.get("studentId") or basic.get("studentId") or "").strip(),
        "schoolName": str(basic.get("schoolName") or "").strip(),
        "grade": str(basic.get("grade") or "").strip(),
        "semesterName": str(basic.get("semesterName") or "").strip(),
        "teacherName": str(basic.get("teacherName") or "").strip(),
        "majorDirection": str(learning.get("majorDirection") or "").strip(),
        "goalSchool": str(learning.get("goalSchool") or "").strip(),
        "goalMajor": str(learning.get("goalMajor") or "").strip(),
        "totalHours": str(hours.get("totalHours") or "").strip(),
        "usedHours": str(hours.get("usedHours") or "").strip(),
        "remainingHours": str(hours.get("remainingHours") or "").strip(),
        "termSummary": str(guidance.get("termSummary") or "").strip(),
        "courseFeedback": str(guidance.get("courseFeedback") or "").strip(),
        "shortTermAdvice": str(guidance.get("shortTermAdvice") or "").strip(),
        "longTermDevelopment": str(guidance.get("longTermDevelopment") or "").strip(),
        "content": str(record.get("content") or "").strip(),
        "remark": str(record.get("remark") or "").strip(),
    }
    values = {k: v for k, v in raw.items() if v}
    aliases: dict[str, str] = {}
    alias_pairs = {
        "姓名": "studentName",
        "学生姓名": "studentName",
        "学号": "studentId",
        "学校": "schoolName",
        "年级": "grade",
        "学期": "semesterName",
        "老师": "teacherName",
        "规划老师": "teacherName",
        "目标学校": "goalSchool",
        "目标专业": "goalMajor",
        "规划方向": "majorDirection",
        "总课时": "totalHours",
        "已用课时": "usedHours",
        "剩余课时": "remainingHours",
        "学期总结": "termSummary",
        "课程反馈": "courseFeedback",
        "短期建议": "shortTermAdvice",
        "长期发展": "longTermDevelopment",
        "备注": "remark",
        "内容": "content",
    }
    for k, key_name in alias_pairs.items():
        val = values.get(key_name, "")
        if val:
            aliases[_normalize_key(k)] = key_name
    return values, aliases


def _extract_table_cells(doc: Document) -> tuple[list[dict[str, Any]], int]:
    cells: list[dict[str, Any]] = []
    table_count = 0
    for t_idx, table in enumerate(doc.tables):
        table_count += 1
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                cells.append(
                    {
                        "cell_id": f"t{t_idx}_r{r_idx}_c{c_idx}",
                        "table_index": t_idx,
                        "row_index": r_idx,
                        "col_index": c_idx,
                        "original_text": (cell.text or "").strip(),
                    },
                )
    return cells, table_count


def _llm_fill_table_cells(
    *,
    student_record: dict[str, Any],
    values: dict[str, str],
    aliases: dict[str, str],
    cells: list[dict[str, Any]],
) -> dict[str, str]:
    api_key = os.getenv("DASHSCOPE_API_KEY") or os.getenv("OPENAI_API_KEY")
    model = (os.getenv("DASHSCOPE_MODEL") or "qwen3-max").strip()
    if not api_key:
        raise RuntimeError("未配置大模型：请设置 DASHSCOPE_API_KEY（或 OPENAI_API_KEY）。")
    max_tokens = int(os.getenv("DASHSCOPE_WORD_MAX_TOKENS", "8192"))
    timeout_sec = int(os.getenv("DASHSCOPE_WORD_TIMEOUT_SEC", "180"))
    profile = student_record.get("profile") if isinstance(student_record.get("profile"), dict) else {}
    system_prompt = (
        "你是学业报告文书填充助手。任务：根据学生数据，逐个判断 Word 表格单元格应填写的最终文本。"
        "你必须保留非目标单元格原文；仅在明确可映射到学生数据时改写。"
        "若单元格原文非空，final_text 不得为空字符串（禁止误清空）。"
        "若单元格含 {{key}} 占位符，优先按 key 替换。"
        "若单元格是“字段名:原值/字段名：原值”，可根据字段名替换。"
        "不要输出解释，只输出 JSON。"
    )
    payload = {
        "student_record": student_record,
        "profile": profile,
        "value_map": values,
        "label_aliases": aliases,
        "table_cells": cells,
    }
    user_prompt = (
        "请输出如下 JSON：\n"
        "{\n"
        '  "cells": [\n'
        '    {"cell_id":"t0_r0_c0","final_text":"..."}\n'
        "  ]\n"
        "}\n"
        "要求：\n"
        "1) 只输出存在于输入中的 cell_id；\n"
        "2) 每个 cell_id 仅出现一次；\n"
        "3) final_text 为最终写回文本；\n"
        "4) 若不应修改，final_text 返回 original_text；\n"
        "5) 全文保持中文风格，不新增无关内容。\n\n"
        f"输入数据：{json.dumps(payload, ensure_ascii=False)}"
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
            "temperature": 0.15,
            "max_tokens": max_tokens,
        },
        timeout=timeout_sec,
    )
    if response.status_code >= 400:
        raise RuntimeError(f"Word 生成模型调用失败：HTTP {response.status_code}")
    data = response.json()
    try:
        content = data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("Word 生成模型返回格式异常。") from exc
    try:
        parsed = extract_json_from_text(content)
    except json.JSONDecodeError as exc:
        raise RuntimeError("Word 生成模型未返回合法 JSON。") from exc
    rows = parsed.get("cells") if isinstance(parsed, dict) else None
    if not isinstance(rows, list):
        raise RuntimeError("Word 生成模型返回缺少 cells 列表。")
    out: dict[str, str] = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        cid = str(row.get("cell_id") or "").strip()
        if not cid:
            continue
        out[cid] = str(row.get("final_text") or "")
    return out


def fill_word_table_for_student(
    task_id: str,
    student_data_id: str,
    *,
    output_path: Path | None = None,
) -> dict[str, Any]:
    tid = (task_id or "").strip()
    sid = (student_data_id or "").strip()
    if not is_safe_task_id(tid):
        raise ValueError("无效的模板任务 ID。")
    if not sid:
        raise ValueError("缺少学生数据 ID。")
    template = config.UPLOAD_DIR / f"{tid}.docx"
    if not template.is_file():
        raise ValueError("未找到 Word 模板文件，请先上传 .docx。")
    rec = db.get_student_record(sid)
    if not rec:
        raise ValueError("未找到学生数据，请重新选择。")
    values, aliases = _build_student_value_map(rec)
    if not values:
        raise ValueError("学生数据为空，无法回填。")

    doc = Document(str(template))
    cells_meta, table_count = _extract_table_cells(doc)
    if table_count <= 0:
        raise ValueError("Word 模板中未检测到表格，无法执行表格回填。")

    llm_updates = _llm_fill_table_cells(
        student_record=rec,
        values=values,
        aliases=aliases,
        cells=cells_meta,
    )

    touched_cells = 0
    placeholder_hits = 0
    kv_rewrite_hits = 0
    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                cid = f"t{t_idx}_r{r_idx}_c{c_idx}"
                original = cell.text or ""
                updated = llm_updates.get(cid, original)
                # 禁止用空串覆盖原有非空内容，避免模型误返回 final_text="" 导致整格被清空
                if (original or "").strip() and not str(updated or "").strip():
                    updated = original
                if updated != original:
                    touched_cells += 1
                    placeholder_hits += len(_PLACEHOLDER_RE.findall(original))
                    kv_rewrite_hits += 1 if _KV_LINE_RE.match(original or "") else 0
                    cell.text = updated

    if output_path is None:
        out = config.FILLED_EXPORT_DIR / f"{tid}_filled.docx"
    else:
        out = output_path
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return {
        "ok": True,
        "task_id": tid,
        "student_data_id": sid,
        "output_kind": "docx",
        "table_count": table_count,
        "touched_cells": touched_cells,
        "placeholder_hits": placeholder_hits,
        "kv_rewrite_hits": kv_rewrite_hits,
        "download_path": str(out),
    }

