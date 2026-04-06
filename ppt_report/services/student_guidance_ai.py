"""根据学生档案其它维度，调用大模型补全「成长指导」四维文本。"""
from __future__ import annotations

import json
import os
from typing import Any

import requests

from ppt_report.services.llm_json import extract_json_from_text

_GUIDANCE_KEYS = ("termSummary", "courseFeedback", "shortTermAdvice", "longTermDevelopment")


def _norm_profile_slice(src: Any) -> dict[str, Any]:
    if not isinstance(src, dict):
        return {}
    return {k: v for k, v in src.items() if v not in (None, "")}


def _trim_text(s: str, max_len: int) -> str:
    t = (s or "").strip()
    if len(t) <= max_len:
        return t
    return t[: max_len - 1] + "…"


def generate_student_guidance(
    profile: dict[str, Any] | None,
    extra_content: str = "",
) -> dict[str, str]:
    """
    返回 guidance 四维英文字段，供前端写入表单。
    依赖 DASHSCOPE_API_KEY（或 OPENAI_API_KEY）与 DASHSCOPE_MODEL。
    """
    api_key = os.getenv("DASHSCOPE_API_KEY") or os.getenv("OPENAI_API_KEY")
    model = (os.getenv("DASHSCOPE_MODEL") or "qwen3-max").strip()
    if not api_key:
        raise RuntimeError("未配置大模型：请设置环境变量 DASHSCOPE_API_KEY（或 OPENAI_API_KEY）。")

    prof = profile if isinstance(profile, dict) else {}
    basic = _norm_profile_slice(prof.get("basic"))
    learning = _norm_profile_slice(prof.get("learning"))
    hours = _norm_profile_slice(prof.get("hours"))
    guidance_existing = _norm_profile_slice(prof.get("guidance"))

    context: dict[str, Any] = {
        "基础信息": basic,
        "学习画像": learning,
        "课时数据": hours,
    }
    if guidance_existing:
        context["已有成长指导草稿"] = guidance_existing
    content_trim = _trim_text(extra_content, 6000)
    if content_trim:
        context["数据内容摘录"] = content_trim

    system_prompt = (
        "你是资深学业规划与成长指导顾问，擅长根据学生档案撰写学期报告中的「成长指导」类文字。"
        "必须严格依据输入中的事实性信息；不得编造具体分数、奖项、学校录取结果等输入中未出现的内容。"
        "若信息不足，可写概括性建议并提示需补充信息。使用自然、专业、积极的中文。"
        "只输出一个 JSON 对象，不要 Markdown 代码围栏，不要其它解释。"
    )
    user_prompt = (
        "请根据下列学生档案 JSON，撰写成长指导四条内容，用于学期报告或家校沟通。\n\n"
        "输出 JSON 的键名必须完全一致（英文 camelCase），值为中文正文，可适当分段，不要使用 Markdown 标题语法：\n"
        "{\n"
        '  "termSummary": "学期表现概述",\n'
        '  "courseFeedback": "课程反馈与建议",\n'
        '  "shortTermAdvice": "短期学习建议",\n'
        '  "longTermDevelopment": "长期发展规划"\n'
        "}\n\n"
        "要求：\n"
        "1) 与档案中的姓名、年级、科目、课时、升学意向等保持一致；\n"
        "2) 四条内容各有侧重，避免完全重复；\n"
        "3) 若「已有成长指导草稿」非空，可在此基础上润色、扩写而非简单照抄。\n\n"
        "档案与摘录：\n"
        f"{json.dumps(context, ensure_ascii=False)}"
    )

    max_tokens = int(os.getenv("DASHSCOPE_GUIDANCE_MAX_TOKENS", "4096"))
    request_timeout = int(os.getenv("DASHSCOPE_GUIDANCE_TIMEOUT_SEC", "120"))

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
            "temperature": 0.35,
            "max_tokens": max_tokens,
        },
        timeout=request_timeout,
    )
    if response.status_code >= 400:
        raise RuntimeError(f"模型调用失败：HTTP {response.status_code} - {response.text[:500]}")

    data = response.json()
    try:
        raw_content = data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("模型返回格式异常，请稍后重试。") from exc

    try:
        parsed = extract_json_from_text(raw_content)
    except json.JSONDecodeError as exc:
        raise RuntimeError("模型未返回合法 JSON，请重试或缩短「数据内容」后再试。") from exc

    if not isinstance(parsed, dict):
        raise RuntimeError("模型返回不是 JSON 对象。")

    out: dict[str, str] = {}
    for k in _GUIDANCE_KEYS:
        v = parsed.get(k)
        out[k] = ("" if v is None else str(v)).strip()

    return out
