"""文案生成：长度策略、构建模型载荷、DashScope 调用与批处理。"""
from __future__ import annotations

import json
import os
from collections.abc import Callable

import requests

from ppt_report import config
from ppt_report.services.llm_json import extract_json_from_text
from ppt_report.services.page_types import compute_chapter_selection_groups
from ppt_report.services.pptx_document import estimate_max_chars


def heading_effective_cap(comp: dict) -> int | None:
    role = str(comp.get("heading_cap_type") or comp.get("type") or "").lower()
    if role not in ("title", "subtitle", "placeholder"):
        return None
    ceiling = max(1, config.GENERATION_HEADING_MAX_CHARS)
    w = float((comp.get("position_cm") or {}).get("width") or 0)
    if w > 0:
        by_width = max(1, min(ceiling, int(w * 0.95)))
        return min(ceiling, by_width)
    return ceiling


def apply_heading_line_limits(comp: dict, max_chars: int, ref_len: int) -> tuple[int, int]:
    cap = heading_effective_cap(comp)
    if cap is None:
        return max_chars, ref_len
    max_chars = min(int(max_chars), cap)
    if ref_len > 0:
        ref_len = min(int(ref_len), cap)
        max_chars = min(max_chars, ref_len)
    return max(1, max_chars), ref_len


def resolve_generation_max_chars(comp: dict) -> tuple[int, int]:
    raw = (comp.get("text") or "").strip()
    pos = comp.get("position_cm") or {}
    w = float(pos.get("width") or 0)
    h = float(pos.get("height") or 0)
    comp_type = str(comp.get("type", "")).lower()
    stored = max(4, int(comp.get("max_chars", 60)))

    if w > 0 and h > 0:
        if comp_type in ("title", "subtitle"):
            geo = estimate_max_chars(w, h, comp_type)
        else:
            area = max(1.0, w * h)
            geo = max(15, min(1200, int(area * 8)))
        layout = max(stored, geo)
    else:
        layout = stored

    if not raw:
        return layout, 0
    olen = len(raw)
    return max(1, olen), olen


def truncate_to_max_chars(text: str, max_chars: int) -> str:
    if max_chars <= 0 or len(text) <= max_chars:
        return text
    cut = text[:max_chars]
    for sep in ("\n", "。", "！", "？", "；", "，", "、"):
        idx = cut.rfind(sep)
        if idx > max_chars * 0.55:
            return cut[: idx + 1].rstrip()
    sp = cut.rfind(" ")
    if sp > max_chars * 0.55:
        return cut[:sp].rstrip()
    return cut.rstrip()


def should_generate_for_component(comp: dict) -> bool:
    if not comp.get("is_text_editable"):
        return False

    comp_type = str(comp.get("type", "")).lower()
    comp_text = (comp.get("text") or "").strip()
    comp_name = str(comp.get("name", "")).lower()

    if comp_type == "table_cell":
        return True
    if comp_type in {"title", "subtitle", "placeholder", "body"}:
        return True
    if comp_text:
        return True

    text_box_markers = ("文本框", "text box", "textbox")
    return any(marker in comp_name for marker in text_box_markers)


def build_model_payload(parsed: dict, selected_slides: list[int], topic: str, extra_content: str) -> dict:
    hard = config.GENERATION_HARD_LENGTH_CAP
    slides_payload = []

    for slide in parsed["slides"]:
        if slide["slide_index"] not in selected_slides:
            continue

        components = []
        for comp in slide["components"]:
            if should_generate_for_component(comp):
                parsed_text = comp.get("text", "")
                if hard:
                    max_chars, ref_len = resolve_generation_max_chars(comp)
                else:
                    max_chars = config.GENERATION_SOFT_MAX_CHARS
                    ref_len = 0
                max_chars, ref_len = apply_heading_line_limits(comp, max_chars, ref_len)
                entry = {
                    "index": comp["index"],
                    "parent_index": comp.get("parent_index"),
                    "level": comp.get("level", 0),
                    "type": comp["type"],
                    "parsed_text": parsed_text,
                    "original_text": parsed_text,
                    "reference_char_count": ref_len,
                    "max_chars": max_chars,
                }
                if comp.get("heading_cap_type"):
                    entry["heading_cap_type"] = comp["heading_cap_type"]
                if hard and ref_len > 0:
                    entry["suggested_length_range"] = {
                        "min_chars": 1,
                        "max_chars": ref_len,
                    }
                components.append(entry)

        if components:
            slides_payload.append(
                {
                    "slide_index": slide["slide_index"],
                    "components": components,
                }
            )

    merged_intent = "\n".join(
        part for part in (topic.strip(), extra_content.strip()) if part
    ).strip()

    return {
        "topic": topic,
        "extra_content": extra_content,
        "primary_intent": merged_intent or topic or extra_content,
        "layout_reference_rule": (
            "parsed_text 只用于参考：是否中英文对照、条目数量、序号样式（如壹贰叁肆）；"
            "具体写什么，必须服从 primary_intent，"
            "禁止沿用与 primary_intent 无关的旧主题（如「开学/新学期」类话术）。"
        ),
        "length_rule": (
            (
                "当某组件 reference_char_count > 0 时：generated_text 的字符数必须严格 <= reference_char_count，"
                "计数方式与参考一致（parsed_text 去首尾空白后的字符个数，含标点、字母、空格）；"
                "超出视为错误；宁可略短不要超长。"
                "reference_char_count 为 0 时：仅受该组件 max_chars 约束。"
            )
            if hard
            else (
                "字数不设硬性上限：以语义完整与幻灯片可读为准；可与原 parsed_text 长短差异较大；"
                "避免无意义重复堆砌。"
                "例外：type 或 heading_cap_type 为 title/subtitle/placeholder 的必须不超过 max_chars（单行，避免换行乱版；上限已按框宽折算）。"
            )
        ),
        "slides": slides_payload,
    }


def normalize_generated_result(model_result: dict, payload: dict) -> dict:
    normalized = {"slides": []}
    model_slides = {}

    for slide in model_result.get("slides", []):
        slide_index = slide.get("slide_index")
        if isinstance(slide_index, int):
            model_slides[slide_index] = slide.get("components", [])

    for slide in payload.get("slides", []):
        slide_index = slide["slide_index"]
        model_components = model_slides.get(slide_index, [])
        by_index = {str(c.get("index")): c for c in model_components}

        out_components = []
        for expected in slide["components"]:
            expected_index = str(expected["index"])
            model_comp = by_index.get(expected_index, {})
            text = str(model_comp.get("generated_text", "")).strip()
            max_chars = int(expected.get("max_chars", 60))
            if not text:
                text = fallback_generated_text(
                    expected,
                    payload.get("primary_intent") or payload.get("topic", ""),
                )
            if len(text) > max_chars:
                text = truncate_to_max_chars(text, max_chars)
            out_components.append(
                {
                    "index": expected_index,
                    "generated_text": text,
                }
            )

        normalized["slides"].append(
            {
                "slide_index": slide_index,
                "components": out_components,
            }
        )

    return normalized


def fallback_generated_text(component: dict, topic: str) -> str:
    original_text = (component.get("original_text") or "").strip()
    if original_text:
        return original_text

    comp_type = str(component.get("type", "")).lower()
    if comp_type == "title":
        cap = max(1, config.GENERATION_HEADING_MAX_CHARS)
        return (topic or "主题").strip()[:cap] or "主题"
    if comp_type == "subtitle":
        return "核心要点"
    if comp_type == "table_cell":
        return "—"
    return "待补充"


def _dashscope_generate_normalized(payload: dict) -> dict:
    api_key = os.getenv("DASHSCOPE_API_KEY") or os.getenv("OPENAI_API_KEY")
    model = os.getenv("DASHSCOPE_MODEL", "qwen3-max")
    if not api_key:
        raise RuntimeError("未检测到环境变量 DASHSCOPE_API_KEY（或 OPENAI_API_KEY）。")

    hard = config.GENERATION_HARD_LENGTH_CAP
    if hard:
        length_system_extra = (
            "4) 字数上限（硬约束）：见 JSON 中的 length_rule。\n"
            "reference_char_count>0 时 generated_text 长度必须 <= reference_char_count；"
            "为 0 时遵守 max_chars。\n"
        )
        parsed_text_hint = (
            "2) parsed_text 仅是版式与长度参考：条数、中英对照、序号字（如壹贰叁肆）、每框大约多长；"
        )
        length_user_extra = (
            "— 长度：每个组件必须同时遵守 length_rule；有 reference_char_count 的，"
            "生成字数不得多于参考 parsed_text 字数（可与参考等长或更短）。\n"
            "— 无参考（reference_char_count 为 0）时遵守 max_chars。\n\n"
        )
    else:
        length_system_extra = (
            "4) 字数：正文类（text/body 等）可展开；"
            "类型 为 标题/副标题/占位符 的必须视为单行标题，generated_text 字符数严格 <= 该组件 max_chars，忌冗长导致换行乱版。\n"
        )
        parsed_text_hint = (
            "2) parsed_text 是版式与结构参考：条数、中英对照、序号字（如壹贰叁肆）；"
            "不必拘泥于与原文字数接近；"
        )
        length_user_extra = (
            "— 长度：title/subtitle/placeholder 每项必须 <= 各自 max_chars；其它组件以版式可读为准。\n\n"
        )

    system_prompt = (
        "你是PPT文案助手。用户意图在 JSON 的 primary_intent（由 topic 与 extra_content 合并）中，这是唯一主题来源。\n"
        "优先级（必须遵守）：\n"
        "1) 所有 generated_text 的语义必须与 primary_intent 一致（人物、场景、文体）。\n"
        f"{parsed_text_hint}"
        "不得保留与 primary_intent 冲突的旧主题词（如用户要求「期末总结」却继续写「开学/新学期」）。\n"
        "3) 当旧文案与 primary_intent 冲突时，必须整体改写为新主题下的目录或段落，而不是轻微微调。\n"
        f"{length_system_extra}"
        "只返回 JSON、不要 Markdown；只输出给定 index。"
    )
    user_prompt = (
        "根据下列 JSON 为每个 index 填写 generated_text。\n\n"
        "输出格式：\n"
        "{\n"
        '  "slides":[\n'
        "    {\n"
        '      "slide_index": 1,\n'
        '      "components":[{"index":"1","generated_text":"..."}]\n'
        "    }\n"
        "  ]\n"
        "}\n\n"
        "硬性规则：\n"
        "— 禁止照抄 parsed_text 中与 primary_intent 无关的章节标题或活动名；目录结构可保留，文字必须换成符合 primary_intent 的内容。\n"
        "— 反例：primary_intent 写「赵锦汉的期末总结」，输出里仍出现「开学计划安排」「新学期的目标」——不允许。\n"
        "— 正例：仍为四条中英目录，但中文与英文都改为「学期回顾/学业表现/不足与反思/寒假与展望」等期末语境。\n"
        "— 只输出输入里已有的 index。\n"
        f"{length_user_extra}"
        "输入：\n"
        f"{json.dumps(payload, ensure_ascii=False)}"
    )

    max_tokens = int(os.getenv("DASHSCOPE_MAX_TOKENS", "8192"))
    request_timeout = int(os.getenv("DASHSCOPE_TIMEOUT_SEC", "120"))

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
            "temperature": 0.45,
            "max_tokens": max_tokens,
        },
        timeout=request_timeout,
    )
    if response.status_code >= 400:
        raise RuntimeError(f"模型调用失败：HTTP {response.status_code} - {response.text}")

    data = response.json()
    content = data["choices"][0]["message"]["content"]
    parsed_model_result = extract_json_from_text(content)
    return normalize_generated_result(parsed_model_result, payload)


def compute_generation_batches(selected_slides: list[int]) -> list[list[int]]:
    selected_sorted = sorted({int(x) for x in selected_slides})
    batch_size = int(os.getenv("GENERATE_SLIDE_BATCH_SIZE", "8"))
    if batch_size <= 0:
        return [selected_sorted]
    return [
        selected_sorted[i : i + batch_size]
        for i in range(0, len(selected_sorted), batch_size)
    ]


def generate_text_with_bailian(
    parsed: dict,
    selected_slides: list[int],
    topic: str,
    extra_content: str,
    progress: Callable[[int, int, list[int]], None] | None = None,
) -> dict:
    batches = compute_generation_batches(selected_slides)

    merged_slides: list[dict] = []
    for batch_index, batch in enumerate(batches):
        if progress:
            progress(batch_index, len(batches), batch)
        payload = build_model_payload(parsed, batch, topic, extra_content)
        if not payload["slides"]:
            continue
        part = _dashscope_generate_normalized(payload)
        merged_slides.extend(part.get("slides", []))

    if not merged_slides:
        raise RuntimeError("所选页面没有可生成的有效文本框。")
    merged_slides.sort(key=lambda s: s["slide_index"])
    return {"slides": merged_slides}


CHAPTER_REF_MUST_USE = (
    "【强制】以上「本章参考数据」中的事实、数字、姓名与表述必须在当前批次的幻灯片 generated_text 中得到体现，"
    "可分配到多个文本框；不得整章忽略。截图类素材未提供，勿臆造图片内容。"
)

POST_CHAPTER_HINT = (
    "【本批为首页与/或目录】在 primary_intent 主题下生成，与前面已写章节在术语、人称与叙事上保持一致，"
    "可作总述或导航；勿引入与主题无关的新剧情。"
)


def format_slot_reference_excluding_screenshots(slot: dict) -> str:
    if not isinstance(slot, dict):
        return ""
    lines: list[str] = []
    tt = str(slot.get("templateTitle") or "").strip()
    if tt:
        lines.append(f"模板章节名：{tt}")
    for f in slot.get("fields") or []:
        if not isinstance(f, dict):
            continue
        label = str(f.get("label") or f.get("key") or "").strip()
        val = str(f.get("value") or "").strip()
        if label or val:
            lines.append(f"- {label or '字段'}：{val}")
    return "\n".join(lines).strip()


def parse_chapter_ref_slots(chapter_ref: dict | None) -> list[dict]:
    if not isinstance(chapter_ref, dict):
        return []
    slots = chapter_ref.get("slots")
    return [x for x in slots if isinstance(x, dict)] if isinstance(slots, list) else []


def generate_single_model_batch(
    parsed: dict,
    slide_indices: list[int],
    topic: str,
    extra_content: str,
) -> dict:
    batch_slides = sorted({int(x) for x in slide_indices})
    payload = build_model_payload(parsed, batch_slides, topic, extra_content)
    if not payload.get("slides"):
        return {"slides": []}
    return _dashscope_generate_normalized(payload)


def generate_text_orchestrated(
    parsed: dict,
    selected_slides: list[int],
    topic: str,
    base_extra: str,
    chapter_ref: dict | None,
    progress: Callable[[int, int, list[int]], None] | None = None,
) -> dict:
    """
    按章节分批调用模型：每一「章」块单独请求，并注入对应 slot 的参考数据（不含截图）。
    全部章完成后，再请求首页+目录；最后请求 misc 块。
    若解析结果中没有任何「章」块，则退回按页数批量（GENERATE_SLIDE_BATCH_SIZE）的旧逻辑。
    """
    selected_set = {int(x) for x in selected_slides}
    groups = compute_chapter_selection_groups(parsed)
    chapters_in_order = [g for g in groups if str(g.get("kind")) == "chapter"]
    slots = parse_chapter_ref_slots(chapter_ref)

    def pick_slides(g: dict) -> list[int]:
        out: list[int] = []
        for s in g.get("slides") or []:
            try:
                n = int(s)
            except (TypeError, ValueError):
                continue
            if n in selected_set:
                out.append(n)
        return sorted(out)

    if not chapters_in_order:
        return generate_text_with_bailian(parsed, selected_slides, topic, base_extra, progress=progress)

    plan: list[tuple[str, list[int], str]] = []

    for ci, g in enumerate(chapters_in_order):
        sl = pick_slides(g)
        if not sl:
            continue
        ref_block = ""
        if ci < len(slots):
            fb = format_slot_reference_excluding_screenshots(slots[ci])
            if fb:
                ref_block = f"【本章参考数据】\n{fb}\n{CHAPTER_REF_MUST_USE}\n"
        extra = "\n\n".join(p for p in (base_extra.strip(), ref_block.strip()) if p)
        plan.append((f"chapter_{ci + 1}", sl, extra))

    cover_slides: list[int] = []
    toc_slides: list[int] = []
    misc_slides: list[int] = []
    for g in groups:
        k = str(g.get("kind"))
        sl = pick_slides(g)
        if not sl:
            continue
        if k == "cover":
            cover_slides.extend(sl)
        elif k == "toc":
            toc_slides.extend(sl)
        elif k == "misc":
            misc_slides.extend(sl)

    post_slides = sorted(set(cover_slides) | set(toc_slides))
    if post_slides:
        extra = "\n\n".join(p for p in (base_extra.strip(), POST_CHAPTER_HINT.strip()) if p)
        plan.append(("cover_toc", post_slides, extra))

    if misc_slides:
        plan.append(("misc", sorted(set(misc_slides)), base_extra.strip()))

    if not plan:
        raise RuntimeError("所选页面没有可生成的有效文本框。")

    total = len(plan)
    merged_slides: list[dict] = []
    for pi, (_phase, sls, extra) in enumerate(plan):
        if progress:
            progress(pi, total, sls)
        part = generate_single_model_batch(parsed, sls, topic, extra)
        merged_slides.extend(part.get("slides", []))

    if not merged_slides:
        raise RuntimeError("所选页面没有可生成的有效文本框。")
    merged_slides.sort(key=lambda s: s.get("slide_index", 0))
    return {"slides": merged_slides}
