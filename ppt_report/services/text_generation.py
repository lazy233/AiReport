"""文案生成：长度策略、构建模型载荷、DashScope 调用与批处理。"""
from __future__ import annotations

import json
import os
import re
from collections.abc import Callable

import requests

from ppt_report import config
from ppt_report.services.chapter_reference_resolve import ppt_reference_slot_rows
from ppt_report.services.llm_json import extract_json_from_text
from ppt_report.services.page_types import compute_chapter_selection_groups
from ppt_report.services.pptx_document import estimate_max_chars

_STAR = "★"


def _star_rating_cap(parsed_text: str) -> int:
    return (parsed_text or "").count(_STAR)


_CJK_RE = re.compile(r"[\u4e00-\u9fff]")
# 连续 3 个及以上拉丁字母视为原文已含英文词（含 API、the 等），不锁「禁止加英文」
_LATIN_RUN_3_RE = re.compile(r"[A-Za-z]{3,}")


def _original_monolingual_zh_lock(parsed_text: str) -> bool:
    """
    原文以中文为主且未出现连续英文字母片段时，生成结果不得擅自加英文或双语翻译。
    用于正文、标题等凡含中文且无英文词的组件。
    """
    t = (parsed_text or "").strip()
    if not t:
        return False
    if not _CJK_RE.search(t):
        return False
    if _LATIN_RUN_3_RE.search(t):
        return False
    return True


def heading_effective_cap(comp: dict) -> int | None:
    role = str(comp.get("heading_cap_type") or comp.get("type") or "").lower()
    if role not in ("title", "subtitle", "placeholder"):
        return None
    layout_mc = max(0, int(comp.get("max_chars") or 0))
    cfg = max(8, config.GENERATION_HEADING_MAX_CHARS)
    # 成品报告标题往往较长：取「环境配置」与「解析时按框宽估算的 max_chars」中较大者，避免压到几十个字符以内
    cap = max(cfg, layout_mc) if layout_mc else cfg
    cap = min(cap, 800)
    w = float((comp.get("position_cm") or {}).get("width") or 0)
    if w > 0:
        wide_hint = max(layout_mc, int(w * 3.2))
        cap = min(max(cap, wide_hint), 800)
    return max(1, cap)


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
                stripped = (parsed_text or "").strip()
                original_len = len(stripped)
                entry = {
                    "index": comp["index"],
                    "parent_index": comp.get("parent_index"),
                    "level": comp.get("level", 0),
                    "type": comp["type"],
                    "parsed_text": parsed_text,
                    "original_text": parsed_text,
                    "reference_char_count": ref_len,
                    "original_text_char_count": original_len,
                    "max_chars": max_chars,
                }
                if comp.get("heading_cap_type"):
                    entry["heading_cap_type"] = comp["heading_cap_type"]
                if _original_monolingual_zh_lock(parsed_text):
                    entry["original_has_no_english"] = True
                cap_stars = _star_rating_cap(parsed_text)
                if cap_stars > 0:
                    entry["star_rating"] = {
                        "apply": True,
                        "max_stars": cap_stars,
                        "rule": (
                            "依据 primary_intent / extra_content（含本章参考数据）中的学生成绩、分数或等级评定本格星级；"
                            "仅使用字符 ★ 重复表示得分（颗数 1～max_stars，整数颗）；"
                            "禁止用 -、—、横线、空白或单独数字替代整块星标；"
                            "除 ★ 序列外(parsed_text 里)的固定文字、标点、换行须保留。"
                        ),
                    }
                if hard and ref_len > 0:
                    entry["suggested_length_range"] = {
                        "min_chars": 1,
                        "max_chars": ref_len,
                    }
                tr = comp.get("text_runs")
                if isinstance(tr, list) and tr:
                    entry["text_runs"] = tr
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
    if not merged_intent:
        merged_intent = (
            "以成品报告各框 parsed_text 为底稿，仅按参考数据替换人名、科目、数字等可变信息；"
            "无额外文字约束。"
        )

    return {
        "topic": topic,
        "extra_content": extra_content,
        "primary_intent": merged_intent or topic or extra_content,
        "layout_reference_rule": (
            "当前幻灯片来自已排版好的成品报告：每个组件的 parsed_text 即该文本框内原有成稿。"
            "任务是「数据与称谓替换」——在尽量保留原有句式、段落结构、条数、标点、语气与叙述角度的前提下，"
            "把人名、学号、科目名、分数、学校等可变信息换成 primary_intent / extra_content（含「本章参考数据」等）中的参考内容。"
            "禁止按主题整段重写、另写一套叙事或随意删改与数据无关的描述性正文；"
            "仅在原文与参考事实明显冲突时，可做最小必要改写以消歧。"
            "【语言与形式】必须严格遵守各组件 parsed_text 的呈现形式："
            "若该框原文仅为中文（无中英对照、无单独成行的英文标题/译句），则 generated_text 也必须仅为中文，"
            "禁止添加英文翻译、括注译文、副标题行或「中文 / English」式双语；"
            "若原文仅为英文或已是固定中英对照/换行双语结构，替换后须保持同样的语言结构与行数关系，不得擅自单语化或加译。"
            "禁止以「更专业」「国际化」等理由自作主张加英文。"
            "【禁止擅自加译】凡组件 JSON 中 original_has_no_english 为 true：表示 parsed_text 为单语中文成稿（无连续英文字母词），"
            "generated_text 必须仍为单语中文——不得新增英文单词、括注译名、尾注英文、独立英文标题行或「中文 / English」式对照；"
            "正文页段落、列表、说明文字同样遵守；替换参考数据时用中文表述，勿附英文说明。"
            "【星级占位符 ★】若某组件 parsed_text 含字符 ★ 或 JSON 中 star_rating.apply 为 true："
            "该格须按学生参考数据（分数、等级、优良中差等）换算为整数颗 ★（不超过 star_rating.max_stars），"
            "输出中对应部分仅由 ★ 组成或与原文同样混排（保留非星文字）；"
            "禁止用 -、— 或纯数字代替星形展示。"
            "【解析样式 text_runs】若某组件 JSON 中含 text_runs 数组：为解析阶段记录的逐 run 文本与样式（bold、font_size_pt、换行等）。"
            "生成结果仍为单列字符串 generated_text，但须据此保留应有的换行（\\n）与段落/条数关系，勿随意压成一行；"
            "加粗字号等无法在本接口的纯文本中表达，回写 PPT 时当前仍按整框纯文本写入，版式以模板为准。"
        ),
        "length_rule": (
            (
                "当某组件 reference_char_count > 0 时：generated_text 字符数必须严格 <= reference_char_count（与硬长度模式一致）。"
                "若参考数据某字段明显长于 parsed_text 中对应片段（如科目名、小标题），须缩写、简称或概括，"
                "使整段仍满足上限且视觉行长与成品报告接近。"
                "reference_char_count 为 0 时：遵守 max_chars，并尽量使总长度接近 original_text_char_count。"
            )
            if hard
            else (
                "以 parsed_text 与 original_text_char_count 为锚：generated_text 应与原版成稿长度同量级，避免无故扩写或大幅缩短。"
                "标题、副标题、表格单元格、单行标签等若需写入更长的参考用语（如「国际经济学与国际形势」替换「数据结构」），"
                "必须压缩为与 original_text_char_count 接近的简称或短语，优先保证版面不换行、不撑版。"
                "type 或 heading_cap_type 为 title/subtitle/placeholder 的另须严格 <= max_chars。"
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
            "5) 字数：见 JSON 中 length_rule。reference_char_count>0 时长度必须 <= 该值；"
            "参考数据长于原文时须缩写后再写入。\n"
        )
        parsed_text_hint = (
            "3) parsed_text 是成品报告该框原文：保留其叙述骨架与版式意图，只替换其中应对齐到参考数据的片段；"
            "每个组件另有 original_text_char_count 表示原文字符数，替换后总长度须服从 length_rule。\n"
        )
        length_user_extra = (
            "— 长度：严格遵守 length_rule 与各组件 reference_char_count / max_chars。\n"
            "— 科目名、小标题等参考用语若远长于 parsed_text 对应片段，必须压缩到与原版字数接近。\n\n"
        )
    else:
        length_system_extra = (
            "5) 字数：以 original_text_char_count 与 max_chars 为锚；"
            "标题类单行且 <= max_chars；长参考词须缩写以贴近原框字数。\n"
        )
        parsed_text_hint = (
            "3) parsed_text 是成品该框原文：在保留句式与结构的前提下替换数据；"
            "original_text_char_count 供你对齐替换后的总长度（尤其标题、科目名）。\n"
        )
        length_user_extra = (
            "— 长度：title/subtitle/placeholder <= max_chars；"
            "其它框生成结果长度宜与 original_text_char_count 同量级，必要时压缩参考数据用语。\n\n"
        )

    system_prompt = (
        "你是 PPT 成品报告「局部替换」助手。JSON 中 primary_intent 由 topic（可选的额外条件/约束）与 extra_content 合并；"
        "若用户未填写 topic，则可能是默认说明。其作用是标明参考数据与可选约束，不是让你整页重写主题的指令。\n"
        "优先级（必须遵守）：\n"
        "1) 默认策略：在每条 parsed_text 基础上做最小必要修改——把人名、称谓、学号、科目、分数、学校等换成参考数据中的对应信息，"
        "保留其余描述性语句、修辞与段落逻辑。\n"
        "1b) 星级占位：若某组件含 star_rating.apply 为 true（或 parsed_text 含 ★），须按 star_rating.max_stars 上限、"
        "结合参考数据中的成绩/等级换算为整数颗 ★；对应片段只使用 ★ 字符，禁止用 -、—、横线或单独数字替代整段星标。\n"
        "2) 语言与版式形式（硬规则）：逐条看该组件的 parsed_text。"
        "原文只有中文、没有成行的英文标题或中英对照结构时，generated_text 也必须只有中文，禁止加英文译名、Summary、括注 (English) 等。"
        "原文只有英文时不要加中文说明行。原文已是中英双语、上下行对照或「中文 / English」格式时，替换后必须保持同一种双语版式，不要改成单语。"
        "不得因个人习惯为纯中文模板「补翻译」。\n"
        "2c) 强约束：若某组件 JSON 含 original_has_no_english: true，该框 parsed_text 判定为「原文无英文词」。"
        "generated_text 不得出现任何新增的连续拉丁字母片段（≥3 个字母视为英文词），"
        "禁止在句末、括号内或换行追加英文；topic、extra_content 即使用英文写说明，也不得改变该框单语中文输出。"
        "适用于标题、正文、表格单元格等所有此类组件。\n"
        f"{parsed_text_hint}"
        "4) 禁止：把成稿改写成与原文风格、长度、结构无关的新文案；无故更换目录条数或中英对照格式；"
        "在参考数据已覆盖应替换信息时仍保留旧人名旧科目等。\n"
        f"{length_system_extra}"
        "只返回 JSON、不要 Markdown；只输出给定 index。"
    )
    user_prompt = (
        "根据下列 JSON，为每个 index 输出替换后的 generated_text（成品报告局部替换，非全文创作）。\n\n"
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
        "— 以 parsed_text 为底稿：能留则留，只改应对齐参考数据的词与数。\n"
        "— 语言形式：parsed_text 纯中文则 generated_text 必须纯中文（无整行/整段新增英文）；"
        "parsed_text 里怎么排版（单语/双语/换行），输出就怎么排版，禁止画蛇添足加翻译。\n"
        "— original_has_no_english 为 true 的组件：输出全文不得含连续 3 个以上拉丁字母（即不要写任何英文单词）；"
        "不得用英文复述参考数据；常见误加如「（Summary）」「 / Introduction」一律禁止。\n"
        "— 正例（中文-only 标题）：原文「学习总结」→ 只替换数据时仍输出中文短标题如「学期总结」，"
        "不得输出「学习总结 / Summary」或「学习总结（Semester Summary）」。\n"
        "— 正例（数据替换）：原文「张三 · 数据结构 · 95」、参考李四、科目「国际经济学与国际形势」、88 → "
        "可「李四 · 国经形势 · 88」等中文缩写，字数贴近原版；若原文无英文，不要加科目英文名。\n"
        "— 星级评分：原文「课堂表现 ★★★★★」且参考数据有分数或等级 → 按 max_stars 与数据给出「课堂表现 ★★★★★」这类结果，"
        "星数与数据一致；禁止输出「课堂表现 -----」或「课堂表现 5」。\n"
        "— 反例：把半页课程介绍整段改成与原文无关的新论述——不允许。\n"
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
            "temperature": 0.28,
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
    "【强制】以上「本章参考数据」中的事实、数字、姓名等必须在当前批次各框的 generated_text 中得到体现（可拆到多个文本框），"
    "但须在保留各框 parsed_text 整体表述与结构的前提下替换，禁止因参考数据而整段重写无关正文。"
    "若某参考字段远长于原文对应片段，须缩写至与 original_text_char_count / 版式相近。截图类素材未提供，勿臆造图片内容。"
    "【语言】各框若 parsed_text 为纯中文则输出不得加英文翻译；双语模板则保持双语结构。"
)

POST_CHAPTER_HINT = (
    "【本批为首页与/或目录】同样以成品 parsed_text 为底稿做局部替换：人名、报告主题称谓等与参考数据、前文章节一致即可，"
    "勿改写无关导航句的整体叙事；字数贴近原版；纯中文框勿加英文译行。"
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


def _chapter_ref_slot_index_for_chapter(
    chapter_index: int,
    slot_rows: list[dict],
    n_slots: int,
) -> int:
    """将第 chapter_index 个「章」块映射到 chapter_ref.slots 下标（兼容仅章 slot 的旧数据）。"""
    if n_slots <= 0 or chapter_index < 0:
        return -1
    has_cover = bool(slot_rows) and str(slot_rows[0].get("kind") or "") == "cover"
    n_ch_rows = sum(1 for g in slot_rows if str(g.get("kind") or "") == "chapter")
    if has_cover:
        if n_slots == len(slot_rows):
            idx = 1 + chapter_index
            return idx if idx < n_slots else -1
        if n_ch_rows > 0 and n_slots == n_ch_rows:
            return chapter_index if chapter_index < n_slots else -1
        idx = 1 + chapter_index
        if idx < n_slots:
            return idx
        return chapter_index if chapter_index < n_slots else -1
    return chapter_index if chapter_index < n_slots else -1


# 原文含这些片段的框更像副标题/信息行，不像整页报告主标题
_COVER_NON_MAIN_TITLE_MARKERS = (
    "姓名",
    "学号",
    "班级",
    "专业",
    "学院",
    "导师",
    "指导教师",
    "日期",
    "时间",
    "答辩",
    "汇报人",
    "提交",
    "电话",
    "邮箱",
    "课程代码",
    "任课教师",
    "评分",
    "成绩",
)


def _score_cover_title_candidate(comp: dict) -> float:
    """分越高越像「封面主标题」框（依据解析 type、原文、版式位）。"""
    ct = str(comp.get("type") or "").lower()
    hc = str(comp.get("heading_cap_type") or "").lower()
    raw = (comp.get("text") or "").strip()
    name = str(comp.get("name") or "")
    name_l = name.lower()
    pos = comp.get("position_cm") or {}
    try:
        top = float(pos.get("top") or 999.0)
    except (TypeError, ValueError):
        top = 999.0
    try:
        width = float(pos.get("width") or 0.0)
    except (TypeError, ValueError):
        width = 0.0

    score = 0.0
    if ct == "title":
        score += 220.0
    elif hc == "title":
        score += 40.0

    # 靠上、偏宽的框更像主标题
    score -= top * 1.2
    score += min(width * 1.8, 36.0)

    if "副标题" in name or "subtitle" in name_l:
        score -= 130.0
    if ("标题" in name or "title" in name_l) and "副" not in name and "sub" not in name_l:
        score += 35.0

    if not raw:
        score += 25.0
    else:
        nlines = raw.count("\n") + 1
        score -= max(0, nlines - 1) * 22.0
        score -= max(0, len(raw) - 32) * 0.75
        if any(m in raw for m in _COVER_NON_MAIN_TITLE_MARKERS):
            score -= 140.0
        digits = sum(1 for c in raw if c.isdigit())
        if digits >= 6:
            score -= min(90.0, digits * 5.0)
        unified = raw.replace(":", "：")
        if "：" in unified:
            a, b = unified.split("：", 1)
            if len(a.strip()) <= 10 and len(b.strip()) > 10:
                score -= 55.0

    return score


def _pick_primary_cover_title_index(parsed: dict, slide_index: int) -> str | None:
    """
    在封面页上，从 type/heading 为 title 的候选框中选出最可能的主标题位（唯一 index）。
    无候选或均不可编辑时返回 None。
    """
    target = int(slide_index)
    candidates: list[dict] = []
    for slide in parsed.get("slides") or []:
        if not isinstance(slide, dict):
            continue
        try:
            si = int(slide.get("slide_index"))
        except (TypeError, ValueError):
            continue
        if si != target:
            continue
        for comp in slide.get("components") or []:
            if not isinstance(comp, dict):
                continue
            if not comp.get("is_text_editable"):
                continue
            ctype = str(comp.get("type") or "").lower()
            if ctype == "table_cell":
                continue
            hc = str(comp.get("heading_cap_type") or "").lower()
            if ctype != "title" and hc != "title":
                continue
            idx = comp.get("index")
            if idx is None:
                continue
            candidates.append(comp)
        break

    if not candidates:
        return None
    best = max(candidates, key=_score_cover_title_candidate)
    return str(best.get("index"))


def apply_cover_main_title_override(
    parsed: dict,
    merged_slides: list[dict],
    cover_slide_indices: list[int],
    template_title: str,
) -> None:
    """生成完成后，仅对每页判定出的「主标题」组件写入用户填写的报告主标题。"""
    tt = (template_title or "").strip()
    if not tt or not merged_slides:
        return
    cover_set = {int(x) for x in cover_slide_indices}
    for block in merged_slides:
        if not isinstance(block, dict):
            continue
        try:
            si = int(block.get("slide_index"))
        except (TypeError, ValueError):
            continue
        if si not in cover_set:
            continue
        primary = _pick_primary_cover_title_index(parsed, si)
        if not primary:
            continue
        comps = block.get("components")
        if not isinstance(comps, list):
            continue
        for c in comps:
            if not isinstance(c, dict):
                continue
            if str(c.get("index")) == primary:
                c["generated_text"] = tt
                break


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
    slot_rows = ppt_reference_slot_rows(parsed)
    n_slots = len(slots)

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
        result = generate_text_with_bailian(parsed, selected_slides, topic, base_extra, progress=progress)
        cover_fb: list[int] = []
        for g in groups:
            if str(g.get("kind")) != "cover":
                continue
            cover_fb.extend(pick_slides(g))
        cover_fb = sorted(set(cover_fb))
        cover_main_fb = ""
        if (
            slot_rows
            and str(slot_rows[0].get("kind") or "") == "cover"
            and n_slots > 0
            and n_slots == len(slot_rows)
            and slots
        ):
            cover_main_fb = str(slots[0].get("templateTitle") or "").strip()
        if cover_main_fb and cover_fb:
            merged_fb = result.get("slides")
            if isinstance(merged_fb, list):
                apply_cover_main_title_override(parsed, merged_fb, cover_fb, cover_main_fb)
        return result

    plan: list[tuple[str, list[int], str]] = []

    for ci, g in enumerate(chapters_in_order):
        sl = pick_slides(g)
        if not sl:
            continue
        ref_block = ""
        si = _chapter_ref_slot_index_for_chapter(ci, slot_rows, n_slots)
        if si >= 0:
            fb = format_slot_reference_excluding_screenshots(slots[si])
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
        cover_ref = ""
        if (
            slot_rows
            and str(slot_rows[0].get("kind") or "") == "cover"
            and n_slots > 0
            and n_slots == len(slot_rows)
        ):
            parts_cov: list[str] = []
            ct0 = str(slots[0].get("templateTitle") or "").strip()
            if ct0:
                hint_lines: list[str] = []
                for csi in sorted(set(cover_slides)):
                    pid = _pick_primary_cover_title_index(parsed, csi)
                    if pid:
                        hint_lines.append(
                            f"第 {csi} 页：仅组件 index 「{pid}」须将 generated_text 设为本条主标题；"
                            "同页其它文本框（含其它 title 类）仍按 parsed_text 做局部数据替换，禁止全部写成同一标题。"
                        )
                block = (
                    "【首页主标题·强制】下列字符串须原样写入各页「主标题位」（由解析依据原文与版式自动判定，见下行 index），"
                    "不得增删、改写或翻译：\n"
                    + ct0
                )
                if hint_lines:
                    block += "\n" + "\n".join(hint_lines)
                parts_cov.append(block)
            fb0 = format_slot_reference_excluding_screenshots(slots[0])
            if fb0:
                parts_cov.append("【首页参考数据】\n" + fb0)
            if parts_cov:
                cover_ref = "\n\n".join(parts_cov) + "\n" + CHAPTER_REF_MUST_USE + "\n"
        extra = "\n\n".join(
            p for p in (base_extra.strip(), cover_ref.strip(), POST_CHAPTER_HINT.strip()) if p
        )
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
    cover_main = ""
    if (
        slot_rows
        and str(slot_rows[0].get("kind") or "") == "cover"
        and n_slots > 0
        and n_slots == len(slot_rows)
        and slots
    ):
        cover_main = str(slots[0].get("templateTitle") or "").strip()
    if cover_main and cover_slides:
        apply_cover_main_title_override(parsed, merged_slides, cover_slides, cover_main)
    return {"slides": merged_slides}
