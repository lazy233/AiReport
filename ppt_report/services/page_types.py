"""页面类型：构建 LLM 载荷、归一化、章节分组、DashScope 分类。"""
from __future__ import annotations

import json
import os

import requests

from ppt_report.constants import PAGE_TYPE_LABELS
from ppt_report.services.llm_json import extract_json_from_text


def page_type_label(page_type: str | None) -> str:
    return PAGE_TYPE_LABELS.get(str(page_type or "").strip(), PAGE_TYPE_LABELS["unknown"])


def build_page_type_payload(parsed: dict) -> dict:
    slides_payload = []
    for slide in parsed.get("slides", []):
        text_components = []
        text_samples = []
        type_counts: dict[str, int] = {}
        for comp in slide.get("components", []):
            comp_type = str(comp.get("type", "unknown"))
            type_counts[comp_type] = type_counts.get(comp_type, 0) + 1
            comp_text = (comp.get("text") or "").strip()
            if comp_text:
                text_components.append(
                    {
                        "index": comp.get("index"),
                        "type": comp_type,
                        "text": comp_text,
                        "level": comp.get("level", 0),
                    }
                )
                text_samples.append(comp_text[:120])

        slides_payload.append(
            {
                "slide_index": slide["slide_index"],
                "component_count": slide.get("component_count", 0),
                "top_level_component_count": slide.get("top_level_component_count", 0),
                "type_counts": type_counts,
                "text_component_count": len(text_components),
                "text_samples": text_samples[:8],
                "text_components": text_components[:12],
            }
        )

    return {
        "task": "请判断每一页 PPT 的页面类型",
        "allowed_page_types": list(PAGE_TYPE_LABELS.keys())[:-1],
        "page_type_definition": {
            "cover": "整份PPT首页/封面，通常包含总标题、副标题、作者/日期等",
            "toc": "目录页，列出多个章节、部分、目录项",
            "chapter_cover": "章节扉页/过渡页，通常只引出某一章节，信息量少于目录页和正文页",
            "content": "正文页，承载具体说明、数据、内容要点",
        },
        "slides": slides_payload,
    }


def normalize_page_types(model_result: dict, parsed: dict) -> dict:
    valid_types = {"cover", "toc", "chapter_cover", "content"}
    by_index = {}
    for item in model_result.get("slides", []):
        slide_index = item.get("slide_index")
        if not isinstance(slide_index, int):
            continue
        raw_type = str(item.get("page_type", "")).strip()
        if raw_type not in valid_types:
            raw_type = "unknown"
        by_index[slide_index] = {
            "page_type": raw_type,
            "page_type_reason": str(item.get("reason", "")).strip(),
        }

    for slide in parsed.get("slides", []):
        info = by_index.get(slide["slide_index"], {"page_type": "unknown", "page_type_reason": ""})
        slide["page_type"] = info["page_type"]
        slide["page_type_label"] = page_type_label(info["page_type"])
        slide["page_type_reason"] = info["page_type_reason"]

    return parsed


def compute_chapter_selection_groups(parsed: dict) -> list[dict]:
    def push_group(entry: dict) -> None:
        sl = entry.get("slides") or []
        entry["default_selected"] = any(int(s) <= 5 for s in sl)
        groups.append(entry)

    slides = sorted(parsed.get("slides", []), key=lambda s: s["slide_index"])
    groups: list[dict] = []
    i = 0
    chapter_seq = 0

    while i < len(slides):
        s = slides[i]
        idx = s["slide_index"]
        ptype = str(s.get("page_type", "unknown")).strip()

        if ptype == "cover":
            push_group({"id": "sel_cover", "kind": "cover", "label": "首页", "slides": [idx]})
            i += 1
        elif ptype == "toc":
            push_group({"id": "sel_toc", "kind": "toc", "label": "目录", "slides": [idx]})
            i += 1
        elif ptype == "chapter_cover":
            chapter_seq += 1
            chapter_slides = [idx]
            i += 1
            while i < len(slides):
                ns = slides[i]
                npt = str(ns.get("page_type", "unknown")).strip()
                if npt == "chapter_cover":
                    break
                if npt in ("content", "unknown"):
                    chapter_slides.append(ns["slide_index"])
                    i += 1
                elif npt in ("cover", "toc"):
                    break
                else:
                    i += 1
            push_group(
                {
                    "id": f"sel_chapter_{chapter_seq}",
                    "kind": "chapter",
                    "label": f"第 {chapter_seq} 章（第 {chapter_slides[0]}–{chapter_slides[-1]} 页）",
                    "slides": chapter_slides,
                }
            )
        else:
            block = [idx]
            i += 1
            while i < len(slides):
                ns = slides[i]
                npt = str(ns.get("page_type", "unknown")).strip()
                if npt in ("cover", "toc", "chapter_cover"):
                    break
                block.append(ns["slide_index"])
                i += 1
            push_group(
                {
                    "id": f"sel_misc_{block[0]}",
                    "kind": "misc",
                    "label": f"其他页面（第 {block[0]}–{block[-1]} 页）",
                    "slides": block,
                }
            )

    return groups


def demo_chapter_selection_groups() -> list[dict]:
    """未选择 PPT 时用于界面展示的示例章节分组（幻灯片页码为示意，不可真实生成）。"""
    raw: list[tuple[str, str, str, list[int]]] = [
        ("sel_cover", "cover", "首页", [1]),
        ("sel_toc", "toc", "目录", [2]),
        ("sel_chapter_1", "chapter", "第 1 章（第 3–4 页）", [3, 4]),
        ("sel_chapter_2", "chapter", "第 2 章（第 5–6 页）", [5, 6]),
        ("sel_chapter_3", "chapter", "第 3 章（第 7–8 页）", [7, 8]),
        ("sel_chapter_4", "chapter", "第 4 章（第 9–11 页）", [9, 10, 11]),
        ("sel_chapter_5", "chapter", "第 5 章（第 12–14 页）", [12, 13, 14]),
        ("sel_chapter_6", "chapter", "第 6 章（第 15–18 页）", [15, 16, 17, 18]),
        ("sel_chapter_7", "chapter", "第 7 章（第 19–23 页）", [19, 20, 21, 22, 23]),
        ("sel_chapter_8", "chapter", "第 8 章（第 24–25 页）", [24, 25]),
    ]
    groups: list[dict] = []
    for eid, kind, label, slides in raw:
        entry: dict = {"id": eid, "kind": kind, "label": label, "slides": slides}
        entry["default_selected"] = any(int(s) <= 5 for s in slides)
        groups.append(entry)
    return groups


def classify_page_types_with_bailian(parsed: dict) -> dict:
    api_key = os.getenv("DASHSCOPE_API_KEY") or os.getenv("OPENAI_API_KEY")
    model = (os.getenv("DASHSCOPE_MODEL") or "qwen3-max").strip()
    if not api_key:
        raise RuntimeError("未检测到环境变量 DASHSCOPE_API_KEY（或 OPENAI_API_KEY）。")

    payload = build_page_type_payload(parsed)
    system_prompt = (
        "你是PPT页面结构分类助手。"
        "你需要仅根据每页的组件统计、文本样本与文本组件列表，判断页面类型。"
        "只允许输出 cover、toc、chapter_cover、content 四种之一。"
        "不要输出 unknown，除非你完全无法判断，此时也尽量在四者中选最接近的。"
        "请只返回 JSON，不要输出 Markdown。"
    )
    user_prompt = (
        "请判断每一页的页面类型。\n"
        "输出 JSON 格式必须为：\n"
        "{\n"
        '  "slides": [\n'
        '    {"slide_index": 1, "page_type": "cover", "reason": "..."}\n'
        "  ]\n"
        "}\n"
        "判断要求：\n"
        "1) 首页/封面 -> cover\n"
        "2) 目录页 -> toc\n"
        "3) 章节扉页/过渡页 -> chapter_cover\n"
        "4) 正文页 -> content\n"
        "5) 优先看页面整体结构，而不是只看单个词。\n"
        "输入数据如下：\n"
        f"{json.dumps(payload, ensure_ascii=False)}"
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
            "temperature": 0.2,
            "max_tokens": 4096,
        },
        timeout=90,
    )
    if response.status_code >= 400:
        raise RuntimeError(f"页面分类失败：HTTP {response.status_code} - {response.text}")

    data = response.json()
    content = data["choices"][0]["message"]["content"]
    parsed_model_result = extract_json_from_text(content)
    return normalize_page_types(parsed_model_result, parsed)
