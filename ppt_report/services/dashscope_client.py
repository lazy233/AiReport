"""阿里云 DashScope OpenAI 兼容 Chat Completions（与 text_generation / 章节引用等共用环境变量）。"""
from __future__ import annotations

import os

import requests

DASHSCOPE_COMPAT_CHAT_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"


def dashscope_credentials() -> tuple[str, str]:
    """返回 (api_key, model)。未配置 api_key 时抛出 RuntimeError。"""
    api_key = os.getenv("DASHSCOPE_API_KEY") or os.getenv("OPENAI_API_KEY")
    model = (os.getenv("DASHSCOPE_MODEL") or "qwen3-max").strip()
    if not api_key:
        raise RuntimeError("未检测到环境变量 DASHSCOPE_API_KEY（或 OPENAI_API_KEY）。")
    return api_key, model


def dashscope_chat_completion(
    *,
    system: str,
    user: str,
    temperature: float = 0.2,
    max_tokens: int | None = None,
    timeout_sec: int | None = None,
) -> str:
    """
    调用 Chat Completions，返回 assistant 的 content 文本。
    默认 max_tokens、超时与 text_generation._dashscope_generate_normalized 一致：
    DASHSCOPE_MAX_TOKENS（默认 8192）、DASHSCOPE_TIMEOUT_SEC（默认 120）。
    """
    api_key, model = dashscope_credentials()
    if max_tokens is None:
        max_tokens = int(os.getenv("DASHSCOPE_MAX_TOKENS", "8192"))
    if timeout_sec is None:
        timeout_sec = int(os.getenv("DASHSCOPE_TIMEOUT_SEC", "120"))

    response = requests.post(
        DASHSCOPE_COMPAT_CHAT_URL,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        json={
            "model": model,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            "temperature": temperature,
            "max_tokens": max_tokens,
        },
        timeout=timeout_sec,
    )
    if response.status_code >= 400:
        raise RuntimeError(
            f"模型调用失败：HTTP {response.status_code} - {response.text[:500]}",
        )
    data = response.json()
    try:
        return data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("模型返回格式异常，请稍后重试。") from exc
