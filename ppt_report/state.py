"""
进程内状态（缓存、异步任务字典）。
服务重启后丢失；解析结果由 models.db 持久化并可回填缓存。
"""
from __future__ import annotations

import threading
from pathlib import Path

PARSE_CACHE: dict[str, dict] = {}
TEMPLATE_PATHS: dict[str, Path] = {}
LAST_GENERATION: dict[str, dict] = {}

GENERATE_JOBS: dict[str, dict] = {}
GENERATE_JOBS_LOCK = threading.Lock()

PARSE_JOBS: dict[str, dict] = {}
PARSE_JOBS_LOCK = threading.Lock()
