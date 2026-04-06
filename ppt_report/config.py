"""应用配置（环境变量 / 默认值）。"""
from __future__ import annotations

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

try:
    from dotenv import load_dotenv

    load_dotenv(BASE_DIR / ".env")
except ImportError:
    pass

UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

# 生成历史对应的回填成品 PPT 缓存目录（与模板 uploads/{task_id}.pptx 分离）
FILLED_EXPORT_DIR = UPLOAD_DIR / "filled_exports"
FILLED_EXPORT_DIR.mkdir(parents=True, exist_ok=True)

# 生成历史与上述成品缓存的保留天数（超时自动清理）
GENERATION_HISTORY_RETENTION_DAYS = max(
    1,
    int(os.getenv("GENERATION_HISTORY_RETENTION_DAYS", "3")),
)

# 后台清理任务间隔（秒）
GENERATION_HISTORY_CLEANUP_INTERVAL_SEC = max(
    60,
    int(os.getenv("GENERATION_HISTORY_CLEANUP_INTERVAL_SEC", "3600")),
)

MAX_UPLOAD_MB = 200

DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "postgresql+psycopg2://postgres:123456@127.0.0.1:5432/ppt_report_platform",
).strip()

GENERATION_HARD_LENGTH_CAP = False
GENERATION_SOFT_MAX_CHARS = int(os.getenv("GENERATION_SOFT_MAX_CHARS", "500000"))
# 标题/副标题/占位符类组件的后处理字数上限（默认勿过小，否则 normalize 会把生成结果截成几个字符）
GENERATION_HEADING_MAX_CHARS = int(os.getenv("GENERATION_HEADING_MAX_CHARS", "128"))


def max_content_bytes() -> int:
    return MAX_UPLOAD_MB * 1024 * 1024
