"""与业务无关的常量。"""

ALLOWED_EXTENSIONS = {"pptx", "ppt"}
ALLOWED_CONTENT_EXTENSIONS = {"txt", "md", "markdown", "json", "csv"}

PAGE_TYPE_LABELS = {
    "cover": "首页",
    "toc": "目录",
    "chapter_cover": "章节扉页",
    "content": "正文页",
    "unknown": "未识别",
}

GENERATE_JOBS_MAX = 64
PARSE_JOBS_MAX = 48
