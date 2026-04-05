"""章节参考截图：按 task_id 分目录存储，供后续填入 PPT。"""
from __future__ import annotations

import re
import shutil
import uuid
from pathlib import Path

from ppt_report import config

CHAPTER_REF_SUBDIR = "chapter_ref"
ALLOWED_IMAGE_EXT = frozenset({"png", "jpg", "jpeg", "gif", "webp"})
# 单张截图上限（与全局 MAX 独立，避免单图过大）
MAX_IMAGE_BYTES = max(1, 15 * 1024 * 1024)

_TASK_ID_RE = re.compile(r"^[a-fA-F0-9]{8,64}$")
_STORED_NAME_RE = re.compile(r"^[a-fA-F0-9]{32}\.(png|jpg|jpeg|gif|webp)$", re.I)


def chapter_ref_root() -> Path:
    root = config.UPLOAD_DIR / CHAPTER_REF_SUBDIR
    root.mkdir(parents=True, exist_ok=True)
    return root


def is_safe_task_id(task_id: str) -> bool:
    return bool(task_id and _TASK_ID_RE.match(task_id.strip()))


def is_safe_stored_filename(name: str) -> bool:
    return bool(name and _STORED_NAME_RE.match(name.strip()))


def task_image_dir(task_id: str) -> Path | None:
    tid = (task_id or "").strip()
    if not is_safe_task_id(tid):
        return None
    d = chapter_ref_root() / tid
    d.mkdir(parents=True, exist_ok=True)
    return d


def ext_from_upload(filename: str, mimetype: str | None) -> str | None:
    if not filename or "." not in filename:
        if mimetype:
            mt = mimetype.lower().split(";")[0].strip()
            mapping = {
                "image/png": "png",
                "image/jpeg": "jpg",
                "image/jpg": "jpg",
                "image/gif": "gif",
                "image/webp": "webp",
            }
            return mapping.get(mt)
        return None
    ext = filename.rsplit(".", 1)[1].lower()
    if ext == "jpeg":
        ext = "jpg"
    return ext if ext in ALLOWED_IMAGE_EXT else None


def save_chapter_ref_image(task_id: str, file_storage) -> tuple[dict[str, str] | None, str | None]:
    """
    保存上传图片。成功返回 (item_dict, None)，item 含 storedFilename / urlPath / originalName。
    """
    tid = (task_id or "").strip()
    if not is_safe_task_id(tid):
        return None, "无效的 task_id。"
    if not file_storage or not getattr(file_storage, "filename", None):
        return None, "请选择图片文件。"
    ext = ext_from_upload(file_storage.filename, getattr(file_storage, "mimetype", None))
    if not ext:
        return None, "仅支持 PNG、JPEG、GIF、WebP 图片。"
    raw = file_storage.read()
    if not raw:
        return None, "文件为空。"
    if len(raw) > MAX_IMAGE_BYTES:
        return None, f"单张图片不能超过 {MAX_IMAGE_BYTES // (1024 * 1024)}MB。"

    name = f"{uuid.uuid4().hex}.{ext}"
    d = task_image_dir(tid)
    if not d:
        return None, "无法创建存储目录。"
    path = d / name
    path.write_bytes(raw)

    url_path = f"/api/chapter-ref-images/{tid}/{name}"
    return (
        {
            "storedFilename": name,
            "url": url_path,
            "originalName": (file_storage.filename or name)[:240],
        },
        None,
    )


def image_file_path(task_id: str, stored_name: str) -> Path | None:
    if not is_safe_task_id(task_id) or not is_safe_stored_filename(stored_name):
        return None
    d = chapter_ref_root() / task_id.strip()
    path = (d / stored_name).resolve()
    try:
        d_resolved = d.resolve()
    except OSError:
        return None
    if not str(path).startswith(str(d_resolved)) or not path.is_file():
        return None
    return path


def delete_chapter_ref_image(task_id: str, stored_name: str) -> bool:
    path = image_file_path(task_id, stored_name)
    if not path:
        return False
    try:
        path.unlink(missing_ok=True)
    except OSError:
        return False
    return True


def purge_chapter_ref_task_dir(task_id: str) -> bool:
    """删除该 task 下章节参考截图目录。目录存在并已删除则返回 True。"""
    tid = (task_id or "").strip()
    if not is_safe_task_id(tid):
        return False
    d = chapter_ref_root() / tid
    if not d.is_dir():
        return False
    try:
        shutil.rmtree(d, ignore_errors=True)
    except OSError:
        return False
    return True
