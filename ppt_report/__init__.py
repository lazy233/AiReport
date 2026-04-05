"""PPT 报告平台 Flask 应用工厂。"""
from __future__ import annotations

import logging
import threading

from flask import Flask, render_template
from werkzeug.exceptions import RequestEntityTooLarge

from ppt_report import config
from ppt_report.blueprints.api import api_bp, export_bp
from ppt_report.blueprints.web import web_bp
from ppt_report.models import db as db_mod

_app_log = logging.getLogger(__name__)


def _generation_history_cleanup_loop(app: Flask) -> None:
    import time

    while True:
        time.sleep(config.GENERATION_HISTORY_CLEANUP_INTERVAL_SEC)
        try:
            with app.app_context():
                n = db_mod.cleanup_expired_generation_history()
                if n:
                    _app_log.info("已自动清理过期生成历史 %s 条", n)
        except Exception:  # noqa: BLE001
            _app_log.exception("生成历史自动清理失败")


def create_app() -> Flask:
    app = Flask(
        __name__,
        template_folder=str(config.BASE_DIR / "templates"),
        static_folder=str(config.BASE_DIR / "static"),
    )
    app.config["MAX_CONTENT_LENGTH"] = config.max_content_bytes()
    app.config["DATABASE_URL"] = config.DATABASE_URL

    db_mod.init_db(config.DATABASE_URL)

    if db_mod.db_enabled():
        try:
            n = db_mod.cleanup_expired_generation_history()
            if n:
                _app_log.info("启动时已清理过期生成历史 %s 条", n)
        except Exception:  # noqa: BLE001
            _app_log.exception("启动时生成历史清理失败")
        threading.Thread(
            target=_generation_history_cleanup_loop,
            args=(app,),
            daemon=True,
            name="generation-history-cleanup",
        ).start()

    @app.context_processor
    def inject_db_flags():
        return {"db_enabled": db_mod.db_enabled()}

    app.register_blueprint(web_bp)
    app.register_blueprint(api_bp)
    app.register_blueprint(export_bp)

    @app.errorhandler(RequestEntityTooLarge)
    def handle_large_file(_exc):
        return (
            render_template(
                "pages/assistant_upload.html",
                error=(
                    f"上传文件过大，当前上限为 {config.MAX_UPLOAD_MB}MB，"
                    "请压缩后重试。"
                ),
                max_upload_mb=config.MAX_UPLOAD_MB,
                generation_hard_length_cap=config.GENERATION_HARD_LENGTH_CAP,
            ),
            413,
        )

    return app
