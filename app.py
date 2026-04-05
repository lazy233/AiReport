"""应用入口：生产/开发请通过本模块的 ``app`` 实例启动。"""
from __future__ import annotations

from ppt_report import create_app

app = create_app()

if __name__ == "__main__":
    app.run(debug=True)
