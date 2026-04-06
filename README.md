# AiReport（PPT 解析与报告生成）

基于 Flask 的 Web 应用：上传 `pptx` 模板，解析幻灯片中的组件（文本、标题、图片、表格、图表、形状等），结合大模型按主题与页面生成填充内容，并支持生成历史、学生数据等配套能力。前后端一体，默认开发地址为 `http://127.0.0.1:5000`。

## 功能概览

- **上传与解析**：上传 `pptx`，按页展示组件类型与说明；仅支持 `pptx`（`ppt` 需先另存为 `pptx`）。
- **智能生成**：按主题/补充说明选择页面，调用模型生成并回填占位文本（阿里百炼 DashScope，可配置模型名与超时等）。
- **章节与引用**：章节模板、参考图等流程（见站内页面与相关服务模块）。
- **数据与历史**：PostgreSQL 持久化；生成历史与成品缓存可配置保留天数与清理间隔。

## 技术栈

| 类别 | 说明 |
|------|------|
| 运行时 | Python 3 + Flask |
| 数据库 | PostgreSQL（SQLAlchemy 2.x，`psycopg2-binary`） |
| PPT | `python-pptx` |
| 大模型 | 通义千问等（DashScope HTTP API；兼容用 `OPENAI_API_KEY` 作为备用键名） |

## 环境准备

1. **Python 3**：建议 3.10+。
2. **PostgreSQL**：本地监听 `127.0.0.1:5432`（或按你的连接串修改），并创建与 `.env` 中一致的数据库（示例库名见下方 `.env.example`）。
3. **依赖安装**：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

4. **环境变量**：复制 `.env.example` 为 `.env`，按需修改数据库连接与模型相关变量。

## 快速启动（Windows）

在项目根目录执行（会尝试启动本机 PostgreSQL 服务，再启动 Flask）：

```powershell
.\start.ps1
```

浏览器访问：`http://127.0.0.1:5000`

若未使用 `start.ps1`，需自行保证数据库已就绪，然后：

```powershell
python app.py
```

生产环境请使用 WSGI 服务器挂载 `app:app`，勿长期以 `debug=True` 暴露公网。

## 主要环境变量

| 变量 | 说明 |
|------|------|
| `DATABASE_URL` | PostgreSQL 连接串，见 `.env.example` |
| `DASHSCOPE_API_KEY` | 阿里百炼 API Key（优先） |
| `OPENAI_API_KEY` | 未设置 `DASHSCOPE_API_KEY` 时作为备用 |
| `DASHSCOPE_MODEL` | 模型名，默认 `qwen3-max` |
| `DASHSCOPE_MAX_TOKENS` / `DASHSCOPE_TIMEOUT_SEC` | 生成请求 token 上限与超时（秒） |
| `GENERATE_SLIDE_BATCH_SIZE` | 按页批处理大小等 |
| `GENERATION_HISTORY_RETENTION_DAYS` | 生成历史/成品缓存保留天数（默认 3） |
| `GENERATION_HISTORY_CLEANUP_INTERVAL_SEC` | 后台清理间隔（秒，默认 3600） |

更多调优项（如章节解析字数上限、学生指导相关超时等）见源码中 `os.getenv` 调用。

## 项目结构（简要）

- `app.py`：应用入口。
- `ppt_report/`：Flask 应用工厂、蓝图、业务服务（解析、生成、缓存、异步任务等）。
- `static/`、`templates/`：前端静态资源与 Jinja 模板。
- `uploads/`：上传与生成导出缓存目录（运行时创建）。

## 许可证与贡献

若仓库未单独声明许可证，以仓库根目录文件为准。欢迎通过 Issue / PR 反馈问题与改进。
