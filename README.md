# PPT 组件解析网页

一个最小可用的网页应用：上传 `pptx` 文件，解析每页组件类型与作用，并在页面中打印结果。

## 功能

- 上传 `pptx` 文件
- 解析每页的组件（文本、标题、图片、表格、图表、线条/连接线、组合等）
- 给出每个组件的作用说明
- 展示结构化结果和原始 JSON
- 输入主题/补充内容，选择页面后调用阿里百炼模型生成组件填充文本

## 运行方式

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

浏览器打开：`http://127.0.0.1:5000`

## 大模型配置

- 优先读取 `DASHSCOPE_API_KEY`
- 若未设置，会回退读取 `OPENAI_API_KEY`
- 模型名读取 `DASHSCOPE_MODEL`，默认 `qwen3.5-plus`

## 说明

- 当前解析核心依赖 `python-pptx`，仅支持 `pptx`。
- 上传 `ppt` 时会提示先转换为 `pptx`。
