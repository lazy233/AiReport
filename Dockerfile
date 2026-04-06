# 生产镜像：Gunicorn 承载 Flask 应用
FROM python:3.12-slim

WORKDIR /app

# 国内访问 deb.debian.org 常较慢：构建前将 APT 源换为阿里云（海外构建传 USE_CN_APT_MIRROR=0）
ARG USE_CN_APT_MIRROR=1
RUN set -eux; \
    if [ "$USE_CN_APT_MIRROR" = "1" ] && [ -f /etc/apt/sources.list.d/debian.sources ]; then \
      sed -i 's|http://deb.debian.org/debian|https://mirrors.aliyun.com/debian|g' /etc/apt/sources.list.d/debian.sources; \
      sed -i 's|https://deb.debian.org/debian|https://mirrors.aliyun.com/debian|g' /etc/apt/sources.list.d/debian.sources; \
      sed -i 's|http://security.debian.org/debian-security|https://mirrors.aliyun.com/debian-security|g' /etc/apt/sources.list.d/debian.sources; \
      sed -i 's|https://security.debian.org/debian-security|https://mirrors.aliyun.com/debian-security|g' /etc/apt/sources.list.d/debian.sources; \
    fi; \
    apt-get update \
    && apt-get install -y --no-install-recommends \
        libpq5 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
# pip 访问 files.pythonhosted.org 在国内易超时；加长超时并支持构建参数换源
# 海外构建可：docker-compose build --build-arg PIP_INDEX_URL=https://pypi.org/simple web
ARG PIP_INDEX_URL=https://pypi.tuna.tsinghua.edu.cn/simple
RUN pip install --no-cache-dir --default-timeout=300 \
    -i "${PIP_INDEX_URL}" \
    -r requirements.txt

COPY . .

# 运行时创建 uploads（与 config.UPLOAD_DIR 一致；若挂载卷则卷会覆盖为空目录）
RUN mkdir -p uploads

ENV PYTHONUNBUFFERED=1 \
    FLASK_APP=app.py

EXPOSE 5000

# 大文件上传与生成可能较慢，超时与 worker 可按机器调整
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--threads", "4", "--timeout", "120", "app:app"]
