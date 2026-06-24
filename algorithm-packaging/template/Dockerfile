# 使用 uv 官方镜像（已内置 Python 和 uv）
FROM ghcr.io/astral-sh/uv:python3.11-bookworm-slim

# 如果算法使用 libsumo，则取消下方 apt-get 注释安装系统依赖
# ENV TZ=Asia/Shanghai
# ENV DEBIAN_FRONTEND=noninteractive
# RUN sed -i 's/deb.debian.org/mirrors.aliyun.com/g' /etc/apt/sources.list.d/debian.sources && \
#     apt-get update && \
#     apt-get install -y \
#     libx11-6 \
#     libxrender1 \
#     libgl1-mesa-glx \
#     libglib2.0-0 \
#     libsm6 \
#     libice6 \
#     libfontconfig1 \
#     libfreetype6 \
#     libxi6 \
#     libxext6 \
#     libxt6 \
#     libxrandr2 \
#     libxcursor1 \
#     libxinerama1 \
#     libxcomposite1 \
#     libxdamage1 \
#     libxfixes3 \
#     libxau6 \
#     libxdmcp6 \
#     libpng16-16 \
#     libexpat1 \
#     && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY . .

# 安装依赖到系统 Python（uv 会读取 pyproject.toml 中配置的阿里云镜像源）
ENV UV_CACHE_DIR=/tmp/.cache/uv
ENV UV_LINK_MODE=copy
# --mount 不展开 ENV 变量，路径须与 UV_CACHE_DIR 保持一致
RUN --mount=type=cache,target=/tmp/.cache/uv \
    uv pip install -r pyproject.toml --system

# 运行主程序（根据实际入口文件修改）
CMD ["python", "main.py"]
