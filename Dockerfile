# =====================================================================
#  发票报销工具 —— 容器版（Docker / NAS）
#
#  同一个程序两种跑法：本机双击 bat 时数据就在程序目录里；这里靠 FB_* 环境
#  变量把数据和监听地址接出来（各变量的含义见 app_paths.py 顶部）。
#
#  构建并启动：docker compose up -d
#  改了代码之后：docker compose up -d --build
# =====================================================================
FROM python:3.12-slim-bookworm

ENV TZ=Asia/Shanghai \
    DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

# 时区：不设的话容器里是 UTC，报销单上的「制表时间」会差 8 小时
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone

# chromium        —— 报销单 PDF 是用无头浏览器打印 HTML 出来的（本机用的是 Edge，
#                    同一个内核、同一套命令行参数），不装的话只能出 Excel。
# fonts-noto-cjk  —— 中文字体。不装的话 PDF 里所有中文会变成方块 ——
#                    这个坑很隐蔽（页面在浏览器里看着好好的，打出来才发现）。
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
        chromium \
        fonts-noto-cjk \
        tzdata \
        ca-certificates \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Python 依赖。前三个都是纯 Python 包，x86_64 / arm64 都装得上，不用编译。
# argon2-cffi：口令哈希用的 Argon2id（改进建议里点名要的），有预编译 wheel；
#   万一这个源拉不到也不让构建失败 —— 程序会自动退回标准库的 scrypt，功能不受影响。
RUN pip install --no-cache-dir openpyxl pypdf pdfminer.six \
 && { pip install --no-cache-dir argon2-cffi \
      || echo "[warn] argon2-cffi 未安装，口令哈希退回标准库 scrypt"; }

COPY . /app

# 服务器模式：绑 0.0.0.0（否则容器外访问不到）、数据落到挂载卷、
# 票据文件夹用容器内的路径。口令不在这里写死 —— 在 compose 里给 FB_PASSWORD。
ENV FB_SERVER=1 \
    FB_HOST=0.0.0.0 \
    FB_PORT=8766 \
    FB_DATA_DIR=/data \
    FB_SOURCE=/票据

RUN mkdir -p /data /票据
VOLUME ["/data"]

EXPOSE 8766

# 探活：/api/ping 是不需要登录的，正好用来做健康检查
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD python -c "import urllib.request;urllib.request.urlopen('http://127.0.0.1:8766/api/ping',timeout=3)"

CMD ["python", "/app/app_web.py"]
