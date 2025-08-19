FROM python:3.10-slim

# 安装 LibreOffice 和中文字体
RUN apt-get update && \
    apt-get install -y libreoffice libreoffice-calc fonts-noto-cjk fonts-arphic-ukai fonts-arphic-uming && \
    apt-get clean

# 设置工作目录
WORKDIR /app

# 拷贝项目文件
COPY . /app

# 安装 Python 依赖
RUN pip install --no-cache-dir -r requirements.txt

# 设置环境变量
ENV PORT=5000

# 启动 Flask 应用
CMD ["python", "app.py"]

