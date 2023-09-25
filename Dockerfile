# pythonイメージ用意
# (macOSの場合、rosseta2が必要、Docker Desktopの設定を有効にする必要あり)
FROM --platform=linux/amd64 python:3.11

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

#install google-chrome
RUN wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | apt-key add && \
    echo 'deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main' | tee /etc/apt/sources.list.d/google-chrome.list && \
    apt-get update && \
    apt-get install -y google-chrome-stable
# echo "set up google-chrome"

ENV PATH /usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin:/opt/google/chrome

# install 7zip
RUN apt-get install -y p7zip-full

# 必要なパッケージインストールはpipenvで行う
COPY Pipfile Pipfile.lock ./
RUN python -m pip install --upgrade pip
RUN pip install pipenv && pipenv install --system --deploy

WORKDIR /app
COPY . /app

# 環境用意あとは、起動しっぱなし。execで入れるようにする
RUN echo "running! msm-gas-prepare"