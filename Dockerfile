FROM python:3.11-slim

# 시스템 폰트 설치 (한글 워터마크용)
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-nanum fonts-nanum-extra \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 의존성 먼저 설치 (캐시 활용)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 앱 코드 복사 (Cloud Run용 app_cloud.py 사용)
COPY app_cloud.py .
COPY cloud_storage_helper.py .
COPY src/ src/
COPY data/ data/

# Streamlit 설정
RUN mkdir -p /root/.streamlit
COPY .streamlit/config.toml /root/.streamlit/config.toml

# Cloud Run은 PORT 환경변수를 제공
ENV PORT=8080
EXPOSE 8080

# reports.pkg는 Cloud Storage에서 로드하므로 복사하지 않음
# 계약정보관리 엑셀도 data/ 폴더에 포함

CMD ["streamlit", "run", "app_cloud.py", \
     "--server.port=8080", \
     "--server.address=0.0.0.0", \
     "--server.headless=true", \
     "--browser.gatherUsageStats=false", \
     "--server.fileWatcherType=none"]
