FROM python:3.11-slim
WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY common_utils.py alert_*.py ./
CMD ["python", "alert_nouki.py"]
