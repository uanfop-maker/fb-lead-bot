FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 8080
CMD gunicorn --workers=1 --bind 0.0.0.0:${PORT:-8080} --timeout 60 app:app
