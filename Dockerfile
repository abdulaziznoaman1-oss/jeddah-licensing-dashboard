FROM python:3.12-slim
WORKDIR /app
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 PORT=8501 APP_ENV=production
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY . /app/
RUN mkdir -p /var/data
EXPOSE 8501
CMD ["python", "server.py"]
