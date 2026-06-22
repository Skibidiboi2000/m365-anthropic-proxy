FROM python:3.12-slim

WORKDIR /app

RUN pip install fastapi uvicorn httpx

COPY proxy.py .

CMD ["uvicorn", "proxy:app", "--host", "0.0.0.0", "--port", "8787"]
