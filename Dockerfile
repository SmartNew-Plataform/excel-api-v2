FROM python:3.13.18-slim-trixie@sha256:e48f59cdba69a3cbd9ed2cf0f348d90d45a8d350d98134cf1070bc106ee9ed90 AS base
WORKDIR /app
RUN apt update -y && apt install -y wget
RUN pip install uvicorn
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
EXPOSE 8000
ENTRYPOINT ["python", "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
