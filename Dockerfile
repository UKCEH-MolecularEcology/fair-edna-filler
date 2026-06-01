FROM python:3.12-slim

WORKDIR /app

# Install LibreOffice for legacy .doc support (optional – adds ~400 MB)
# Uncomment if you need .doc file support:
# RUN apt-get update && apt-get install -y --no-install-recommends libreoffice && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Create runtime directories
RUN mkdir -p uploads outputs

ENV FAIRE_TOOL_HOST=0.0.0.0
ENV FAIRE_TOOL_PORT=8787

EXPOSE 8787

CMD ["python3", "server.py"]
