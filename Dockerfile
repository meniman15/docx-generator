FROM python:3.10-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy source code and template
COPY app.py templateToDocx.py templateV2.docx ./

# Set environment variables
ENV PORT=8080
ENV TEMPLATE_PATH=/app/templateV2.docx
ENV PYTHONUNBUFFERED=1

EXPOSE 8080

CMD ["python", "app.py"]
