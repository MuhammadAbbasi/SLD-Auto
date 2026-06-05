FROM python:3.11-slim

WORKDIR /app

# Dependencies first - cached unless requirements.txt changes
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Application source
COPY app.py sld_core.py ./
COPY templates/ templates/

# Generation engine - self-contained inside the webapp folder
COPY code/ code/

EXPOSE 5000

CMD ["python", "app.py"]
