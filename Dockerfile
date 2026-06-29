FROM python:3.12-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    pandoc \
    build-essential \
    gcc \
    libsndfile1 \
    libcairo2 \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements_colab.txt /app/requirements.txt
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r /app/requirements.txt

COPY . /app

EXPOSE 10000

CMD ["python", "app.py"]
