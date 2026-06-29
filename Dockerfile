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

ENV GRADIO_SERVER_NAME=0.0.0.0
ENV GRADIO_SERVER_PORT=10000
ENV PYTHONUNBUFFERED=1

EXPOSE 10000

# Render already provides a public URL. Gradio's localhost accessibility check can fail
# inside Render's container network, so we patch the launch call at startup.
CMD ["/bin/sh", "-c", "sed -i 's/share=False/share=True/' /app/app.py && python app.py"]
