FROM python:3.10-slim

# System deps:
# - pandoc: DOCX export
# - build-essential + gcc: some python wheels may compile on slim
# - libsndfile1: soundfile runtime dependency
# - libcairo2: for cairosvg (LaTeX rendering)
RUN apt-get update && apt-get install -y --no-install-recommends \
    git \
    pandoc \
    build-essential \
    gcc \
    libsndfile1 \
    libcairo2 \
    && rm -rf /var/lib/apt/lists/*

# Create non-root user (HF recommendation)
RUN useradd -m -u 1000 user
USER user
ENV PATH="/home/user/.local/bin:$PATH"

WORKDIR /app

COPY --chown=user ./requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir --upgrade -r /app/requirements.txt

COPY --chown=user . /app

EXPOSE 7860

# Run Gradio in HF Docker (must bind 0.0.0.0:7860)
CMD ["python", "app.py"]