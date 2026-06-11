FROM python:3.11-slim

WORKDIR /code

# Copy requirements and install dependencies
COPY ./requirements.txt /code/requirements.txt
RUN pip install --no-cache-dir --upgrade -r /code/requirements.txt

# Copy application files
COPY ./app.py /code/app.py
COPY ./templates /code/templates
COPY ./static /code/static

# Create temp directory for uploads
RUN mkdir -p /code/temp_uploads && chmod 777 /code/temp_uploads

# Expose port 7860 (Hugging Face default port)
EXPOSE 7860

# Run uvicorn server
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "7860"]
