FROM python:3.7
COPY secrets secrets/
COPY swapify swapify/
COPY requirements.txt requirements.txt
RUN pip install --upgrade pip && \
    pip install -r requirements.txt
CMD ["python", "swapify/swapify.py"]