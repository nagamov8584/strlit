FROM python:3.11-slim

WORKDIR /Users/erebor/PycharmProjects/strlit_test

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

COPY main.py /app/main.py


EXPOSE 8501

CMD ["streamlit", "run", "main.py"]
