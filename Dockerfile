FROM python:3.11

WORKDIR /app

COPY backend/ ./backend/

RUN pip install --no-cache-dir -r backend/requirements.txt

EXPOSE 5000

CMD ["python", "backend/app.py"]