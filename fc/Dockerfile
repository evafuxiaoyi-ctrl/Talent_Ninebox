FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    APP_ACCESS_PASSWORD=change-me \
    APP_SESSION_SECRET=change-me-session-secret

WORKDIR /code

COPY requirements.txt /code/requirements.txt
RUN pip install --no-cache-dir -r /code/requirements.txt

COPY talent_ninebox /code/talent_ninebox

EXPOSE 9000

CMD ["python", "-m", "uvicorn", "talent_ninebox.web.app:app", "--host", "0.0.0.0", "--port", "9000"]
