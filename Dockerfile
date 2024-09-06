FROM python:3.10

COPY . ./

RUN pip install .

EXPOSE 8080

CMD exec gunicorn --bind :8080 --workers 1 --threads 8 --timeout 0 src.app.main:app
