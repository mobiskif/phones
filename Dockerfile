FROM python:3.13-alpine

WORKDIR /app

COPY index.html structure.xlsx pass.txt server.py ./

EXPOSE 8080

CMD ["python", "server.py"]
