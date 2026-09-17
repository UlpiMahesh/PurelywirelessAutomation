FROM python:3.10

WORKDIR /app

COPY . .

RUN pip install -r requirements.txt

RUN playwright install --with-deps

CMD ["sh", "-c", "streamlit run app.py --server.port=$PORT --server.address=0.0.0.0"]