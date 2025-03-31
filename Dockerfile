# Use a imagem Python 3.10.12 oficial como base
FROM python:3.12

# Instala o Git (necessário para clonar o repositório)
RUN apt-get update && apt-get install -y git

# Define o diretório de trabalho dentro do container
WORKDIR /app

# Clona o repositório diretamente do GitHub
RUN git clone https://github.com/renatoeco/docs-generator-web.git /

# Instala as dependências do seu script (caso haja)
RUN pip install --no-cache-dir -r requirements.txt

# Expõe a porta padrão do Streamlit (8501)
EXPOSE 8501

# Comando para rodar a aplicação
CMD ["streamlit", "run", "main.py", "--server.port=8501", "--server.address=0.0.0.0"]

