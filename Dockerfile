# Usa a imagem oficial do Python como base
FROM python:3.12

# Instala o Git (necessário para clonar o repositório)
RUN apt-get update && apt-get install -y git && rm -rf /var/lib/apt/lists/*

# Define o diretório de trabalho dentro do container
WORKDIR /app

# Clona o repositório diretamente do GitHub
RUN git clone https://github.com/renatoeco/Gerador_Docs_Web.git /app

# Garante que estamos na versão mais recente do código
WORKDIR /app
RUN git pull origin main

# Instala as dependências do projeto
RUN pip install --no-cache-dir -r /app/requirements.txt

# Expõe a porta padrão do Streamlit (8501)
EXPOSE 8501

# Comando para rodar a aplicação
CMD ["streamlit", "run", "main.py", "--server.port=8501", "--server.address=0.0.0.0"]
