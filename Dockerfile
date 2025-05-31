# 1. Use uma imagem Python pública e estável.
#    Use a versão do Python que você está usando (3.11 é uma boa escolha, como no seu antigo).
FROM python:3.11-slim-buster

# 2. Define o diretório de trabalho dentro do contêiner.
#    É o lugar onde sua aplicação será copiada e executada.
WORKDIR /app

# 3. Copia o arquivo requirements.txt e instala as dependências.
#    O Render (e Docker) automaticamente buscará este arquivo na raiz do seu repositório.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 4. Copia o código da sua aplicação Flask (app.py) e a pasta 'templates'.
#    Certifique-se que 'app.py' e a pasta 'templates/' estão na raiz do seu repositório GitHub.
COPY app.py .
COPY templates/ templates/

# EXCLUÍDO:
# - As linhas 'COPY extrator_dgt', 'COPY extrator_dgt/settings_base.py'
# - A instalação do Oracle Instant Client (RUN apt-get update ... fakeroot alien ...)
# - O comando 'chmod +x /extrator_dgt/servicos.sh'
# Tudo isso era específico do ambiente do tribunal e do Oracle, e não é mais necessário.

# 5. Informa qual porta a aplicação irá escutar.
#    É mais para documentação do Docker. O Render injeta a variável $PORT.
EXPOSE 8000

# 6. Comando para iniciar a aplicação Flask com Gunicorn.
#    Este comando será executado quando o contêiner iniciar.
#    O Render injeta a variável de ambiente $PORT, então a aplicação deve escutar nela.
CMD ["gunicorn", "-b", "0.0.0.0:${PORT}", "app:app"]
# Explicação de "app:app":
# - O primeiro "app" se refere ao nome do seu arquivo Python principal (app.py).
# - O segundo "app" se refere à instância do seu aplicativo Flask dentro desse arquivo (app = Flask(__name__)).
# Se o seu arquivo principal for outro (ex: main.py) ou a instância do Flask tiver outro nome (ex: meu_app),
# ajuste "app:app" para "nome_do_arquivo:nome_da_instancia" (ex: "main:meu_app").
