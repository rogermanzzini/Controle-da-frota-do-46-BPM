# Guia de Publicação em Produção e Nuvem - Sistema 46º BPM

Este documento reúne todas as opções preparadas para colocar o sistema em produção e publicado na internet.

---

## 1. Publicação Imediata na Internet (Já em Execução)

O aplicativo já está rodando em produção pelo servidor WSGI **Waitress** e transmitido para a internet por túnel seguro HTTPS da **Cloudflare**.

- **Link Público Ativo**: [https://determination-pipes-positive-homeless.trycloudflare.com](https://determination-pipes-positive-homeless.trycloudflare.com)
- **Como Iniciar no Dia a Dia**:
  Basta dar dois cliques no arquivo:
  ```
  publicar_na_internet.bat
  ```
  Ele iniciará o servidor de produção e gerará um link HTTPS seguro para acessar do celular, tablet ou qualquer outro computador.

---

## 2. Publicação Permanente em Servidor de Nuvem 24/7 (Render.com)

O Render é gratuito e suporta deploy contínuo direto do GitHub.

### Arquivos já configurados no projeto:
- `Procfile`: Comando de inicialização Gunicorn para a nuvem.
- `render.yaml`: Configuração automática de ambiente, portas e variáveis.
- `wsgi.py`: Ponto de entrada padrão WSGI.
- `requirements.txt`: Todas as dependências (incluindo `gunicorn`).

### Passo a passo para subir no Render:
1. Crie uma conta gratuita em [render.com](https://render.com).
2. Suba a pasta deste projeto em um repositório no seu GitHub (público ou privado).
3. No painel do Render, clique em **New +** > **Web Service**.
4. Conecte com o seu repositório do GitHub.
5. O Render detectará automaticamente o `Procfile` e o `requirements.txt`.
6. Defina:
   - **Environment**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn --workers=4 --bind=0.0.0.0:$PORT wsgi:application`
7. Clique em **Create Web Service**. Em 2 a 3 minutos, seu sistema terá uma URL permanente como `https://bpm46-frota.onrender.com`.

---

## 3. Publicação em Nuvem via Docker (Railway / Fly.io / VPS)

Também foi gerado um `Dockerfile` profissional pronto para rodar em qualquer hospedagem Docker:

```dockerfile
# Para testar ou subir no Railway / Fly.io:
docker build -t bpm46-frota .
docker run -p 5000:5000 bpm46-frota
```

---

## 4. Segurança e Produção

- **Sessões e Chaves**: A chave `SECRET_KEY` agora pode ser definida por variável de ambiente para total segurança.
- **Proxy Reverso**: Adicionado `ProxyFix` do Werkzeug para compatibilidade total com HTTPS e cabeçalhos de IPs de clientes.
- **Servidor WSGI**: O modo desenvolvimento (`app.run(debug=True)`) foi substituído por servidores WSGI multithread de alta performance (`waitress` no Windows e `gunicorn` no Linux/Nuvem).
