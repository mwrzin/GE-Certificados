# Publicar o Gerador de Certificados para o GEPetro

O app principal para publicar é `certificados_web.py`.

## Opção Recomendada: Render

1. Crie um repositório no GitHub, por exemplo `gerador-certificados-gepetro`.
2. Envie estes arquivos para o repositório:
   - `certificados_web.py`
   - `certificados_gui.py`
   - `requirements_certificados.txt`
   - `Procfile`
   - `Certificado_Gabrielly_Mayara.pdf`, se quiser deixar o modelo atual carregando automaticamente
3. No Render, crie um **New Web Service** conectado a esse repositório.
4. Configure:
   - **Build Command:** `pip install -r requirements_certificados.txt`
   - **Start Command:** `sh -c 'gunicorn certificados_web:app --bind 0.0.0.0:${PORT:-5000} --workers 2 --timeout 180'`
5. Em **Environment Variables**, adicione:
   - `CERT_WEB_USER` = `gepetro`
   - `CERT_WEB_PASSWORD` = uma senha compartilhada forte
   - Health check path, se o serviço pedir: `/health`
6. Depois do deploy, o Render vai fornecer uma URL parecida com:
   - `https://gerador-certificados-gepetro.onrender.com`

Compartilhe essa URL, o usuário `gepetro` e a senha apenas com quem deve acessar.

## Segurança

- Não coloque senha de app do Gmail dentro do código.
- A senha de app/key deve ser digitada apenas na tela do app quando for enviar e-mails.
- Use uma senha forte em `CERT_WEB_PASSWORD`.
- Evite publicar o repositório com planilhas reais, CPFs, certificados finais ou logs.
- Para produção real com muitos usuários, prefira adicionar login individual e armazenamento persistente.

## Teste Local

```bash
source .venv_certificados/bin/activate
python certificados_web.py
```

Abra:

```text
http://127.0.0.1:5000
```

Para testar com senha local:

```bash
CERT_WEB_USER=gepetro CERT_WEB_PASSWORD=senha-teste python certificados_web.py
```
