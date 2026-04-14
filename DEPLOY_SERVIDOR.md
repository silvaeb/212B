# Deploy no servidor Ubuntu com Nginx

Este projeto e um sistema Flask. O servidor que voce mostrou ja tem Nginx ativo, entao a forma mais simples e segura e:

1. subir o codigo para uma pasta nova em /home/suporte/212B
2. criar uma venv Python nessa pasta
3. usar Waitress ouvindo em 127.0.0.1:8000
4. configurar Nginx como proxy reverso
5. registrar um servico systemd para reinicio automatico

## 1. Copiar o projeto para o servidor

No Windows, dentro da pasta do projeto:

```powershell
scp -r E:\212B suporte@VM-7CTA-CSUP-INTRANET-HOMOLOGACAO:/home/suporte/
```

Se preferir manter um nome sem espacos e sem risco de confusao:

```bash
mv /home/suporte/212B /home/suporte/app_212b
```

## 2. Criar ambiente Python no servidor

```bash
cd /home/suporte/app_212b
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

## 3. Criar o arquivo .env

Exemplo:

```bash
cat > /home/suporte/app_212b/.env <<'EOF'
APP_ENV=production
FLASK_ENV=production
SECRET_KEY=troque-esta-chave-por-uma-forte
EXCEL_PATH=Extra PDRLOG.xlsx
# Opcional: para SQLite em caminho absoluto
DATABASE_URL=sqlite:////home/suporte/app_212b/instance/sistema_pdrlog.db
EOF
```

Se voce usar SQLite, garanta que a pasta instance exista:

```bash
mkdir -p /home/suporte/app_212b/instance
```

## 4. Testar a aplicacao manualmente

```bash
cd /home/suporte/app_212b
source .venv/bin/activate
python -c "from wsgi import application; print(application)"
waitress-serve --listen=127.0.0.1:8000 wsgi:application
```

Em outro terminal no servidor:

```bash
curl http://127.0.0.1:8000
```

## 5. Criar servico systemd

Crie o arquivo /etc/systemd/system/212b.service com este conteudo:

```ini
[Unit]
Description=Sistema 212B Flask via Waitress
After=network.target

[Service]
User=suporte
Group=suporte
WorkingDirectory=/home/suporte/app_212b
EnvironmentFile=/home/suporte/app_212b/.env
ExecStart=/home/suporte/app_212b/.venv/bin/waitress-serve --listen=127.0.0.1:8000 wsgi:application
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

Depois:

```bash
sudo systemctl daemon-reload
sudo systemctl enable 212b
sudo systemctl start 212b
sudo systemctl status 212b
```

## 6. Configurar Nginx

Crie /etc/nginx/sites-available/212b com este conteudo:

```nginx
server {
    listen 80;
    server_name _;

    client_max_body_size 50M;

    location /static/ {
        alias /home/suporte/app_212b/static/;
        expires 7d;
        add_header Cache-Control "public";
    }

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_read_timeout 300;
    }
}
```

Ativar:

```bash
sudo ln -s /etc/nginx/sites-available/212b /etc/nginx/sites-enabled/212b
sudo nginx -t
sudo systemctl reload nginx
```

## 7. Backup do banco atual

Antes de trocar o sistema antigo, descubra onde esta o banco real. No seu log, o arquivo nao estava em /home/suporte/venv_opapp/instance. Procure assim:

```bash
find /home/suporte -name "*.db" 2>/dev/null
find /home/suporte/op_app_dev -maxdepth 3 -type f | grep -Ei "db|sqlite|instance"
```

Se o sistema antigo usar MariaDB, faca dump do banco certo:

```bash
mysqldump -u USUARIO -p NOME_DO_BANCO > backup_antigo.sql
```

## 8. Comandos de verificacao

```bash
systemctl status 212b
journalctl -u 212b -n 100 --no-pager
sudo nginx -t
curl -I http://127.0.0.1
```

## Observacoes importantes

- O seu usuario suporte esta no grupo sudo, mas o sudo so vai funcionar se voce souber a senha desse usuario.
- O pendrive nao apareceu em lsblk; entao o servidor nao detectou dispositivo USB conectado.
- Nao substitua o sistema antigo sem primeiro identificar o banco em uso e testar o novo sistema em outra pasta.
- Se quiser manter o sistema antigo no ar, publique o novo em /home/suporte/app_212b e teste por uma porta interna antes de apontar o Nginx.
