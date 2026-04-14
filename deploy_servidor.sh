#!/bin/bash
# =============================================================
# deploy_servidor.sh  -  Deploy automatico do sistema 212B
# Execute este script no servidor como usuario "suporte":
#   bash /home/suporte/deploy_servidor.sh
# =============================================================

set -e

APP_DIR="/home/suporte/app_212b"
VENV_DIR="$APP_DIR/.venv"
SECRET_KEY=$(python3 -c "import secrets; print(secrets.token_hex(32))")
DB_PATH="$APP_DIR/instance/sistema_pdrlog.db"
INSTANCE_DIR="$APP_DIR/instance"
SERVICE_NAME="app_212b"
PORT=8001   # porta interna; Nginx fara o proxy

echo "======================================================"
echo " Deploy do Sistema 212B"
echo "======================================================"

# ----------------------------------------------------------
# 1. Verificar se o projeto foi enviado para o servidor
# ----------------------------------------------------------
if [ ! -f "$APP_DIR/app.py" ]; then
    echo ""
    echo "ERRO: Pasta $APP_DIR nao encontrada ou app.py ausente."
    echo "Envie o projeto primeiro (veja instrucoes no final)."
    exit 1
fi

echo "[1/8] Projeto encontrado em $APP_DIR"

# ----------------------------------------------------------
# 2. Criar ambiente virtual Python
# ----------------------------------------------------------
echo "[2/8] Criando ambiente virtual Python..."
cd "$APP_DIR"
python3 -m venv "$VENV_DIR"
source "$VENV_DIR/bin/activate"
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt
echo "      Dependencias instaladas."

# ----------------------------------------------------------
# 3. Criar pasta instance e banco SQLite
# ----------------------------------------------------------
echo "[3/8] Criando pasta instance..."
mkdir -p "$INSTANCE_DIR"

# Copiar banco existente se houver na pasta antiga
if [ -f "/home/suporte/op_app_dev/instance/sistema_pdrlog.db" ]; then
    echo "      Banco antigo encontrado! Copiando..."
    cp "/home/suporte/op_app_dev/instance/sistema_pdrlog.db" "$DB_PATH"
else
    echo "      Nenhum banco antigo detectado. Sera criado novo."
fi

# ----------------------------------------------------------
# 4. Criar arquivo .env de producao
# ----------------------------------------------------------
echo "[4/8] Criando .env de producao..."
cat > "$APP_DIR/.env" <<EOF
APP_ENV=production
FLASK_ENV=production
SECRET_KEY=$SECRET_KEY
DATABASE_URL=sqlite:///$DB_PATH
EXCEL_PATH=Extra PDRLOG.xlsx
EOF
echo "      .env criado com SECRET_KEY gerada automaticamente."

# ----------------------------------------------------------
# 5. Inicializar banco de dados
# ----------------------------------------------------------
echo "[5/8] Inicializando banco de dados..."
cd "$APP_DIR"
source "$VENV_DIR/bin/activate"
python3 - <<'PYEOF'
import os
os.chdir("/home/suporte/app_212b")
from dotenv import load_dotenv
load_dotenv()
from app import app, init_database
init_database()
print("      Banco inicializado com sucesso.")
PYEOF

# ----------------------------------------------------------
# 6. Testar se o app sobe corretamente
# ----------------------------------------------------------
echo "[6/8] Testando importacao do app..."
python3 -c "
import os
os.chdir('/home/suporte/app_212b')
from dotenv import load_dotenv; load_dotenv()
from wsgi import application
print('      App importado OK:', application)
"

# ----------------------------------------------------------
# 7. Criar servico systemd
# ----------------------------------------------------------
echo "[7/8] Configurando servico systemd..."
SERVICE_FILE="/tmp/${SERVICE_NAME}.service"
cat > "$SERVICE_FILE" <<EOF
[Unit]
Description=Sistema 212B - Flask via Waitress
After=network.target

[Service]
User=suporte
Group=suporte
WorkingDirectory=$APP_DIR
EnvironmentFile=$APP_DIR/.env
ExecStart=$VENV_DIR/bin/waitress-serve --listen=127.0.0.1:$PORT wsgi:application
Restart=always
RestartSec=5
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
EOF

if command -v sudo &>/dev/null && sudo -n true 2>/dev/null; then
    sudo mv "$SERVICE_FILE" "/etc/systemd/system/${SERVICE_NAME}.service"
    sudo systemctl daemon-reload
    sudo systemctl enable "$SERVICE_NAME"
    sudo systemctl restart "$SERVICE_NAME"
    sleep 2
    sudo systemctl status "$SERVICE_NAME" --no-pager
    echo "      Servico systemd configurado e iniciado."
else
    echo ""
    echo "  AVISO: sudo nao disponivel sem senha."
    echo "  Execute manualmente (precisa de sudo):"
    echo ""
    echo "    sudo mv $SERVICE_FILE /etc/systemd/system/${SERVICE_NAME}.service"
    echo "    sudo systemctl daemon-reload"
    echo "    sudo systemctl enable $SERVICE_NAME"
    echo "    sudo systemctl restart $SERVICE_NAME"
    echo "    sudo systemctl status $SERVICE_NAME"
fi

# ----------------------------------------------------------
# 8. Gerar config Nginx e ativar
# ----------------------------------------------------------
echo "[8/8] Gerando configuracao Nginx..."
NGINX_CONF="/tmp/212b_nginx.conf"
cat > "$NGINX_CONF" <<EOF
server {
    listen 80;
    server_name _;

    client_max_body_size 50M;

    location /static/ {
        alias $APP_DIR/static/;
        expires 7d;
        add_header Cache-Control "public";
    }

    location / {
        proxy_pass http://127.0.0.1:$PORT;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
        proxy_read_timeout 300;
    }
}
EOF

if command -v sudo &>/dev/null && sudo -n true 2>/dev/null; then
    sudo cp "$NGINX_CONF" /etc/nginx/sites-available/212b
    sudo ln -sf /etc/nginx/sites-available/212b /etc/nginx/sites-enabled/212b
    sudo nginx -t && sudo systemctl reload nginx
    echo "      Nginx configurado e recarregado."
else
    echo ""
    echo "  AVISO: sudo nao disponivel sem senha."
    echo "  A configuracao do Nginx foi salva em: $NGINX_CONF"
    echo "  Execute manualmente:"
    echo ""
    echo "    sudo cp $NGINX_CONF /etc/nginx/sites-available/212b"
    echo "    sudo ln -sf /etc/nginx/sites-available/212b /etc/nginx/sites-enabled/212b"
    echo "    sudo nginx -t && sudo systemctl reload nginx"
fi

echo ""
echo "======================================================"
echo " Deploy concluido!"
echo ""
echo " Teste local:  curl http://127.0.0.1:$PORT"
echo " Teste via web: http://10.166.67.220"
echo ""
echo " Logs do servico: journalctl -u $SERVICE_NAME -f"
echo "======================================================"
