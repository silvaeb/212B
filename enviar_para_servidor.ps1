# =============================================================
# enviar_para_servidor.ps1
# Execute no PowerShell do Windows (na pasta do projeto):
#   cd E:\212B
#   .\enviar_para_servidor.ps1
# =============================================================

$SERVIDOR_IP   = "10.166.67.220"
$SERVIDOR_USER = "suporte"
$PROJETO_LOCAL = "E:\212B"
$DESTINO       = "/home/suporte/app_212b"
$ZIP_TEMP      = "$env:TEMP\212b_deploy.tar.gz"

Write-Host "======================================================"
Write-Host " Envio do Sistema 212B para o Servidor"
Write-Host "======================================================"

# ----------------------------------------------------------
# 1. Compactar o projeto (exclui .venv, __pycache__, .git etc.)
# ----------------------------------------------------------
Write-Host "[1/3] Compactando projeto..."

$excluir = @(
    ".venv", "__pycache__", ".git", "*.pyc",
    "*.xlsx", "*.db", "*.tar.gz", "*.zip",
    "instance", ".vscode", "backup_*"
)

$arquivos = Get-ChildItem -Path $PROJETO_LOCAL -Recurse | Where-Object {
    $caminho = $_.FullName
    $excluido = $false
    foreach ($pat in $excluir) {
        if ($caminho -like "*\$pat*" -or $caminho -like "*/$pat*") {
            $excluido = $true
            break
        }
    }
    -not $excluido -and -not $_.PSIsContainer
}

Write-Host "   Total de arquivos: $($arquivos.Count)"

# Usar tar do Windows (disponivel no Windows 10+)
Push-Location $PROJETO_LOCAL
tar --exclude=".venv" `
    --exclude="__pycache__" `
    --exclude=".git" `
    --exclude="*.pyc" `
    --exclude="instance" `
    --exclude=".vscode" `
    --exclude="*.xlsx" `
    --exclude="*.db" `
    --exclude="backup_*" `
    -czf $ZIP_TEMP . 2>&1

Pop-Location

if (Test-Path $ZIP_TEMP) {
    $tamanho = (Get-Item $ZIP_TEMP).Length / 1MB
    Write-Host "   Arquivo criado: $ZIP_TEMP ($([math]::Round($tamanho,1)) MB)"
} else {
    Write-Host "ERRO: Falha ao criar arquivo compactado." -ForegroundColor Red
    exit 1
}

# ----------------------------------------------------------
# 2. Criar estrutura no servidor e enviar arquivo
# ----------------------------------------------------------
Write-Host "[2/3] Enviando para $SERVIDOR_USER@$SERVIDOR_IP ..."
Write-Host "      (sera solicitada sua senha SSH)"

# Criar pasta no servidor
ssh "${SERVIDOR_USER}@${SERVIDOR_IP}" "mkdir -p $DESTINO"

# Enviar o tar.gz
scp $ZIP_TEMP "${SERVIDOR_USER}@${SERVIDOR_IP}:/tmp/212b_deploy.tar.gz"

# Descompactar no servidor
ssh "${SERVIDOR_USER}@${SERVIDOR_IP}" @"
    cd $DESTINO
    tar -xzf /tmp/212b_deploy.tar.gz -C $DESTINO
    rm -f /tmp/212b_deploy.tar.gz
    echo 'Arquivos extraidos em $DESTINO'
    ls -la $DESTINO
"@

# ----------------------------------------------------------
# 3. Enviar o script de deploy e executar
# ----------------------------------------------------------
Write-Host "[3/3] Enviando e executando script de deploy..."

scp "${PROJETO_LOCAL}\deploy_servidor.sh" "${SERVIDOR_USER}@${SERVIDOR_IP}:/home/suporte/deploy_servidor.sh"

ssh "${SERVIDOR_USER}@${SERVIDOR_IP}" "chmod +x /home/suporte/deploy_servidor.sh && bash /home/suporte/deploy_servidor.sh"

Write-Host ""
Write-Host "======================================================"
Write-Host " Pronto! Acesse: http://$SERVIDOR_IP"
Write-Host "======================================================"
