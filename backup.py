# backup.py
import os
import shutil
from datetime import datetime

def fazer_backup():
    """Faz backup do banco de dados"""
    if os.path.exists('sistema_pdrlog.db'):
        # Criar pasta de backup se não existir
        if not os.path.exists('backups'):
            os.makedirs('backups')
        
        # Nome do arquivo de backup com timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_file = f'backups/sistema_pdrlog_backup_{timestamp}.db'
        
        # Copiar arquivo do banco
        shutil.copy2('sistema_pdrlog.db', backup_file)
        print(f"✅ Backup criado: {backup_file}")
    else:
        print("❌ Arquivo do banco não encontrado para backup")

def restaurar_backup(arquivo_backup):
    """Restaura um backup do banco de dados"""
    if os.path.exists(arquivo_backup):
        shutil.copy2(arquivo_backup, 'sistema_pdrlog.db')
        print(f"✅ Backup restaurado: {arquivo_backup}")
    else:
        print(f"❌ Arquivo de backup não encontrado: {arquivo_backup}")

if __name__ == '__main__':
    fazer_backup()