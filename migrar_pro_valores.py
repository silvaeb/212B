
import sqlite3
# Caminho correto do banco de dados SQLite
DB_PATH = 'instance/sistema_pdrlog.db'

conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

try:
    cursor.execute("ALTER TABLE pro ADD COLUMN valor_total FLOAT DEFAULT 0;")
    print("Coluna 'valor_total' adicionada com sucesso.")
except Exception as e:
    print("Aviso:", e)

try:
    cursor.execute("ALTER TABLE pro ADD COLUMN valor_restante FLOAT DEFAULT 0;")
    print("Coluna 'valor_restante' adicionada com sucesso.")
except Exception as e:
    print("Aviso:", e)

conn.commit()
conn.close()
print("Migração concluída.")
