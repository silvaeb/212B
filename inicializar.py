# inicializar.py
import os
from app import app
from database import db
from models import Usuario

def criar_usuarios_iniciais():
    """Cria usuários iniciais"""
    if not Usuario.query.first():
        usuarios = [
            Usuario(nome='Administrador', email='admin@pdrlog.com', departamento='COLOG', nivel_acesso='admin'),
            Usuario(nome='João Silva', email='joao@empresa.com', departamento='Produção', nivel_acesso='supervisor'),
            Usuario(nome='Maria Santos', email='maria@empresa.com', departamento='COLOG', nivel_acesso='gerente'),
        ]
        
        # Atribui senhas com hash
        usuarios[0].set_password('admin123')
        usuarios[1].set_password('123456')
        usuarios[2].set_password('123456')

        for usuario in usuarios:
            db.session.add(usuario)
        
        db.session.commit()

def inicializar_banco_apenas_se_necessario():
    """Inicializa o banco apenas se não existir"""
    with app.app_context():
        try:
            # Verificar se o banco já existe
            if os.path.exists('sistema_pdrlog.db'):
                print("✅ Banco já existe. Nenhuma ação necessária.")
                return
            
            print("📝 Criando novo banco de dados...")
            db.create_all()
            criar_usuarios_iniciais()
            print("🎉 Banco criado com sucesso!")
            
        except Exception as e:
            print(f"❌ Erro ao criar banco: {e}")
            import traceback
            traceback.print_exc()

if __name__ == '__main__':
    inicializar_banco_apenas_se_necessario()