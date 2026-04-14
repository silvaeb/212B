from datetime import datetime
from flask_login import UserMixin
from sqlalchemy.orm import validates
from werkzeug.security import generate_password_hash, check_password_hash
from database import db


class Usuario(db.Model, UserMixin):
    __tablename__ = 'usuario'

    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(100), unique=True, nullable=False)
    senha = db.Column(db.String(100), nullable=False)
    departamento = db.Column(db.String(100))
    nivel_acesso = db.Column(db.String(20), default='usuario')
    data_criacao = db.Column(db.DateTime, default=datetime.now)

    solicitacoes = db.relationship('SolicitacaoExtraPDRLOG', backref='usuario', lazy=True, cascade='all, delete-orphan')
    ncs = db.relationship('NotaCredito', backref='usuario', lazy=True, cascade='all, delete-orphan')
    pros = db.relationship('Pro', backref='usuario', lazy=True, cascade='all, delete-orphan')

    def set_password(self, raw_password: str) -> None:
        self.senha = generate_password_hash(raw_password) if raw_password else None

    def check_password(self, raw_password: str) -> bool:
        if not self.senha or raw_password is None:
            return False
        # Handle Werkzeug-hashed passwords (pbkdf2, scrypt, etc.); fall back to legacy plain text
        if self.senha.startswith(('pbkdf2:', 'scrypt:', 'bcrypt:', 'argon2:')):
            return check_password_hash(self.senha, raw_password)
        return self.senha == raw_password


class SolicitacaoExtraPDRLOG(db.Model):
    __tablename__ = 'solicitacao_extra_pdrlog'

    id = db.Column(db.Integer, primary_key=True)
    numero = db.Column(db.String(50), unique=True, nullable=False)
    diex = db.Column(db.String(100), nullable=False)
    orgao_demandante = db.Column(db.String(100))
    data_solicitacao = db.Column(db.DateTime, default=datetime.now)
    descricao = db.Column(db.Text)
    parecer_analise = db.Column(db.Text)
    despacho = db.Column(db.Text)
    status = db.Column(db.String(50), default='Aguardando Análise')
    finalidade = db.Column(db.String(100))
    modalidade = db.Column(db.String(20))
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    data_criacao = db.Column(db.DateTime, default=datetime.now)
    data_atualizacao = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    tem_fsv = db.Column(db.String(10))
    diex_dom = db.Column(db.String(100))
    nr_opus = db.Column(db.String(100))
    destinatario_colog = db.Column(db.String(50))

    pedidos = db.relationship('PedidoSolicitacao', backref='solicitacao', lazy=True, cascade='all, delete-orphan')
    notas_credito = db.relationship('NotaCredito', backref='solicitacao_origem', lazy=True, cascade='all, delete-orphan')
    pros_relacionadas = db.relationship('Pro', backref='solicitacao_origem', lazy=True, cascade='all, delete-orphan')

    @property
    def total_solicitado(self):
        total = 0
        for pedido in self.pedidos:
            for item in pedido.itens:
                total += (item.valor_solicitado or 0)
        return total

    @property
    def total_aprovado(self):
        total = 0
        for pedido in self.pedidos:
            for item in pedido.itens:
                total += (item.valor_aprovado or 0)
        return total


class PedidoSolicitacao(db.Model):
    __tablename__ = 'pedido_solicitacao'

    id = db.Column(db.Integer, primary_key=True)
    solicitacao_id = db.Column(db.Integer, db.ForeignKey('solicitacao_extra_pdrlog.id', ondelete='CASCADE'), nullable=False)
    om = db.Column(db.String(200), nullable=False)
    codom = db.Column(db.String(20))
    codug = db.Column(db.String(20))
    sigla_ug = db.Column(db.String(100))
    descricao_om = db.Column(db.Text)

    itens = db.relationship('ItemPedido', backref='pedido', lazy=True, cascade='all, delete-orphan')


class ItemPedido(db.Model):
    __tablename__ = 'item_pedido'

    id = db.Column(db.Integer, primary_key=True)
    pedido_id = db.Column(db.Integer, db.ForeignKey('pedido_solicitacao.id', ondelete='CASCADE'), nullable=False)
    nd = db.Column(db.String(20), nullable=False)
    finalidade = db.Column(db.String(100))
    pi = db.Column(db.String(50))
    valor_solicitado = db.Column(db.Float, nullable=False)
    valor_aprovado = db.Column(db.Float, default=0)
    valor_restante = db.Column(db.Float)

    @validates('pi')
    def normalizar_pi(self, key, value):
        if isinstance(value, list):
            if not value:
                return ''
            return str(value[0]).strip().upper()
        if value is None:
            return ''
        return str(value).strip().upper()


class NotaCredito(db.Model):
    __tablename__ = 'nota_credito'

    id = db.Column(db.Integer, primary_key=True)
    numero = db.Column(db.String(50), unique=True, nullable=False)

    cod_ug = db.Column(db.String(20))
    sigla_ug = db.Column(db.String(100))
    pi = db.Column(db.String(50))
    nd = db.Column(db.String(20))
    valor = db.Column(db.Float, default=0)
    ref_sisnc = db.Column(db.String(100))
    nc_siafi = db.Column(db.String(100))
    diex_credito = db.Column(db.String(100))
    descricao = db.Column(db.Text)

    status = db.Column(db.String(50), default='Pendente')
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    data_criacao = db.Column(db.DateTime, default=datetime.now)
    data_atualizacao = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    motivo_cancelamento = db.Column(db.Text)

    solicitacao_pdrlog_id = db.Column(db.Integer, db.ForeignKey('solicitacao_extra_pdrlog.id', ondelete='SET NULL'))
    pro_id = db.Column(db.Integer, db.ForeignKey('pro.id', ondelete='SET NULL'))
    item_id = db.Column(db.Integer, db.ForeignKey('item_pedido.id', ondelete='SET NULL'))



class Pro(db.Model):
    __tablename__ = 'pro'

    id = db.Column(db.Integer, primary_key=True)
    numero = db.Column(db.String(50), unique=True, nullable=False)

    cod_ug = db.Column(db.String(20))
    sigla_ug = db.Column(db.String(100))
    descricao = db.Column(db.Text)

    status = db.Column(db.String(50), default='Pendente')
    valor_total = db.Column(db.Float, default=0)
    valor_restante = db.Column(db.Float, default=0)

    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    solicitacao_pdrlog_id = db.Column(db.Integer, db.ForeignKey('solicitacao_extra_pdrlog.id', ondelete='SET NULL'))

    data_criacao = db.Column(db.DateTime, default=datetime.now)
    data_atualizacao = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    notas_credito = db.relationship('NotaCredito', backref='pro_origem', lazy=True, cascade='all, delete-orphan')

    @property
    def valor(self):
        return self.valor_total

    @valor.setter
    def valor(self, value):
        self.valor_total = value

    @property
    def vencida(self):
        # Considera vencida se status for aguardando licitação, parcialmente convertida ou pendente, e passou de 120 dias
        if self.status in ['Aguardando término do processo licitatório', 'Parcialmente Convertida', 'Pendente']:
            if self.data_criacao and (datetime.now() - self.data_criacao).days > 120:
                return True
        return False

    def marcar_vencida(self):
        self.status = 'Vencida (mais de 120 dias)'
        self.data_atualizacao = datetime.now()

    def prorrogar(self):
        self.status = 'Aguardando término do processo licitatório'
        self.data_atualizacao = datetime.now()

    id = db.Column(db.Integer, primary_key=True)
    numero = db.Column(db.String(50), unique=True, nullable=False)

    cod_ug = db.Column(db.String(20))
    sigla_ug = db.Column(db.String(100))
    descricao = db.Column(db.Text)

    status = db.Column(db.String(50), default='Pendente')
    valor_total = db.Column(db.Float, default=0)
    valor_restante = db.Column(db.Float, default=0)

    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    solicitacao_pdrlog_id = db.Column(db.Integer, db.ForeignKey('solicitacao_extra_pdrlog.id', ondelete='SET NULL'))

    data_criacao = db.Column(db.DateTime, default=datetime.now)
    data_atualizacao = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    notas_credito = db.relationship('NotaCredito', backref='pro_origem', lazy=True, cascade='all, delete-orphan')

    @property
    def saldo_restante(self):
        if self.valor_restante is not None:
            return self.valor_restante
        if self.valor_total is not None:
            total_nc = sum(nc.valor for nc in self.notas_credito if nc.status != 'Cancelada')
            return self.valor_total - total_nc
        return 0
