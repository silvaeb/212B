from flask import Flask, render_template, request, jsonify, redirect, url_for, flash, send_file, make_response, session, current_app
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from database import db
from models import Usuario, SolicitacaoExtraPDRLOG, PedidoSolicitacao, ItemPedido, NotaCredito, Pro
from datetime import datetime, timedelta
import csv
import json
from io import StringIO, BytesIO
import os
import sys
import math
import pandas as pd
import unicodedata
import re
from functools import wraps
from dotenv import load_dotenv
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

load_dotenv()

from config import get_config_class

app = Flask(__name__)
app.config.from_object(get_config_class())

# ==== ROTA UTILITÁRIA: ATUALIZAR STATUS EM LOTE ==== (deve vir após definição do app)
def atualizar_status_ncs():
    """Atualiza o status de todas as Notas de Crédito conforme os campos REF SISNC e NC SIAFI."""
    from models import NotaCredito
    ncs = NotaCredito.query.all()
    total_processada = 0
    total_sisnc = 0
    total_pendente = 0
    for nc in ncs:
        status_anterior = nc.status
        ref_sisnc = (nc.ref_sisnc or '').strip()
        nc_siafi = (nc.nc_siafi or '').strip()
        if ref_sisnc and nc_siafi:
            nc.status = 'Processada SIAFI'
            total_processada += 1
        elif ref_sisnc and not nc_siafi:
            nc.status = 'Cadastrada SISNC'
            total_sisnc += 1
        else:
            total_pendente += 1
        if status_anterior != nc.status:
            print(f"NC {nc.numero}: {status_anterior} -> {nc.status}")
    db.session.commit()
    flash(f'Status atualizado: {total_processada} "Processada SIAFI", {total_sisnc} "Cadastrada SISNC", {total_pendente} "Pendente/Outro".', 'success')
    return redirect(url_for('listar_ncs'))

app.add_url_rule('/atualizar_status_ncs', view_func=atualizar_status_ncs, methods=['POST'])



# ==== ROTA UTILITÁRIA: ATUALIZAR STATUS EM LOTE ==== (deve vir após definição do app)
# (deixe este bloco após a configuração do app e dos imports principais)
# COLOQUE ESTE BLOCO APÓS app = Flask(__name__) E SUAS CONFIGS

# ...existing code...

# --- INÍCIO DO BLOCO CORRETO ---

# (coloque este bloco após a linha app = Flask(__name__))


# ...existing code...


# --- FIM DO BLOCO CORRETO ---
def gerar_chave_correspondencia(codom, codug, pi, nd, valor):
    """Gera chave de correspondência padronizada para NC SIAFI: CODOM, CODUG, PI, ND, Valor."""
    codom_norm = normalizar_codom(codom)
    codug_norm = normalizar_cod_ug(codug)
    pi_norm = str(pi or '').strip().upper()
    nd_norm = str(nd or '').strip().upper()
    # Valor pode ser float ou string; usar normalização robusta
    try:
        valor_norm = f"{normalizar_valor(valor):.2f}"
    except Exception:
        try:
            valor_norm = f"{float(str(valor or '').replace(',', '.')):.2f}"
        except Exception:
            valor_norm = '0.00'
    return f"{codom_norm}|{codug_norm}|{pi_norm}|{nd_norm}|{valor_norm}"

def normalizar_codom(codom):
    """Remove pontos, espaços e zeros à esquerda do CODOM."""
    if codom is None:
        return ''
    # Remove tudo que não é dígito
    codom_str = re.sub(r'\D', '', str(codom))
    # Remove zeros à esquerda
    return codom_str.lstrip('0') or '0'


# ==== IMPORTS: devem estar no início do arquivo ====
from flask import Flask, render_template, request, jsonify, redirect, url_for, flash, send_file, make_response, session, current_app
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from database import db
from models import Usuario, SolicitacaoExtraPDRLOG, PedidoSolicitacao, ItemPedido, NotaCredito, Pro
from datetime import datetime, timedelta
import csv
import json
from io import StringIO, BytesIO
import os
import sys
import math
import pandas as pd
import unicodedata
import re
from functools import wraps
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from config import get_config_class


# ==== CONFIGURAÇÃO BÁSICA DO APP ==== (app definition é aqui)

app = Flask(__name__)
app.config.from_object(get_config_class())

# Status automaticamente derivado dos campos de referencia da NC.
def _status_automatico_nc(ref_sisnc='', nc_siafi='', status_atual=''):
    status_atual_txt = str(status_atual or '').strip()
    if status_atual_txt in ('Cancelada', 'Nao Conferida', 'Não Conferida'):
        return status_atual_txt

    ref_txt = str(ref_sisnc or '').strip()
    siafi_txt = str(nc_siafi or '').strip()
    if siafi_txt:
        return 'Processada SIAFI'
    if ref_txt:
        return 'Cadastrada SISNC'
    return 'Pendente'

# ==== ROTA UTILITÁRIA: ATUALIZAR STATUS EM LOTE ==== (deve vir após definição do app)
def atualizar_status_ncs():
    """Atualiza o status de todas as Notas de Crédito conforme os campos REF SISNC e NC SIAFI."""
    from models import NotaCredito
    ncs = NotaCredito.query.all()
    total_processada = 0
    total_sisnc = 0
    total_pendente = 0
    for nc in ncs:
        status_anterior = nc.status
        nc.status = _status_automatico_nc(
            ref_sisnc=nc.ref_sisnc,
            nc_siafi=nc.nc_siafi,
            status_atual=nc.status
        )
        if nc.status == 'Processada SIAFI':
            total_processada += 1
        elif nc.status == 'Cadastrada SISNC':
            total_sisnc += 1
        else:
            total_pendente += 1
        if status_anterior != nc.status:
            print(f"NC {nc.numero}: {status_anterior} -> {nc.status}")
    db.session.commit()
    flash(f'Status atualizado: {total_processada} "Processada SIAFI", {total_sisnc} "Cadastrada SISNC", {total_pendente} "Pendente/Outro".', 'success')
    return redirect(url_for('listar_ncs'))

app.add_url_rule('/atualizar_status_ncs', view_func=atualizar_status_ncs, methods=['POST'])

MENU_MODULO_SESSION_KEY = 'menu_modulo'
MENU_MODULO_PDR = 'pdr_log'
MENU_MODULO_EXTRA = 'extra_pdr_log'


def definir_modulo_menu(modulo):
    if modulo in (MENU_MODULO_PDR, MENU_MODULO_EXTRA):
        session[MENU_MODULO_SESSION_KEY] = modulo


def obter_modulo_menu():
    return session.get(MENU_MODULO_SESSION_KEY, MENU_MODULO_EXTRA)


@app.context_processor
def inject_menu_context():
    return {'menu_modulo': obter_modulo_menu()}

# ==== ROTA DASHBOARD DEVE VIR APÓS DEFINIÇÃO DO APP ====
@app.route('/dashboard')
@login_required
def dashboard():
    finalidade_filter = request.args.get('finalidade', '').strip()
    orgao_filter = request.args.get('orgao', '').strip()
    nd_filter = request.args.get('nd', '').strip()
    gnd_filter = request.args.get('gnd', '').strip()
    descricao_filter = request.args.get('descricao', '').strip()

    # Query base de solicitações
    solicitacao_query = SolicitacaoExtraPDRLOG.query
    if finalidade_filter:
        solicitacao_query = solicitacao_query.join(PedidoSolicitacao).join(ItemPedido).filter(ItemPedido.finalidade == finalidade_filter)
    if orgao_filter:
        solicitacao_query = solicitacao_query.filter(SolicitacaoExtraPDRLOG.orgao_demandante == orgao_filter)
    if descricao_filter:
        solicitacao_query = solicitacao_query.filter(SolicitacaoExtraPDRLOG.descricao.ilike(f'%{descricao_filter}%'))
    if nd_filter:
        solicitacao_query = solicitacao_query.join(PedidoSolicitacao).join(ItemPedido).filter(ItemPedido.nd == nd_filter)
    if gnd_filter:
        solicitacao_query = solicitacao_query.filter(SolicitacaoExtraPDRLOG.gnd == gnd_filter)

    # Subquery para reuso
    from sqlalchemy import func, select
    solicitacao_subq = solicitacao_query.with_entities(SolicitacaoExtraPDRLOG.id).subquery()

    def soma_solicitacoes(status_val=None):
        q = db.session.query(func.coalesce(func.sum(ItemPedido.valor_solicitado), 0))\
            .join(PedidoSolicitacao, ItemPedido.pedido_id == PedidoSolicitacao.id)\
            .join(SolicitacaoExtraPDRLOG, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)\
            .filter(SolicitacaoExtraPDRLOG.id.in_(select(solicitacao_subq)))
        if status_val:
            q = q.filter(SolicitacaoExtraPDRLOG.status == status_val)
        return q.scalar() or 0

    def conta_solicitacoes(status_val=None):
        q = solicitacao_query
        if status_val:
            q = q.filter(SolicitacaoExtraPDRLOG.status == status_val)
        return q.count()

    total_solicitacoes = conta_solicitacoes()
    total_solicitacoes_valor = soma_solicitacoes()

    solicitacoes_aguardando = conta_solicitacoes('Aguardando Análise')
    solicitacoes_aguardando_valor = soma_solicitacoes('Aguardando Análise')

    solicitacoes_analise = conta_solicitacoes('Em Análise')
    solicitacoes_analise_valor = soma_solicitacoes('Em Análise')

    solicitacoes_aprovadas = conta_solicitacoes('Aprovado Ch Sup')
    solicitacoes_aprovadas_valor = soma_solicitacoes('Aprovado Ch Sup')

    solicitacoes_negadas = conta_solicitacoes('Negado Ch Sup')
    solicitacoes_negadas_valor = soma_solicitacoes('Negado Ch Sup')

    solicitacoes_arquivadas = conta_solicitacoes('Arquivado')
    solicitacoes_arquivadas_valor = soma_solicitacoes('Arquivado')

    # NCs com os mesmos filtros (via solicitação origem)
    nc_query = NotaCredito.query
    if finalidade_filter or orgao_filter or descricao_filter:
        nc_query = nc_query.join(SolicitacaoExtraPDRLOG, NotaCredito.solicitacao_pdrlog_id == SolicitacaoExtraPDRLOG.id)
        if finalidade_filter:
            nc_query = (
                nc_query
                .join(PedidoSolicitacao, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
                .join(ItemPedido, ItemPedido.pedido_id == PedidoSolicitacao.id)
                .filter(ItemPedido.finalidade == finalidade_filter)
                .distinct()
            )
        if orgao_filter:
            nc_query = nc_query.filter(SolicitacaoExtraPDRLOG.orgao_demandante == orgao_filter)
        if descricao_filter:
            nc_query = nc_query.filter(SolicitacaoExtraPDRLOG.descricao.ilike(f'%{descricao_filter}%'))

    total_ncs = nc_query.count()
    total_ncs_valor = (nc_query.with_entities(func.coalesce(func.sum(NotaCredito.valor), 0)).scalar() or 0)

    ncs_pendentes = nc_query.filter(NotaCredito.status == 'Pendente').count()
    ncs_pendentes_valor = (nc_query.filter(NotaCredito.status == 'Pendente')
                                      .with_entities(func.coalesce(func.sum(NotaCredito.valor), 0)).scalar() or 0)

    # PROs com filtros na solicitação de origem
    pro_query = Pro.query
    if finalidade_filter or orgao_filter or descricao_filter:
        pro_query = pro_query.join(SolicitacaoExtraPDRLOG, Pro.solicitacao_pdrlog_id == SolicitacaoExtraPDRLOG.id)
        if finalidade_filter:
            pro_query = (
                pro_query
                .join(PedidoSolicitacao, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
                .join(ItemPedido, ItemPedido.pedido_id == PedidoSolicitacao.id)
                .filter(ItemPedido.finalidade == finalidade_filter)
                .distinct()
            )
        if orgao_filter:
            pro_query = pro_query.filter(SolicitacaoExtraPDRLOG.orgao_demandante == orgao_filter)
        if descricao_filter:
            pro_query = pro_query.filter(SolicitacaoExtraPDRLOG.descricao.ilike(f'%{descricao_filter}%'))

    total_pros = pro_query.count()
    # Valor de PRO baseado no total solicitado da solicitação de origem
    total_pros_valor = (db.session.query(func.coalesce(func.sum(ItemPedido.valor_solicitado), 0))
            .join(PedidoSolicitacao, ItemPedido.pedido_id == PedidoSolicitacao.id)
            .join(SolicitacaoExtraPDRLOG, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
            .join(Pro, Pro.solicitacao_pdrlog_id == SolicitacaoExtraPDRLOG.id)
            .filter(Pro.id.in_(pro_query.with_entities(Pro.id)))
            .scalar() or 0)

    pros_pendentes = pro_query.filter(Pro.status == 'Pendente').count()
    pros_pendentes_valor = (db.session.query(func.coalesce(func.sum(ItemPedido.valor_solicitado), 0))
            .join(PedidoSolicitacao, ItemPedido.pedido_id == PedidoSolicitacao.id)
            .join(SolicitacaoExtraPDRLOG, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
            .join(Pro, Pro.solicitacao_pdrlog_id == SolicitacaoExtraPDRLOG.id)
            .filter(Pro.id.in_(pro_query.filter(Pro.status == 'Pendente').with_entities(Pro.id)))
            .scalar() or 0)

    # Séries para gráficos (aplicam os mesmos filtros)
    # Somatório por finalidade (valor solicitado)
    FINALIDADES = [
        'Quantitativo de Rancho (QR)',
        'Quantitativo de Subsistência (QS)', 
        'Reserva Regional (RR)',
        'Alimentação em Campanha',
        'PASA',
        'Manutenção de Órgão Provedor',
        'Hub Logístico',
        'Remanejamento',
        'Câmaras frigoríficas',
        'PASA-DEC',
        'Mnt OP-DEC',
        'Publicações',
        'Diárias',
        'Passagens',
        'Solenidades',
        'Outros'
    ]
    soma_por_finalidade = {f: 0 for f in FINALIDADES}
    finais_rows = (
            db.session.query(
                ItemPedido.finalidade,
                func.coalesce(func.sum(ItemPedido.valor_solicitado), 0)
            )
            .join(PedidoSolicitacao, ItemPedido.pedido_id == PedidoSolicitacao.id)
            .join(SolicitacaoExtraPDRLOG, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
            .filter(SolicitacaoExtraPDRLOG.id.in_(select(solicitacao_subq.c.id)))
            .group_by(ItemPedido.finalidade)
            .all()
        )

    for fin, total_val in finais_rows:
        if fin in soma_por_finalidade:
            soma_por_finalidade[fin] = total_val or 0

    solicitacoes_por_finalidade = [soma_por_finalidade[f] for f in FINALIDADES]

    # Somatório por órgão demandante (valor solicitado)
    orgaos_labels = []
    orgaos_valores = []
    orgaos_rows = (
            db.session.query(
                SolicitacaoExtraPDRLOG.orgao_demandante,
                               func.coalesce(func.sum(ItemPedido.valor_solicitado), 0)
            )
            .join(PedidoSolicitacao, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
            .join(ItemPedido, ItemPedido.pedido_id == PedidoSolicitacao.id)
            .filter(SolicitacaoExtraPDRLOG.id.in_(select(solicitacao_subq.c.id)))
            .group_by(SolicitacaoExtraPDRLOG.orgao_demandante)
            .order_by(func.coalesce(func.sum(ItemPedido.valor_solicitado), 0).desc())
            .all()
        )

    for orgao, total_val in orgaos_rows:
        orgaos_labels.append(orgao or '---')
        orgaos_valores.append(total_val or 0)

    STATUS_SOLICITACAO = ['Aguardando Análise', 'Em Análise', 'Aguardando despacho', 'Aprovado Ch Sup', 'Aprovado Parcialmente', 'Negado Ch Sup', 'Arquivado', 'Devolvido para correções']
    solicitacoes_por_status = {}
    for status in STATUS_SOLICITACAO:
        count = solicitacao_query.filter(SolicitacaoExtraPDRLOG.status == status).count()
        solicitacoes_por_status[status] = count

    # Solicitações por mês (últimos 6 meses) por valor solicitado
    from datetime import datetime
    meses_labels = []
    meses_valores = []

    # Opções para filtros
    gnd_opcoes = sorted({getattr(s, 'gnd', None) for s in SolicitacaoExtraPDRLOG.query if getattr(s, 'gnd', None)})
    nd_possiveis = sorted({item.nd for item in ItemPedido.query if item.nd})

    # Lista de órgãos demandantes (opções estáticas)
    orgaos_demandantes = [
        '1ª RM', '2ª RM', '3ª RM', '4ª RM', '5ª RM', '6ª RM', '7ª RM', '8ª RM', '9ª RM', '10ª RM', '11ª RM', '12ª RM',
        'DGP', 'DEC', 'DECEx', 'SEF', 'DCT', 'COLOG', 'COTER', 'GAB CMT EX', 'EME', 'SGEx', 'CIE', 'CCOMSEx',
        'CMS', 'CML', 'CMO', 'CMSE', 'CMP', 'CMNE', 'CMN', 'CMA', '3º GPT LOG', '8º GPT LOG', '9º GPT LOG'
    ]

    return render_template(
        'dashboard.html',
        finalidades=FINALIDADES,
        orgaos_demandantes=orgaos_demandantes,
        nd_possiveis=nd_possiveis,
        gnd_opcoes=gnd_opcoes,
        finalidade_filter=finalidade_filter,
        orgao_filter=orgao_filter,
        nd_filter=nd_filter,
        gnd_filter=gnd_filter,
        descricao_filter=descricao_filter,
        solicitacoes_por_finalidade=solicitacoes_por_finalidade,
        solicitacoes_por_status=solicitacoes_por_status,
        solicitacoes_por_mes={'meses': meses_labels, 'valores': meses_valores},
        orgaos_labels=orgaos_labels,
        orgaos_valores=orgaos_valores,
        total_solicitacoes=total_solicitacoes,
        total_solicitacoes_valor=total_solicitacoes_valor,
        solicitacoes_aguardando=solicitacoes_aguardando,
        solicitacoes_aguardando_valor=solicitacoes_aguardando_valor,
        solicitacoes_analise=solicitacoes_analise,
        solicitacoes_analise_valor=solicitacoes_analise_valor,
        solicitacoes_aprovadas=solicitacoes_aprovadas,
        solicitacoes_aprovadas_valor=solicitacoes_aprovadas_valor,
        solicitacoes_negadas=solicitacoes_negadas,
        solicitacoes_negadas_valor=solicitacoes_negadas_valor,
        solicitacoes_arquivadas=solicitacoes_arquivadas,
        solicitacoes_arquivadas_valor=solicitacoes_arquivadas_valor,
        total_ncs=total_ncs,
        total_ncs_valor=total_ncs_valor,
        ncs_pendentes=ncs_pendentes,
        ncs_pendentes_valor=ncs_pendentes_valor,
        total_pros=total_pros,
        total_pros_valor=total_pros_valor,
        pros_pendentes=pros_pendentes,
        pros_pendentes_valor=pros_pendentes_valor
    )

# Usa caminho absoluto no instance/ quando estiver no SQLite padrao.
if not os.getenv('DATABASE_URL') and app.config.get('SQLALCHEMY_DATABASE_URI') == 'sqlite:///sistema_pdrlog.db':
    os.makedirs(app.instance_path, exist_ok=True)
    instance_db = os.path.join(app.instance_path, 'sistema_pdrlog.db').replace('\\', '/')
    app.config['SQLALCHEMY_DATABASE_URI'] = f"sqlite:///{instance_db}"

# Inicializa extensões
db.init_app(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# ...existing code...

# DECORATOR PARA CONTROLE DE ACESSO
def acesso_requerido(*niveis_permitidos):
    """Decorator para controlar acesso baseado no nível do usuário"""
    def decorator(f):
        @wraps(f)
        @login_required
        def decorated_function(*args, **kwargs):
            if current_user.nivel_acesso not in niveis_permitidos:
                if current_user.nivel_acesso == 'nc_only':
                    # Usuário NC Only é redirecionado para lista de NCs
                    flash('Acesso permitido apenas para Notas de Crédito.', 'warning')
                    return redirect(url_for('listar_ncs'))
                else:
                    flash('Acesso não autorizado para esta funcionalidade.', 'error')
                    return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return decorated_function
    return decorator

@app.route('/exportar_ncs_excel')
@acesso_requerido('admin', 'usuario')
def exportar_ncs_excel():
    # Get filters from request
    tipo = request.args.get('tipo', '') or request.form.get('tipo', '')
    finalidade = request.args.get('finalidade', '') or request.form.get('finalidade', '')
    status = request.args.get('status', '') or request.form.get('status', '')
    data_inicio = request.args.get('data_inicio', '') or request.form.get('data_inicio', '')
    data_fim = request.args.get('data_fim', '') or request.form.get('data_fim', '')

    df = gerar_relatorio_ncs_dataframe(tipo, finalidade, status, data_inicio, data_fim)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Notas de Crédito', index=False)
        worksheet = writer.sheets['Notas de Crédito']
        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).str.len().max(), len(col)) + 2
            worksheet.column_dimensions[chr(65 + idx)].width = max_len
    output.seek(0)
    return send_file(
                output,
                download_name=f'ncs_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
                as_attachment=True,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )


def gerar_relatorio_ncs_dataframe(tipo='', finalidade='', status='', data_inicio='', data_fim=''):
    query = NotaCredito.query
    if tipo:
        if tipo == 'Notas de Crédito':
            pass
        elif tipo == 'Solicitações':
            query = query.filter(NotaCredito.solicitacao_pdrlog_id.isnot(None))
        elif tipo == 'PRO':
            query = query.filter(NotaCredito.pro_id.isnot(None))
        elif tipo == 'Transferência Interna COLOG':
            query = query.filter(NotaCredito.modalidade == 'Transferência Interna COLOG')
    if finalidade:
        query = query.join(NotaCredito.solicitacao_origem).filter(SolicitacaoExtraPDRLOG.finalidade == finalidade)
    if status:
        query = query.filter(NotaCredito.status == status)
    if data_inicio:
        try:
            dt_inicio = datetime.strptime(data_inicio, '%Y-%m-%d')
            query = query.filter(NotaCredito.data_criacao >= dt_inicio)
        except Exception:
            pass
    if data_fim:
        try:
            dt_fim = datetime.strptime(data_fim, '%Y-%m-%d')
            query = query.filter(NotaCredito.data_criacao <= dt_fim)
        except Exception:
            pass

    ncs = query.order_by(NotaCredito.data_criacao.desc()).all()

    # Agrupamento por CODOM, CODUG, PI, ND (OM padronizada)
    agrupados = {}
    for nc in ncs:
        solicitacao = nc.solicitacao_origem if hasattr(nc, 'solicitacao_origem') else None
        om_nome = ''
        om_nome_raw = ''
        cod_ug = nc.cod_ug or ''
        sigla_ug = nc.sigla_ug or ''
        if nc.pro_id and nc.pro_origem:
            om_nome = nc.pro_origem.cod_ug or ''
            om_nome_raw = om_nome
        elif nc.solicitacao_pdrlog_id and solicitacao:
            for pedido in solicitacao.pedidos:
                om_nome = pedido.om or pedido.codom or ''
                om_nome_raw = pedido.om or pedido.codom or ''
                if om_nome:
                    break
        codom = ''
        if nc.solicitacao_pdrlog_id and solicitacao:
            for pedido in solicitacao.pedidos:
                codom = pedido.codom or ''
                if codom:
                    break
        codom_normalizado = normalizar_codom(codom)
        # Chave de agrupamento IGNORA diferenças de nomenclatura na OM (usa normalização)
        chave = (normalizar_om_solicitante_chave(om_nome), codom_normalizado, str(cod_ug).strip().upper(), str(nc.pi or '').strip().upper(), str(nc.nd or '').strip().upper())
        if chave not in agrupados:
            agrupados[chave] = {
                'OMs': [],
                'OMs_raw': [],
                'Valor NC': 0.0,
                'NCs': [],
                'PI': nc.pi or '',
                'COD OM': codom_normalizado,
                'COD UG': cod_ug,
                'Sigla UG': sigla_ug,
                'ND': nc.nd or '',
                'Status NC': nc.status or '',
            }
        agrupados[chave]['OMs'].append(normalizar_om_solicitante_chave(om_nome_raw))
        agrupados[chave]['OMs_raw'].append(om_nome_raw)
        agrupados[chave]['Valor NC'] += float(nc.valor or 0)
        agrupados[chave]['NCs'].append(nc.numero)

    # Adotar OM padronizada (mais frequente ou primeira, mas mostrar nome original)
    data = []
    for info in agrupados.values():
        if info['OMs_raw']:
            # Escolhe o nome original mais frequente
            nomes = [n for n in info['OMs_raw'] if n]
            om_padrao = max(set(nomes), key=nomes.count) if nomes else ''
        else:
            om_padrao = ''
        data.append({
            'OM Solicitante': om_padrao,
            'Valor NC': info['Valor NC'],
            'NCs': ', '.join(sorted(set(info['NCs']))),
            'PI': info['PI'],
            'COD OM': info['COD OM'],
            'COD UG': info['COD UG'],
            'Sigla UG': info['Sigla UG'],
            'ND': info['ND'],
            'Status NC': info['Status NC'],
        })
    return pd.DataFrame(data)


# ...existing code...


# ...existing code...


@app.template_filter('format_currency')
def format_currency(value):
    """Formata valores numéricos em R$ com separadores brasileiros."""
    try:
        valor = float(value)
        if pd.isna(valor) or math.isinf(valor):
            valor = 0.0
        return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return "R$ 0,00"


def normalizar_texto_simples(valor):
    if valor is None:
        return ''
    texto = str(valor).strip()
    return texto


def normalizar_cod_ug(valor):
    try:
        if valor is None or (isinstance(valor, float) and pd.isna(valor)):
            return ''
        s = str(valor).strip()
        s = re.sub(r'\D', '', s)
        return s
    except Exception:
        return ''


def normalizar_valor(v):
    """Converte formatos comuns de moeda para float.

    Exemplos suportados:
    - '1.234,56' -> 1234.56
    - '1234,56'  -> 1234.56
    - '1234.56'  -> 1234.56
    - 1234, 1234.56 -> float
    """
    try:
        if v is None:
            return 0.0
        # numerics
        if isinstance(v, (int, float)):
            try:
                if pd.isna(v) or math.isinf(float(v)):
                    return 0.0
            except Exception:
                return 0.0
            return float(v)

        texto = str(v).strip()
        if texto == '' or texto.lower() in ('nan', 'none', 'null'):
            return 0.0

        # tratar parênteses como negativo: (1.234,56) -> -1234.56
        negativo = False
        if texto.startswith('(') and texto.endswith(')'):
            negativo = True
            texto = texto[1:-1].strip()

        # sinal de negativo no final ou início
        if texto.endswith('-'):
            negativo = True
            texto = texto[:-1].strip()
        if texto.startswith('-'):
            negativo = True
            texto = texto[1:].strip()

        # remover palavras/cifras comuns (ex: R$, BRL) e espaços usados como milhares
        texto = re.sub(r'(?i)r\$|brl', '', texto)
        texto = texto.replace('\xa0', '').replace(' ', '')

        # heurística para separadores: caso possua '.' e ',' assume '.' milhares e ',' decimal
        if texto.count('.') > 0 and texto.count(',') > 0:
            texto = texto.replace('.', '').replace(',', '.')
        elif texto.count(',') > 0 and texto.count('.') == 0:
            # provável formato brasileiro 1234,56
            texto = texto.replace(',', '.')
        elif texto.count('.') > 0 and texto.count(',') == 0:
            # se houver múltiplos pontos, provavelmente são separadores de milhares
            if texto.count('.') > 1:
                texto = texto.replace('.', '')
            else:
                # se parte decimal tiver 3 dígitos, pode ser milhares (ex: 1.234)
                partes = texto.split('.')
                if len(partes) >= 2 and len(partes[-1]) == 3:
                    texto = ''.join(partes)

        # remove tudo que não for dígito, ponto ou sinal
        texto = re.sub(r'[^0-9\.\-]', '', texto)
        if texto in ('', '.', '-', '-.', '.-'):
            return 0.0

        valor = float(texto)
        if negativo:
            valor = -abs(valor)

        # arredondar a 2 casas para comparação consistente
        return round(valor, 2)
    except Exception:
        return 0.0


def normalizar_nd(valor):
    """Normaliza um código ND removendo caracteres não numéricos."""
    try:
        if valor is None:
            return ''
        if isinstance(valor, int):
            return str(valor)
        if isinstance(valor, float):
            if pd.isna(valor):
                return ''
            if valor.is_integer():
                return str(int(valor))
            s = str(valor)
        else:
            s = str(valor).strip()
            if s == '':
                return ''
            m_decimal_zero = re.fullmatch(r'(\d+)[\.,]0+', s)
            if m_decimal_zero:
                return m_decimal_zero.group(1)
        s = re.sub(r'\D', '', s)
        return s
    except Exception:
        return ''


def normalizar_pi(valor):
    """Normaliza PI (uppercase, sem espaços)."""
    try:
        if valor is None:
            return ''
        s = str(valor).strip().upper()
        s = re.sub(r'\s+', '', s)
        return s
    except Exception:
        return ''


def normalizar_om_solicitante_chave(valor):
    """Normaliza OM solicitante para comparação de duplicidades."""
    try:
        if valor is None:
            return ''
        s = str(valor).strip().upper()
        s = unicodedata.normalize('NFD', s)
        s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
        s = re.sub(r'[^A-Z0-9]+', '', s)
        return s
    except Exception:
        return ''


def _exportar_planilhas_excel(nome_base, planilhas):
    """Gera e envia um arquivo Excel com multiplas planilhas."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for nome_planilha, dados in planilhas:
            if isinstance(dados, pd.DataFrame):
                df = dados.copy()
            else:
                df = pd.DataFrame(dados or [])

            if df.empty:
                df = pd.DataFrame([{'info': 'Sem dados'}])

            nome_limpo = re.sub(r'[\\/*?:\[\]]', '_', str(nome_planilha or 'Planilha'))[:31]
            if not nome_limpo:
                nome_limpo = 'Planilha'

            df.to_excel(writer, sheet_name=nome_limpo, index=False)

    output.seek(0)
    carimbo = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(
        output,
        download_name=f'{nome_base}_{carimbo}.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


def normalizar_diex(valor):
    """Normaliza DIEx removendo espaços e caracteres estranhos, em maiúsculas."""
    try:
        if valor is None:
            return ''
        s = str(valor).strip().upper()
        s = re.sub(r'[^A-Z0-9\-]', '', s)
        return s
    except Exception:
        return ''


def _normalizar_numero_identificador(valor):
    texto = re.sub(r'\D', '', str(valor or ''))
    if not texto:
        return ''
    texto = texto.lstrip('0')
    return texto or '0'


def extrair_numero_diex(texto):
    """Extrai apenas a parte numérica do DIEx, normalizando zeros à esquerda."""
    try:
        if not texto:
            return ''
        valor = str(texto)

        if not re.search(r'DI\s*EX|DIEX', valor, flags=re.IGNORECASE):
            return ''

        m = re.search(r'DIEX[^0-9]{0,20}([0-9]{1,12})(?=\s*[-/]|\b)', valor, flags=re.IGNORECASE)
        if m:
            return _normalizar_numero_identificador(m.group(1))

        m = re.search(r'\b([0-9]{1,12})\s*-(?:[A-Z]|\d)', valor, flags=re.IGNORECASE)
        if m:
            return _normalizar_numero_identificador(m.group(1))

        return ''
    except Exception:
        return ''


def extrair_codom_numerico(texto):
    """Extrai COD OM numérico do padrão '(1234-XXXX)' e normaliza zeros à esquerda."""
    try:
        if not texto:
            return ''
        valor = str(texto)

        m = re.search(r'\(\s*([0-9]{1,12})\s*-', valor)
        if m:
            return _normalizar_numero_identificador(m.group(1))

        m = re.search(r'\(\s*([0-9]{1,12})\s*\)', valor)
        if m:
            return _normalizar_numero_identificador(m.group(1))

        m = re.search(r'\(\s*([0-9]{1,12})\b', valor)
        if m:
            return _normalizar_numero_identificador(m.group(1))

        m = re.fullmatch(r'\s*([0-9]{1,12})\s*', valor)
        if m:
            return _normalizar_numero_identificador(m.group(1))

        return ''
    except Exception:
        return ''


def extrair_codom_sistema_nc(nc):
    """Obtém COD OM numérico da NC/sua solicitação de origem."""
    for candidato in (getattr(nc, 'descricao', ''), getattr(nc, 'numero', ''), getattr(nc, 'diex_credito', '')):
        codom = extrair_codom_numerico(candidato)
        if codom:
            return codom

    try:
        solicitacao = nc.solicitacao_origem
    except Exception:
        solicitacao = None

    if solicitacao is not None:
        for pedido in getattr(solicitacao, 'pedidos', []) or []:
            codom = extrair_codom_numerico(getattr(pedido, 'codom', ''))
            if codom:
                return codom
            codom = extrair_codom_numerico(getattr(pedido, 'descricao_om', ''))
            if codom:
                return codom

        codom = extrair_codom_numerico(getattr(solicitacao, 'descricao', ''))
        if codom:
            return codom

    return ''


def extrair_diex_do_texto(texto):
    """Tenta extrair um identificador DIEx de um texto livre usando regex."""
    try:
        if not texto:
            return ''
        t = str(texto)
        # procurar padrões do tipo E6SUSOLA1QR ou palavras alfanuméricas curtas
        m = re.search(r'\bE\d[A-Z0-9]{5,}\b', t.upper())
        if m:
            return m.group(0)
        # procurar DIEx explícito
        m2 = re.search(r'\bDIEX[:\s]*([A-Z0-9\-]+)\b', t.upper())
        if m2:
            return m2.group(1)
        # fallback: palavra alfanumérica curta
        m3 = re.search(r'\b[A-Z0-9\-]{3,20}\b', t.upper())
        if m3:
            return m3.group(0)
        return ''
    except Exception:
        return ''


def valor_para_chave(v):
    """Converte um valor (numérico ou texto) para uma chave de comparação '0.00'."""
    try:
        f = normalizar_valor(v)
        valor_float = float(f)
        if pd.isna(valor_float) or math.isinf(valor_float):
            return '0.00'
        return f"{round(valor_float, 2):.2f}"
    except Exception:
        return '0.00'
def carregar_dados_fallback():
    """Dados de fallback quando o Excel não pode ser carregado"""
    print("🔄 Usando dados de fallback...")
    return {
        'OMs': [
            {'OM': 'AMAN', 'CODOM': '109', 'CODUG': '160001', 'SIGLA_UG': 'AMAN'},
            {'OM': '7º BEC', 'CODOM': '109', 'CODUG': '160001', 'SIGLA_UG': '7º BEC'},
            {'OM': 'MNMSGM', 'CODOM': '414', 'CODUG': '160068', 'SIGLA_UG': 'MNMSGM'},
            {'OM': 'AHEx', 'CODOM': '513', 'CODUG': '160068', 'SIGLA_UG': 'AHEx'},
            {'OM': 'DESMIL', 'CODOM': '45724', 'CODUG': '160068', 'SIGLA_UG': 'DESMIL'},
            {'OM': 'DETMIL', 'CODOM': '45732', 'CODUG': '160068', 'SIGLA_UG': 'DETMIL'},
            {'OM': 'DEPA', 'CODOM': '45740', 'CODUG': '160068', 'SIGLA_UG': 'DEPA'},
            {'OM': 'DPHCEx', 'CODOM': '46128', 'CODUG': '160068', 'SIGLA_UG': 'DPHCEx'},
            {'OM': 'AGR', 'CODOM': '703', 'CODUG': '160285', 'SIGLA_UG': 'AGR'},
            {'OM': 'AGSP', 'CODOM': '802', 'CODUG': '160529', 'SIGLA_UG': 'AGSP'}
        ]
    }


OMS_TABELA_ARQUIVO = 'oms_mapeamento.csv'
OM_GESTORA_RM_ARQUIVO = 'OM Gestora por RM.xlsx'
RM_CODUG_CACHE = {
    'arquivo': None,
    'mtime': None,
    'mapa': {}
}


def _caminho_tabela_oms():
    instance_dir_projeto = os.path.join(app.root_path, 'instance')
    os.makedirs(instance_dir_projeto, exist_ok=True)
    return os.path.join(instance_dir_projeto, OMS_TABELA_ARQUIVO)


def _normalizar_nome_coluna_excel(valor):
    txt = str(valor or '').strip().lower()
    txt = unicodedata.normalize('NFKD', txt)
    txt = ''.join(c for c in txt if not unicodedata.combining(c))
    return re.sub(r'[^a-z0-9]+', '', txt)


def _normalizar_rm_texto(valor):
    txt = str(valor or '').strip()
    if txt.lower() in ('nan', 'none'):
        return ''
    if re.fullmatch(r'\d+\.0+', txt):
        txt = txt.split('.', 1)[0]
    return txt


def _carregar_mapa_rm_por_codug():
    caminhos = [
        os.path.join(app.root_path, OM_GESTORA_RM_ARQUIVO),
        os.path.join(app.instance_path, OM_GESTORA_RM_ARQUIVO),
    ]
    existentes = [c for c in caminhos if os.path.exists(c)]
    if not existentes:
        RM_CODUG_CACHE.update({'arquivo': None, 'mtime': None, 'mapa': {}})
        return {}

    caminho = max(existentes, key=os.path.getmtime)
    mtime = os.path.getmtime(caminho)
    if RM_CODUG_CACHE.get('arquivo') == caminho and RM_CODUG_CACHE.get('mtime') == mtime:
        return RM_CODUG_CACHE.get('mapa', {})

    mapa = {}
    try:
        excel = pd.ExcelFile(caminho)
        for header_idx in (0, 1, 2, 3):
            encontrou_aba = False
            for aba in excel.sheet_names:
                try:
                    df = pd.read_excel(excel, sheet_name=aba, dtype=str, header=header_idx)
                except Exception:
                    continue
                if df is None or df.empty:
                    continue

                colunas_norm = {col: _normalizar_nome_coluna_excel(col) for col in df.columns}
                col_codug = next(
                    (
                        col for col, norm in colunas_norm.items()
                        if norm in ('codug', 'codigoug') or 'codug' in norm
                    ),
                    None
                )
                col_rm = next(
                    (
                        col for col, norm in colunas_norm.items()
                        if norm == 'rm' or norm.endswith('rm') or 'regiaomilitar' in norm
                    ),
                    None
                )
                if not col_codug or not col_rm:
                    continue

                encontrou_aba = True
                for _, row in df.iterrows():
                    codug = normalizar_cod_ug(row.get(col_codug, ''))
                    rm = _normalizar_rm_texto(row.get(col_rm, ''))
                    if codug and rm and codug not in mapa:
                        mapa[codug] = rm
            if encontrou_aba and mapa:
                break
    except Exception:
        mapa = {}

    RM_CODUG_CACHE.update({'arquivo': caminho, 'mtime': mtime, 'mapa': mapa})
    return mapa


def _obter_rm_por_codug(codug):
    codug_norm = normalizar_cod_ug(codug)
    if not codug_norm:
        return ''
    return _carregar_mapa_rm_por_codug().get(codug_norm, '')


def _normalizar_linha_om(om='', codom='', codug='', sigla_ug='', rm=''):
    om_txt = str(om or '').strip()
    if om_txt.lower() in ('nan', 'none'):
        om_txt = ''
    codom_txt = re.sub(r'\D', '', str(codom or '').strip())
    codug_txt = re.sub(r'\D', '', str(codug or '').strip())
    rm_txt = _normalizar_rm_texto(rm)
    if not rm_txt:
        rm_txt = _obter_rm_por_codug(codug_txt)
    sigla_txt = str(sigla_ug or '').strip()
    if sigla_txt.lower() in ('nan', 'none'):
        sigla_txt = ''
    return {
        'RM': rm_txt,
        'OM': om_txt,
        'CODOM': codom_txt,
        'CODUG': codug_txt,
        'SIGLA_UG': sigla_txt,
    }


def _salvar_tabela_oms(oms_lista):
    caminho = _caminho_tabela_oms()
    df = pd.DataFrame([_normalizar_linha_om(
        rm=item.get('RM', ''),
        om=item.get('OM', ''),
        codom=item.get('CODOM', ''),
        codug=item.get('CODUG', ''),
        sigla_ug=item.get('SIGLA_UG', '')
    ) for item in (oms_lista or [])])

    if not df.empty:
        df = df[df['OM'].astype(str).str.strip() != '']
        df = df.drop_duplicates(subset=['OM'], keep='first')

    df.to_csv(caminho, index=False, encoding='utf-8-sig')


def _carregar_tabela_oms():
    caminho = _caminho_tabela_oms()
    if not os.path.exists(caminho):
        return []
    try:
        df = pd.read_csv(caminho, dtype=str)
    except Exception:
        return []

    if df.empty:
        return []

    retorno = []
    for _, row in df.iterrows():
        item = _normalizar_linha_om(
            rm=row.get('RM', ''),
            om=row.get('OM', ''),
            codom=row.get('CODOM', ''),
            codug=row.get('CODUG', ''),
            sigla_ug=row.get('SIGLA_UG', ''),
        )
        if item['OM']:
            retorno.append(item)

    retorno.sort(key=lambda x: x.get('OM', '').upper())
    return retorno


def _consolidar_oms_por_nome(oms_lista):
    grupos = {}
    for item in oms_lista or []:
        om_nome = str(item.get('OM', '') or '').strip()
        if not om_nome:
            continue
        chave = om_nome.upper()
        grupos.setdefault(chave, {'OM': om_nome, 'itens': []})
        grupos[chave]['itens'].append(_normalizar_linha_om(
            om=om_nome,
            codom=item.get('CODOM', ''),
            codug=item.get('CODUG', ''),
            sigla_ug=item.get('SIGLA_UG', ''),
            rm=item.get('RM', ''),
        ))

    consolidados = []
    for dados in grupos.values():
        itens = dados['itens']
        contagem = {}
        for it in itens:
            par = (it.get('CODOM', ''), it.get('CODUG', ''), it.get('SIGLA_UG', ''), it.get('RM', ''))
            contagem[par] = contagem.get(par, 0) + 1

        melhor_par = sorted(
            contagem.items(),
            key=lambda kv: (kv[1], 1 if kv[0][1] else 0, 1 if kv[0][0] else 0),
            reverse=True
        )[0][0]

        consolidados.append({
            'RM': melhor_par[3],
            'OM': dados['OM'],
            'CODOM': melhor_par[0],
            'CODUG': melhor_par[1],
            'SIGLA_UG': melhor_par[2],
        })

    consolidados.sort(key=lambda x: x.get('OM', '').upper())
    return consolidados


def _obter_oms_data():
    # Prioriza sempre a tabela persistida (fonte de verdade editável no sistema)
    tabela_oms = _carregar_tabela_oms()
    if tabela_oms:
        return tabela_oms

    dados_globais = globals().get('DADOS_PLANILHA', {'OMs': []})
    if dados_globais and dados_globais.get('OMs'):
        return dados_globais['OMs']
    return []


def _resolver_correspondencia_om(nome_om, codom_form='', codug_form='', sigla_form=''):
    nome_ref = str(nome_om or '').strip().upper()
    base_oms = _obter_oms_data()

    # Match exato por OM (mais confiável).
    for om in base_oms:
        if (om.get('OM', '') or '').strip().upper() == nome_ref:
            return {
                'codom': (om.get('CODOM', '') or '').strip(),
                'codug': (om.get('CODUG', '') or '').strip(),
                'sigla_ug': (om.get('SIGLA_UG', '') or '').strip(),
                'fonte': 'tabela_oms'
            }

    # Match tolerante por chave normalizada para variações de pontuação/acentuação.
    chave_ref = normalizar_om_solicitante_chave(nome_om)
    if chave_ref:
        for om in base_oms:
            chave_om = normalizar_om_solicitante_chave(om.get('OM', ''))
            chave_sigla = normalizar_om_solicitante_chave(om.get('SIGLA_UG', ''))
            if chave_ref and (chave_ref == chave_om or chave_ref == chave_sigla):
                return {
                    'codom': (om.get('CODOM', '') or '').strip(),
                    'codug': (om.get('CODUG', '') or '').strip(),
                    'sigla_ug': (om.get('SIGLA_UG', '') or '').strip(),
                    'fonte': 'tabela_oms_normalizada'
                }

    return {
        'codom': str(codom_form or '').strip(),
        'codug': str(codug_form or '').strip(),
        'sigla_ug': str(sigla_form or '').strip(),
        'fonte': 'form_fallback'
    }


def _extrair_oms_da_extra_pdrlog_xlsx(arquivo_upload=None):
    def limpar(valor):
        if pd.isna(valor):
            return ''
        texto = str(valor).strip()
        if texto.lower() in ('nan', 'none'):
            return ''
        if re.fullmatch(r'\d+\.0+', texto):
            texto = texto.split('.', 1)[0]
        return texto

    def extrair_oms_do_dataframe(df_xlsx):
        df_xlsx = df_xlsx.copy()
        df_xlsx.columns = [str(col).strip() for col in df_xlsx.columns]
        oms_local = []

        # Bloco oficial de correspondência (lado direito da planilha CODOM)
        if {'OM', 'COD OM.1', 'COD UG.1', 'SIGLA UG.1'}.issubset(set(df_xlsx.columns)):
            bloco_oficial = []
            for _, row in df_xlsx.iterrows():
                om_nome = limpar(row.get('OM', ''))
                if not om_nome:
                    continue
                bloco_oficial.append({
                    'OM': om_nome,
                    'CODOM': limpar(row.get('COD OM.1', '')),
                    'CODUG': re.sub(r'\D', '', limpar(row.get('COD UG.1', ''))),
                    'SIGLA_UG': limpar(row.get('SIGLA UG.1', '')),
                })
            if bloco_oficial:
                return bloco_oficial

        # Fallback para layouts antigos
        if {'COD OM', 'SIGLA OM', 'COD UG', 'SIGLA UG'}.issubset(set(df_xlsx.columns)):
            for _, row in df_xlsx.iterrows():
                om_nome = limpar(row.get('SIGLA OM', ''))
                if not om_nome:
                    continue
                oms_local.append({
                    'OM': om_nome,
                    'CODOM': limpar(row.get('COD OM', '')),
                    'CODUG': re.sub(r'\D', '', limpar(row.get('COD UG', ''))),
                    'SIGLA_UG': limpar(row.get('SIGLA UG', '')),
                })

        if not oms_local:
            col_om = next((c for c in df_xlsx.columns if str(c).strip().upper() in ('OM', 'SIGLA OM')), None)
            col_codom = next((c for c in df_xlsx.columns if 'COD OM' in str(c).upper()), None)
            col_codug = next((c for c in df_xlsx.columns if 'COD UG' in str(c).upper()), None)
            col_sigla = next((c for c in df_xlsx.columns if 'SIGLA UG' in str(c).upper()), None)
            if col_om and col_codug:
                for _, row in df_xlsx.iterrows():
                    om_nome = limpar(row.get(col_om, ''))
                    if not om_nome:
                        continue
                    oms_local.append({
                        'OM': om_nome,
                        'CODOM': limpar(row.get(col_codom, '')) if col_codom else '',
                        'CODUG': re.sub(r'\D', '', limpar(row.get(col_codug, ''))),
                        'SIGLA_UG': limpar(row.get(col_sigla, '')) if col_sigla else '',
                    })

        return oms_local

    def ler_aba_preferencial(caminho_arquivo):
        excel_local = pd.ExcelFile(caminho_arquivo)
        aba = next((a for a in excel_local.sheet_names if str(a).strip().lower() == 'codom'), excel_local.sheet_names[0])
        return pd.read_excel(caminho_arquivo, sheet_name=aba, dtype=str)

    if arquivo_upload is not None:
        try:
            arquivo_upload.stream.seek(0)
        except Exception:
            pass
        excel = pd.ExcelFile(arquivo_upload)
        aba_codom = next((aba for aba in excel.sheet_names if str(aba).strip().lower() == 'codom'), excel.sheet_names[0])
        df_xlsx = pd.read_excel(excel, sheet_name=aba_codom, dtype=str)
        oms = extrair_oms_do_dataframe(df_xlsx)
        return _consolidar_oms_por_nome(oms)
    else:
        caminhos_xlsx = [
            os.path.join(app.root_path, 'Extra PDRLOG.xlsx'),
            os.path.join(app.instance_path, 'Extra PDRLOG.xlsx')
        ]
        candidatos = [c for c in caminhos_xlsx if os.path.exists(c)]
        if not candidatos:
            raise FileNotFoundError('Arquivo Extra PDRLOG.xlsx não encontrado.')

        candidatos_ordenados = sorted(candidatos, key=os.path.getmtime, reverse=True)
        melhor_oms = []
        for caminho in candidatos_ordenados:
            try:
                df_xlsx = ler_aba_preferencial(caminho)
                oms_caminho = extrair_oms_do_dataframe(df_xlsx)
                if len(oms_caminho) > len(melhor_oms):
                    melhor_oms = oms_caminho
            except Exception:
                continue

        if not melhor_oms:
            raise ValueError('Nenhum mapeamento OM/CODOM/CODUG válido foi encontrado na planilha Extra PDRLOG.xlsx.')

        return _consolidar_oms_por_nome(melhor_oms)

def normalizar_oms_dados(dados):
    """Remove NaN e converte códigos para strings para JSON seguro."""
    oms_normalizados = []

    def limpar(valor):
        if pd.isna(valor):
            return ''
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        return str(valor).strip()

    for om in dados.get('OMs', []):
        oms_normalizados.append({
            'RM': limpar(om.get('RM', '')),
            'OM': limpar(om.get('OM', '')),
            'CODOM': limpar(om.get('CODOM', '')),
            'CODUG': limpar(om.get('CODUG', '')),
            'SIGLA_UG': limpar(om.get('SIGLA_UG', '')),
        })

    return {'OMs': oms_normalizados}

# Função segura para carregar dados da planilha com fallback
def carregar_dados_planilha():
    """Carrega OMs priorizando `Extra PDRLOG.xlsx`; em fallback usa CSV/cache e por fim dados fixos."""
    try:
        # 0) Prioridade máxima: tabela persistida no sistema (editável no painel)
        tabela_oms = _carregar_tabela_oms()
        if tabela_oms:
            return {'OMs': tabela_oms}

        def limpar(valor):
            if pd.isna(valor):
                return ''
            texto = str(valor).strip()
            if texto.lower() in ('nan', 'none'):
                return ''
            if re.fullmatch(r'\d+\.0+', texto):
                texto = texto.split('.', 1)[0]
            return texto

        # 1) Prioridade: planilha Extra PDRLOG.xlsx (raiz e/ou instance)
        caminhos_xlsx = [
            os.path.join(app.root_path, 'Extra PDRLOG.xlsx'),
            os.path.join(app.instance_path, 'Extra PDRLOG.xlsx')
        ]
        for caminho_xlsx in caminhos_xlsx:
            if not os.path.exists(caminho_xlsx):
                continue

            try:
                excel = pd.ExcelFile(caminho_xlsx)
                aba_codom = next((aba for aba in excel.sheet_names if str(aba).strip().lower() == 'codom'), excel.sheet_names[0])
                df_xlsx = pd.read_excel(caminho_xlsx, sheet_name=aba_codom, dtype=str)
                df_xlsx.columns = [str(col).strip() for col in df_xlsx.columns]

                oms = []

                # Bloco principal: COD OM | SIGLA OM | COD UG | SIGLA UG
                if {'COD OM', 'SIGLA OM', 'COD UG', 'SIGLA UG'}.issubset(set(df_xlsx.columns)):
                    for _, row in df_xlsx.iterrows():
                        om_nome = limpar(row.get('SIGLA OM', ''))
                        if not om_nome:
                            continue
                        oms.append({
                            'OM': om_nome,
                            'CODOM': limpar(row.get('COD OM', '')),
                            'CODUG': re.sub(r'\D', '', limpar(row.get('COD UG', ''))),
                            'SIGLA_UG': limpar(row.get('SIGLA UG', '')),
                        })

                # Bloco lateral: OM | COD OM.1 | COD UG.1 | SIGLA UG.1
                if {'OM', 'COD OM.1', 'COD UG.1', 'SIGLA UG.1'}.issubset(set(df_xlsx.columns)):
                    for _, row in df_xlsx.iterrows():
                        om_nome = limpar(row.get('OM', ''))
                        if not om_nome:
                            continue
                        oms.append({
                            'OM': om_nome,
                            'CODOM': limpar(row.get('COD OM.1', '')),
                            'CODUG': re.sub(r'\D', '', limpar(row.get('COD UG.1', ''))),
                            'SIGLA_UG': limpar(row.get('SIGLA UG.1', '')),
                        })

                # Caso não esteja no layout CODOM, tenta detecção genérica simples
                if not oms:
                    col_om = next((c for c in df_xlsx.columns if str(c).strip().upper() in ('OM', 'SIGLA OM')), None)
                    col_codom = next((c for c in df_xlsx.columns if 'COD OM' in str(c).upper()), None)
                    col_codug = next((c for c in df_xlsx.columns if 'COD UG' in str(c).upper()), None)
                    col_sigla = next((c for c in df_xlsx.columns if 'SIGLA UG' in str(c).upper()), None)

                    if col_om and col_codug:
                        for _, row in df_xlsx.iterrows():
                            om_nome = limpar(row.get(col_om, ''))
                            if not om_nome:
                                continue
                            oms.append({
                                'OM': om_nome,
                                'CODOM': limpar(row.get(col_codom, '')) if col_codom else '',
                                'CODUG': re.sub(r'\D', '', limpar(row.get(col_codug, ''))),
                                'SIGLA_UG': limpar(row.get(col_sigla, '')) if col_sigla else '',
                            })

                if oms:
                    unicos = _consolidar_oms_por_nome(oms)
                    if unicos:
                        return {'OMs': unicos}
            except Exception:
                pass

        # 2) Fallback: CSV cache legado
        caminho = os.path.join(app.instance_path, 'Extra_PDRLOG_cache.csv')
        if os.path.exists(caminho):
            df = pd.read_csv(caminho, dtype=str)

            def find_col(possibilities):
                for col in df.columns:
                    norm_col = ''.join(c for c in str(col).lower() if c.isalnum())
                    for p in possibilities:
                        alvo = ''.join(c for c in str(p).lower() if c.isalnum())
                        if norm_col == alvo or alvo in norm_col:
                            return col
                return None

            col_om = find_col(['OM', 'NOME', 'ORGAO'])
            col_codom = find_col(['CODOM', 'COD OM', 'CODIGO OM'])
            col_codug = find_col(['CODUG', 'COD UG', 'UG EXECUTORA', 'FAVORECIDO DOC NUMERO'])
            col_sigla = find_col(['SIGLA_UG', 'SIGLA'])

            # Fallback para layout recente do CSV (OM em Unnamed: 1 e CODUG em UG Executora)
            if not col_om and 'Unnamed: 1' in df.columns:
                col_om = 'Unnamed: 1'
            if not col_codug:
                primeira_coluna = str(df.columns[0]) if len(df.columns) else ''
                if 'ug' in primeira_coluna.lower() and 'executora' in primeira_coluna.lower():
                    col_codug = df.columns[0]

            oms = []
            for _, row in df.iterrows():
                om_valor = limpar(row.get(col_om, '') if col_om else '')
                codom_valor = limpar(row.get(col_codom, '') if col_codom else '')
                codug_valor = limpar(row.get(col_codug, '') if col_codug else '')
                sigla_valor = limpar(row.get(col_sigla, '') if col_sigla else '')

                # Limpa CODUG em formato 160001.0
                if re.fullmatch(r'\d+\.0+', codug_valor):
                    codug_valor = codug_valor.split('.', 1)[0]
                codug_limpo = re.sub(r'\D', '', codug_valor)
                if codug_limpo:
                    codug_valor = codug_limpo

                # Ignora linhas sem dados úteis de OM
                if not om_valor:
                    continue

                oms.append({
                    'OM': om_valor,
                    'CODOM': codom_valor,
                    'CODUG': codug_valor,
                    'SIGLA_UG': sigla_valor,
                })

            if oms:
                unicos = _consolidar_oms_por_nome(oms)
                return {'OMs': unicos}
        # Se não existir arquivo ou não obteve dados, usar fallback
        return carregar_dados_fallback()
    except Exception:
        return carregar_dados_fallback()

# Garante que DADOS_PLANILHA sempre exista, mesmo se houver erro na leitura da planilha
def _recarregar_dados_planilha_globais():
    global DADOS_PLANILHA
    try:
        DADOS_PLANILHA = normalizar_oms_dados(carregar_dados_planilha())
        if not DADOS_PLANILHA or 'OMs' not in DADOS_PLANILHA:
            DADOS_PLANILHA = {'OMs': []}
        # Semeia a tabela persistida na primeira execução, se ainda não existir.
        if DADOS_PLANILHA.get('OMs') and not _carregar_tabela_oms():
            _salvar_tabela_oms(DADOS_PLANILHA.get('OMs', []))
    except Exception:
        DADOS_PLANILHA = {'OMs': []}
    return DADOS_PLANILHA


_recarregar_dados_planilha_globais()
print(f"OMs carregadas: {len(DADOS_PLANILHA['OMs'])}")

PDR_LOG_CACHE = {
    'mtime': None,
    'data': []
}
PDR_LOG_NC_CACHE_FILE = 'pdr_log_nc_siafi_cache.csv'
AUDITORIA_UPLOAD_DIR = 'auditoria_upload'
AUDITORIA_UPLOAD_ARQUIVO_BASE = 'auditoria_unificada.xlsx'
AUDITORIA_UPLOAD_EXTRA_CACHE = 'auditoria_extra_cache.csv'
AUDITORIA_UPLOAD_PDR_CACHE = 'auditoria_pdr_log_cache.csv'
AUDITORIA_UPLOAD_SISNC_ARQUIVO_BASE = 'sisnc_unificada.xlsx'
AUDITORIA_UPLOAD_SISNC_EXTRA_CACHE = 'sisnc_extra_cache.csv'
AUDITORIA_UPLOAD_SISNC_PDR_CACHE = 'sisnc_pdr_log_cache.csv'


def _normalizar_colunas_numericas(df):
    cols = list(df.columns)
    novas_colunas = []
    alterou = False

    # Primeiro: converte nomes de coluna que são apenas dígitos ("0","1",...) em ints
    for col in cols:
        texto = str(col).strip()
        if re.fullmatch(r"\d+", texto):
            novas_colunas.append(int(texto))
            alterou = True
        else:
            novas_colunas.append(col)

    # Se não houve conversão por dígitos, tenta detectar formatos comuns de exportação
    # que usam nomes genéricos como 'COLUNA', 'COLUNA_2', 'COLUNA 3' e normaliza para índices
    if not alterou:
        padrao_coluna_generica = re.compile(r"^COLUNA(?:[_\s]?\d+)?$", re.IGNORECASE)
        if all(padrao_coluna_generica.match(str(c).strip()) for c in cols):
            novas_colunas = list(range(len(cols)))
            alterou = True

    if alterou:
        df.columns = novas_colunas
    return df


def _caminhos_upload_auditoria():
    pasta = os.path.join(app.instance_path, AUDITORIA_UPLOAD_DIR)
    os.makedirs(pasta, exist_ok=True)
    return {
        'dir': pasta,
        'base': os.path.join(pasta, AUDITORIA_UPLOAD_ARQUIVO_BASE),
        'extra': os.path.join(pasta, AUDITORIA_UPLOAD_EXTRA_CACHE),
        'pdr': os.path.join(pasta, AUDITORIA_UPLOAD_PDR_CACHE),
        'meta': os.path.join(pasta, 'meta.json'),
        'sisnc_base': os.path.join(pasta, AUDITORIA_UPLOAD_SISNC_ARQUIVO_BASE),
        'sisnc_extra': os.path.join(pasta, AUDITORIA_UPLOAD_SISNC_EXTRA_CACHE),
        'sisnc_pdr': os.path.join(pasta, AUDITORIA_UPLOAD_SISNC_PDR_CACHE),
        'sisnc_meta': os.path.join(pasta, 'meta_sisnc.json')
    }


def _carregar_meta_upload_auditoria(fonte='tg'):
    caminhos = _caminhos_upload_auditoria()
    caminho_meta = caminhos['meta'] if fonte != 'sisnc' else caminhos['sisnc_meta']
    if not os.path.exists(caminho_meta):
        return None
    try:
        with open(caminho_meta, 'r', encoding='utf-8') as f:
            meta = json.load(f)
    except Exception:
        return None

    uploaded_at = str(meta.get('uploaded_at', '') or '').strip()
    uploaded_at_fmt = ''
    if uploaded_at:
        try:
            uploaded_at_fmt = datetime.fromisoformat(uploaded_at).strftime('%d/%m/%Y %H:%M:%S')
        except Exception:
            uploaded_at_fmt = uploaded_at

    return {
        'uploaded_at': uploaded_at,
        'uploaded_at_fmt': uploaded_at_fmt,
        'total': int(meta.get('total', 0) or 0),
        'pdr': int(meta.get('pdr', 0) or 0),
        'extra': int(meta.get('extra', 0) or 0),
        'col_pi': str(meta.get('col_pi', '') or '')
    }


def _normalizar_nome_coluna_upload(valor):
    texto = str(valor or '')
    texto = unicodedata.normalize('NFD', texto)
    texto = ''.join(ch for ch in texto if unicodedata.category(ch) != 'Mn')
    texto = texto.lower().strip()
    return re.sub(r'[^a-z0-9]+', '', texto)


def _encontrar_coluna_upload(df, candidatos):
    mapa = {}
    for col in df.columns:
        mapa[_normalizar_nome_coluna_upload(col)] = col

    for cand in candidatos:
        chave = _normalizar_nome_coluna_upload(cand)
        if chave in mapa:
            return mapa[chave]

    for cand in candidatos:
        chave = _normalizar_nome_coluna_upload(cand)
        for chave_col, col_original in mapa.items():
            if chave and chave in chave_col:
                return col_original

    return None


def _inferir_coluna_pi_sisnc_por_conteudo(df):
    """Tenta inferir a coluna PI em layouts SISNC com cabecalho nao padronizado."""
    melhor_coluna = None
    melhor_score = 0
    padrao_pi = re.compile(r'(PLJ|SOL|\bE\d[A-Z0-9]{8,}\b)', re.IGNORECASE)

    for col in df.columns:
        try:
            serie = df[col].fillna('').astype(str).str.upper().str.strip()
        except Exception:
            continue

        if serie.empty:
            continue

        score = int(serie.str.contains(padrao_pi, na=False, regex=True).sum())
        if score > melhor_score:
            melhor_score = score
            melhor_coluna = col

    if melhor_score > 0:
        return melhor_coluna
    return None


def _detectar_dataframe_sisnc_upload(arquivo_upload):
    """Le o Excel SISNC com tolerancia a diferentes linhas de cabecalho."""

    candidatos_pi = ['PI', 'Plano Interno', 'Projeto/Atividade', 'Projeto Atividade', 'PI/Localizador']
    candidatos_nd = ['ND', 'Natureza da Despesa', 'Natureza Despesa']
    candidatos_codug = [
        'Código UG', 'Codigo UG', 'COD UG', 'CODUG',
        'UG Emitente', 'UG Emitente Cod', 'UG Código'
    ]
    candidatos_doc = [
        'Finalidade', 'Finalidade/Observação', 'Finalidade/Observacao',
        'Finalidade Observação', 'Finalidade Observacao',
        'Doc - Observação', 'Doc - Observacao', 'Doc Observacao',
        'Observação', 'Observacao', 'Historico'
    ]
    candidatos_valor = ['Valor', 'Valor Total', 'Valor R$', 'Valor (R$)', 'Vlr']

    nome_arquivo = str(getattr(arquivo_upload, 'filename', '') or '').lower()

    def _reset_upload():
        try:
            arquivo_upload.stream.seek(0)
            return
        except Exception:
            pass
        try:
            arquivo_upload.seek(0)
        except Exception:
            pass

    def _ler_excel_sisnc(**kwargs):
        alvo = getattr(arquivo_upload, 'stream', arquivo_upload)
        params = dict(kwargs)

        if 'engine' not in params:
            if nome_arquivo.endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
                params['engine'] = 'openpyxl'
            elif nome_arquivo.endswith('.xls'):
                params['engine'] = 'xlrd'

        _reset_upload()
        try:
            return pd.read_excel(alvo, **params)
        except Exception as e:
            if params.get('engine') == 'xlrd':
                raise ValueError(
                    'Não foi possível ler arquivo SISNC em formato .xls neste ambiente. '
                    'Converta para .xlsx e tente novamente.'
                ) from e
            if 'engine' in params:
                _reset_upload()
                try:
                    return pd.read_excel(alvo, **kwargs)
                except Exception:
                    pass

            # Alguns exportadores do SISNC geram HTML com extensao .xls/.xlsx.
            # Nesse caso, tenta extrair a primeira tabela via read_html.
            _reset_upload()
            try:
                bruto = alvo.read()
            except Exception:
                bruto = b''

            if isinstance(bruto, bytes) and bruto:
                texto_html = None
                for enc in ('utf-8', 'cp1252', 'latin-1'):
                    try:
                        texto_html = bruto.decode(enc)
                        break
                    except Exception:
                        continue
                if texto_html:
                    try:
                        tabelas = pd.read_html(StringIO(texto_html), header=kwargs.get('header', None))
                        if tabelas:
                            return tabelas[0]
                    except Exception:
                        pass
            raise

    def _limpar_df(df_local):
        if df_local is None:
            return pd.DataFrame()
        df_local = _normalizar_colunas_numericas(df_local)
        df_local.columns = [str(col).strip() for col in df_local.columns]
        df_local = df_local.dropna(how='all').dropna(axis=1, how='all')
        return df_local

    def _score_cabecalho_sisnc(valores_linha):
        tokens = [
            'PI', 'PROJETOATIVIDADE', 'PLANOINTERNO',
            'ND', 'NATUREZADESPESA',
            'CODIGOUG', 'CODUG', 'UG',
            'FINALIDADE', 'DOCOBSERVACAO', 'OBSERVACAO',
            'VALOR', 'VALORTOTAL',
            'NCSIAFI', 'REFERENCIA'
        ]
        score = 0
        for valor in valores_linha:
            cel = _normalizar_nome_coluna_upload(valor)
            if not cel or cel.startswith('unnamed'):
                continue
            if any(tok in cel for tok in tokens):
                score += 1
        return score

    def _pontuar_df_sisnc(df_local):
        if df_local is None or df_local.empty:
            return -1

        score = 0
        if _encontrar_coluna_upload(df_local, candidatos_pi) or _inferir_coluna_pi_sisnc_por_conteudo(df_local):
            score += 1
        if _encontrar_coluna_upload(df_local, candidatos_nd):
            score += 2
        if _encontrar_coluna_upload(df_local, candidatos_codug):
            score += 2
        if _encontrar_coluna_upload(df_local, candidatos_doc):
            score += 2
        if _encontrar_coluna_upload(df_local, candidatos_valor):
            score += 2
        return score

    try:
        df_bruto = _ler_excel_sisnc(header=None, dtype=object)
    except Exception as e:
        raise ValueError(f'Erro ao ler o Excel SISNC: {str(e)}')

    if df_bruto is None or df_bruto.empty:
        return pd.DataFrame()

    candidatos = []

    # Mantem compatibilidade com layout esperado anteriormente (header=1).
    try:
        candidatos.append(_limpar_df(_ler_excel_sisnc(header=1)))
    except Exception:
        pass

    # Fallback comum quando cabecalho vem na primeira linha.
    try:
        candidatos.append(_limpar_df(_ler_excel_sisnc(header=0)))
    except Exception:
        pass

    # Detecta dinamicamente a linha de cabecalho com melhor pontuacao.
    melhor_idx = None
    melhor_score = -1
    limite = min(len(df_bruto), 20)
    for idx in range(limite):
        score = _score_cabecalho_sisnc(df_bruto.iloc[idx].tolist())
        if score > melhor_score:
            melhor_score = score
            melhor_idx = idx

    if melhor_idx is not None and melhor_score >= 2:
        cabecalho = [str(v).strip() if pd.notna(v) else '' for v in df_bruto.iloc[melhor_idx].tolist()]
        df_detectado = df_bruto.iloc[melhor_idx + 1:].copy()
        df_detectado.columns = cabecalho
        candidatos.append(_limpar_df(df_detectado))

    melhor_df = None
    melhor_score_df = -1
    for df_cand in candidatos:
        score_df = _pontuar_df_sisnc(df_cand)
        if score_df > melhor_score_df:
            melhor_score_df = score_df
            melhor_df = df_cand

    if melhor_df is not None and not melhor_df.empty:
        return melhor_df

    # Ultimo fallback: usa o bruto e promove a primeira linha a cabecalho.
    cabecalho = [str(v).strip() if pd.notna(v) else '' for v in df_bruto.iloc[0].tolist()]
    df_fallback = df_bruto.iloc[1:].copy()
    df_fallback.columns = cabecalho
    return _limpar_df(df_fallback)


def _normalizar_texto_bi(valor):
    texto = str(valor or '')
    texto = unicodedata.normalize('NFD', texto)
    texto = ''.join(ch for ch in texto if unicodedata.category(ch) != 'Mn')
    return texto.upper().strip()


def _carregar_dataframe_bi_upload(arquivo_upload):
    def _reset_stream_upload(upload_obj):
        try:
            upload_obj.stream.seek(0)
        except Exception:
            pass

    def _find_col_bi(df_local, possibilidades):
        for col in df_local.columns:
            norm_col = ''.join(c for c in _normalizar_texto_bi(col).lower() if c.isalnum())
            for p in possibilidades:
                norm_p = ''.join(c for c in _normalizar_texto_bi(p).lower() if c.isalnum())
                if norm_col == norm_p or norm_p in norm_col:
                    return col
        return None

    def _pontuar_layout_bi(df_local):
        encontrados = 0
        if _find_col_bi(df_local, ['PI', 'PROJETO/ATIVIDADE', 'PROJETO ATIVIDADE']):
            encontrados += 1
        if _find_col_bi(df_local, ['UG EXECUTORA', 'UG', 'UG_EXECUTORA', 'UGEXECUTORA']):
            encontrados += 1
        if _find_col_bi(df_local, ['CREDITO DISPONIVEL', 'CRÉDITO DISPONÍVEL', '19']):
            encontrados += 1
        if _find_col_bi(df_local, ['DESPESAS EMPENHADAS', 'EMPENHADO', '23']):
            encontrados += 1
        if _find_col_bi(df_local, ['DESPESAS LIQUIDADAS', 'LIQUIDADO', '25']):
            encontrados += 1
        if _find_col_bi(df_local, ['DESPESAS PAGAS', 'PAGO', '28']):
            encontrados += 1
        if _find_col_bi(df_local, ['GRUPO DESPESA', 'GND']):
            encontrados += 1
        return encontrados

    _reset_stream_upload(arquivo_upload)
    try:
        df_bruto = pd.read_excel(arquivo_upload, header=None, dtype=object)
    except Exception as e:
        nome = str(getattr(arquivo_upload, 'filename', '') or '').lower()
        if nome.endswith('.xls'):
            raise ValueError(
                'Não foi possível ler arquivo .xls neste ambiente. '
                'Converta a planilha BI para .xlsx e tente novamente.'
            ) from e
        raise ValueError(f'Erro ao ler planilha BI: {str(e)}') from e

    if df_bruto is None or df_bruto.empty:
        raise ValueError('Planilha BI vazia ou sem dados válidos.')

    max_linhas = min(len(df_bruto), 25)
    melhor_idx = 0
    melhor_score = -1

    tokens_relevantes = [
        'PI', 'PROJETO/ATIVIDADE', 'PROJETO ATIVIDADE',
        'GRUPO DESPESA', 'GND',
        'UG EXECUTORA', 'UG',
        'CREDITO DISPONIVEL', 'CREDITO',
        'DESPESAS EMPENHADAS', 'EMPENHADO',
        'DESPESAS LIQUIDADAS', 'LIQUIDADO',
        'DESPESAS PAGAS', 'PAGO'
    ]
    tokens_norm = [_normalizar_texto_bi(t) for t in tokens_relevantes]

    for idx in range(max_linhas):
        valores = [_normalizar_texto_bi(v) for v in df_bruto.iloc[idx].tolist()]
        if not any(valores):
            continue

        score = 0
        for celula in valores:
            if not celula:
                continue
            if celula.startswith('UNNAMED'):
                continue
            for token in tokens_norm:
                if token and token in celula:
                    score += 1

        if score > melhor_score:
            melhor_score = score
            melhor_idx = idx

    if melhor_score <= 0:
        # fallback para leitura padrão caso não seja possível identificar um cabeçalho útil
        _reset_stream_upload(arquivo_upload)
        try:
            df_fallback = pd.read_excel(arquivo_upload)
        except Exception as e:
            raise ValueError(f'Erro ao ler planilha BI (fallback): {str(e)}') from e

        if df_fallback is None or df_fallback.empty:
            raise ValueError('Planilha BI sem dados após fallback de leitura.')
        df_fallback = df_fallback.dropna(how='all').dropna(axis=1, how='all')
        return df_fallback, {'header_row': 1, 'score': 0, 'modo': 'padrao'}

    cabecalho = [str(v).strip() if pd.notna(v) else '' for v in df_bruto.iloc[melhor_idx].tolist()]
    df = df_bruto.iloc[melhor_idx + 1:].copy()
    df.columns = cabecalho
    df = df.dropna(how='all').dropna(axis=1, how='all')

    # Garante nomes de colunas únicos para evitar ambiguidades em processamento posterior
    nomes = []
    contagem = {}
    for col in df.columns:
        base = str(col).strip() if str(col).strip() else 'COLUNA'
        qtd = contagem.get(base, 0)
        contagem[base] = qtd + 1
        nomes.append(base if qtd == 0 else f'{base}_{qtd + 1}')
    df.columns = nomes

    if df.empty:
        raise ValueError('Planilha BI não possui linhas de dados após identificar o cabeçalho.')

    # Se o layout detectado não apresentar colunas típicas de BI, usa fallback padrão.
    if _pontuar_layout_bi(df) < 2:
        _reset_stream_upload(arquivo_upload)
        try:
            df_fallback = pd.read_excel(arquivo_upload)
        except Exception as e:
            raise ValueError(f'Erro ao ler planilha BI (fallback): {str(e)}') from e
        if df_fallback is None or df_fallback.empty:
            raise ValueError('Planilha BI sem dados após fallback de leitura.')
        df_fallback = df_fallback.dropna(how='all').dropna(axis=1, how='all')
        return df_fallback, {'header_row': 1, 'score': int(melhor_score), 'modo': 'padrao'}

    return df, {'header_row': int(melhor_idx + 1), 'score': int(melhor_score), 'modo': 'detectado'}


def _detectar_coluna_pi_upload(df):
    candidatos_nome = {'PI', 'PROJETO/ATIVIDADE', 'PROJETO ATIVIDADE'}
    for col in df.columns:
        col_upper = str(col).strip().upper()
        if col_upper in candidatos_nome:
            return col

    melhor_coluna = None
    melhor_score = 0
    for col in df.columns:
        try:
            score = int(df[col].astype(str).str.upper().str.contains(r'\bE\d[A-Z0-9]{8,}\b', na=False, regex=True).sum())
        except Exception:
            score = 0
        if score > melhor_score:
            melhor_score = score
            melhor_coluna = col

    if melhor_score > 0:
        return melhor_coluna
    return None


def _processar_upload_auditoria_unificada(arquivo_upload):
    try:
        df_upload = pd.read_excel(
            arquivo_upload,
            header=None,
            skiprows=5,
            converters={
                0: str,
                1: str,
                4: str,
                5: str,
                6: str,
                7: str
            }
        )
    except Exception as e:
        raise ValueError(f'Erro ao ler o Excel: {str(e)}')

    df_upload = _normalizar_colunas_numericas(df_upload)
    df_upload = df_upload.dropna(how='all')
    if 0 in df_upload.columns:
        df_upload = df_upload[df_upload[0].notna()]

    if df_upload.empty:
        raise ValueError('A planilha enviada não possui linhas válidas para auditoria.')

    col_pi = _detectar_coluna_pi_upload(df_upload)
    if col_pi is None:
        raise ValueError('Não foi possível identificar a coluna PI na planilha enviada.')

    serie_pi = df_upload[col_pi].fillna('').astype(str).str.upper().str.strip()
    mask_pdr = serie_pi.str.contains('PLJ', na=False)
    mask_extra = serie_pi.str.contains('SOL', na=False)

    df_pdr = df_upload[mask_pdr].copy()
    df_extra = df_upload[mask_extra].copy()

    caminhos = _caminhos_upload_auditoria()
    try:
        arquivo_upload.stream.seek(0)
    except Exception:
        pass
    arquivo_upload.save(caminhos['base'])

    df_extra.to_csv(caminhos['extra'], index=False, encoding='utf-8-sig')
    df_pdr.to_csv(caminhos['pdr'], index=False, encoding='utf-8-sig')

    resumo = {
        'total': int(len(df_upload)),
        'pdr': int(len(df_pdr)),
        'extra': int(len(df_extra)),
        'col_pi': str(col_pi),
        'uploaded_at': datetime.now().isoformat(timespec='seconds')
    }
    try:
        with open(caminhos['meta'], 'w', encoding='utf-8') as f:
            json.dump(resumo, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    return resumo


def _processar_upload_auditoria_sisnc(arquivo_upload):
    df_upload = _detectar_dataframe_sisnc_upload(arquivo_upload)

    if df_upload.empty:
        raise ValueError('A planilha SISNC enviada não possui linhas válidas para auditoria.')

    col_pi = _encontrar_coluna_upload(df_upload, [
        'PI',
        'Plano Interno',
        'Projeto/Atividade',
        'Projeto Atividade',
        'PI/Localizador'
    ])
    if col_pi is None:
        col_pi = _inferir_coluna_pi_sisnc_por_conteudo(df_upload)

    col_nd = _encontrar_coluna_upload(df_upload, ['ND', 'Natureza da Despesa', 'Natureza Despesa'])
    col_codug = _encontrar_coluna_upload(df_upload, [
        'Código UG', 'Codigo UG', 'COD UG', 'CODUG',
        'UG Emitente', 'UG Emitente Cod', 'UG Código'
    ])
    col_doc = _encontrar_coluna_upload(df_upload, [
        'Finalidade', 'Finalidade/Observação', 'Finalidade/Observacao',
        'Finalidade Observação', 'Finalidade Observacao',
        'Doc - Observação', 'Doc - Observacao', 'Doc Observacao',
        'Observação', 'Observacao', 'Historico'
    ])
    col_valor = _encontrar_coluna_upload(df_upload, ['Valor', 'Valor Total', 'Valor R$', 'Valor (R$)', 'Vlr'])
    col_nc = _encontrar_coluna_upload(df_upload, ['NC SIAFI', 'NC_SIAFI', 'NCSIAFI'])
    col_ref = _encontrar_coluna_upload(df_upload, ['Referência', 'Referencia', 'Número Referência', 'Numero Referencia', 'Referencia SISNC'])

    if col_pi is None:
        colunas_detectadas = ', '.join([str(c) for c in df_upload.columns[:20]])
        raise ValueError(
            'Não foi possível identificar a coluna PI na planilha SISNC enviada. '
            f'Colunas detectadas: {colunas_detectadas}'
        )
    if col_nd is None or col_codug is None or col_doc is None or col_valor is None:
        raise ValueError('Não foi possível identificar todas as colunas obrigatórias da SISNC (ND, Código UG, Finalidade e Valor).')

    registros = []
    for _, row in df_upload.iterrows():
        pi_norm = normalizar_pi(row.get(col_pi, ''))
        if not pi_norm:
            continue

        nd_norm = normalizar_nd(row.get(col_nd, ''))
        codug_norm = normalizar_cod_ug(row.get(col_codug, ''))
        doc_texto = str(row.get(col_doc, '') or '').strip()
        valor_num = normalizar_valor(row.get(col_valor, 0))

        nc_raw = row.get(col_nc, '') if col_nc else ''
        nc_siafi = re.sub(r'\D', '', str(nc_raw or ''))
        if nc_siafi:
            nc_siafi = nc_siafi[-6:]

        ref_raw = row.get(col_ref, '') if col_ref else ''
        ref_sisnc = str(ref_raw or '').strip()

        registros.append({
            0: nc_siafi,
            1: codug_norm,
            2: ref_sisnc,
            3: nd_norm,
            5: pi_norm,
            6: doc_texto,
            7: valor_num
        })

    df_normalizado = pd.DataFrame(registros)
    if df_normalizado.empty:
        raise ValueError('A planilha SISNC não possui linhas com PI válido para auditoria.')

    serie_pi = df_normalizado[5].fillna('').astype(str).str.upper().str.strip()
    mask_pdr = serie_pi.str.contains('PLJ', na=False)
    mask_extra = serie_pi.str.contains('SOL', na=False)

    df_pdr = df_normalizado[mask_pdr].copy()
    df_extra = df_normalizado[mask_extra].copy()

    caminhos = _caminhos_upload_auditoria()
    try:
        arquivo_upload.stream.seek(0)
    except Exception:
        pass
    arquivo_upload.save(caminhos['sisnc_base'])

    df_extra.to_csv(caminhos['sisnc_extra'], index=False, encoding='utf-8-sig')
    df_pdr.to_csv(caminhos['sisnc_pdr'], index=False, encoding='utf-8-sig')

    resumo = {
        'total': int(len(df_normalizado)),
        'pdr': int(len(df_pdr)),
        'extra': int(len(df_extra)),
        'col_pi': str(col_pi),
        'col_ref': str(col_ref) if col_ref is not None else '',
        'uploaded_at': datetime.now().isoformat(timespec='seconds')
    }
    try:
        with open(caminhos['sisnc_meta'], 'w', encoding='utf-8') as f:
            json.dump(resumo, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    return resumo


def _carregar_upload_auditoria_cache(tipo):
    caminhos = _caminhos_upload_auditoria()
    if tipo == 'pdr':
        caminho = caminhos['pdr']
    elif tipo == 'extra':
        caminho = caminhos['extra']
    else:
        raise ValueError('Tipo de cache de auditoria inválido.')

    if not os.path.exists(caminho):
        raise FileNotFoundError('Cache de upload da auditoria não encontrado.')

    df = pd.read_csv(caminho)
    df = _normalizar_colunas_numericas(df)
    df = df.dropna(how='all')
    if 0 in df.columns:
        df = df[df[0].notna()]
    df['__fonte_upload'] = 'Tesouro Gerencial'
    return df


def _carregar_upload_auditoria_sisnc_cache(tipo):
    caminhos = _caminhos_upload_auditoria()
    if tipo == 'pdr':
        caminho = caminhos['sisnc_pdr']
    elif tipo == 'extra':
        caminho = caminhos['sisnc_extra']
    else:
        raise ValueError('Tipo de cache SISNC inválido.')

    if not os.path.exists(caminho):
        raise FileNotFoundError('Cache SISNC da auditoria não encontrado.')

    df = pd.read_csv(caminho)
    df = _normalizar_colunas_numericas(df)
    df = df.dropna(how='all')
    if 0 in df.columns:
        df = df[df[0].notna() | df[5].notna()]
    df['__fonte_upload'] = 'SISNC'
    return df


def _carregar_upload_auditoria_cache_combinado(tipo):
    dfs = []
    try:
        dfs.append(_carregar_upload_auditoria_cache(tipo))
    except FileNotFoundError:
        pass

    try:
        dfs.append(_carregar_upload_auditoria_sisnc_cache(tipo))
    except FileNotFoundError:
        pass

    if not dfs:
        raise FileNotFoundError('Nenhuma base de upload da auditoria encontrada (TG ou SISNC).')

    if len(dfs) == 1:
        return dfs[0]

    combinado = pd.concat(dfs, ignore_index=True, sort=False)
    combinado = _normalizar_colunas_numericas(combinado)
    combinado = combinado.dropna(how='all')
    return combinado


def _normalizar_numero_solicitacao_pdr(valor):
    texto = _normalizar_texto_planilha(valor)
    return re.sub(r'\s+', '', texto).upper()


def _normalizar_om_chave_cache_pdr(valor):
    texto = _normalizar_texto_planilha(valor)
    texto = re.sub(r'[^A-Z0-9]+', ' ', texto.upper()).strip()
    return re.sub(r'\s+', ' ', texto)


def _chave_cache_nc_pdr(codom, codug, pi, nd, valor):
    """Compatibilidade: usa a nova função de chave de correspondência."""
    return gerar_chave_correspondencia(codom, codug, pi, nd, valor)


def _decompor_cache_key_nc_pdr(cache_key):
    chave = str(cache_key or '')
    if '|' not in chave:
        return {
            'tipo': 'legacy',
            'cache_key': chave,
            'numero': _normalizar_numero_solicitacao_pdr(chave),
            'codom': '',
            'codug': '',
            'pi': '',
            'nd': '',
            'valor': '',
            'om_chave': ''
        }

    partes = chave.split('|')
    if len(partes) >= 5:
        return {
            'tipo': 'composite',
            'cache_key': chave,
            'numero': '',
            'codom': partes[0],
            'codug': partes[1],
            'pi': partes[2],
            'nd': partes[3],
            'valor': partes[4],
            'om_chave': ''
        }

    while len(partes) < 3:
        partes.append('')
    return {
        'tipo': 'legacy-piped',
        'cache_key': chave,
        'numero': _normalizar_numero_solicitacao_pdr(partes[0]),
        'codom': '',
        'codug': partes[1],
        'pi': '',
        'nd': '',
        'valor': '',
        'om_chave': partes[2]
    }


def _cache_key_para_partes_pdr(cache_key):
    partes = _decompor_cache_key_nc_pdr(cache_key)
    return partes.get('numero', ''), partes.get('codug', ''), partes.get('om_chave', '')


def _caminho_cache_nc_pdr():
    os.makedirs(app.instance_path, exist_ok=True)
    return os.path.join(app.instance_path, PDR_LOG_NC_CACHE_FILE)


def _carregar_cache_nc_pdr():
    caminho = _caminho_cache_nc_pdr()
    if not os.path.exists(caminho):
        return {}

    mapa = {}
    try:
        with open(caminho, mode='r', encoding='utf-8-sig', newline='') as arquivo:
            leitor = csv.DictReader(arquivo)
            for row in leitor:
                cache_key_raw = str(row.get('cache_key', '') or '').strip()
                numero_raw = row.get('numero_solicitacao', '')
                nc_raw = row.get('nc_siafi', '')
                codug_raw = row.get('codug', '')
                om_raw = row.get('om_chave', '')
                nc_siafi = re.sub(r'\D', '', str(nc_raw or ''))
                if not nc_siafi:
                    continue

                if cache_key_raw:
                    mapa[cache_key_raw] = nc_siafi
                    continue

                # Recupera CSV legado salvo no formato truncado: codom em numero_solicitacao,
                # codug em codug e pi|nd|valor em om_chave.
                om_partes = str(om_raw or '').split('|') if '|' in str(om_raw or '') else []
                if om_partes:
                    while len(om_partes) < 3:
                        om_partes.append('')
                    chave_legada = _chave_cache_nc_pdr(
                        codom=numero_raw,
                        codug=codug_raw,
                        pi=om_partes[0],
                        nd=om_partes[1],
                        valor=om_partes[2]
                    )
                    if chave_legada:
                        mapa[chave_legada] = nc_siafi
                        continue

                chave = _chave_cache_nc_pdr(
                    codom=row.get('codom', ''),
                    codug=codug_raw,
                    pi=row.get('pi', ''),
                    nd=row.get('nd', ''),
                    valor=row.get('valor', '')
                )
                if chave:
                    mapa[chave] = nc_siafi
                    continue

                numero = _normalizar_numero_solicitacao_pdr(numero_raw)
                if numero:
                    mapa[numero] = nc_siafi
    except Exception:
        return {}

    return mapa


def _salvar_cache_nc_pdr(mapa_nc):
    caminho = _caminho_cache_nc_pdr()
    with open(caminho, mode='w', encoding='utf-8-sig', newline='') as arquivo:
        writer = csv.DictWriter(
            arquivo,
            fieldnames=['cache_key', 'numero_solicitacao', 'codom', 'codug', 'pi', 'nd', 'valor', 'om_chave', 'nc_siafi']
        )
        writer.writeheader()
        for cache_key in sorted(mapa_nc.keys()):
            nc_siafi = re.sub(r'\D', '', str(mapa_nc[cache_key] or ''))
            if nc_siafi:
                partes = _decompor_cache_key_nc_pdr(cache_key)
                writer.writerow({
                    'cache_key': str(cache_key or ''),
                    'numero_solicitacao': partes.get('numero', ''),
                    'codom': partes.get('codom', ''),
                    'codug': partes.get('codug', ''),
                    'pi': partes.get('pi', ''),
                    'nd': partes.get('nd', ''),
                    'valor': partes.get('valor', ''),
                    'om_chave': partes.get('om_chave', ''),
                    'nc_siafi': nc_siafi
                })


def _limpar_conflitos_nc_siafi_pdr(mapa_nc, nc_siafi_alvo, numeros_alvo, base_por_numero, codug_set_alvo=None, om_chave_alvo=''):
    nc_limpa = re.sub(r'\D', '', str(nc_siafi_alvo or ''))
    if not nc_limpa:
        return 0

    codug_set_norm = set(codug_set_alvo or set())
    numeros_preservados = {n for n in (numeros_alvo or set()) if n}
    removidos = 0

    def _om_chave_local(valor):
        texto = _normalizar_texto_planilha(valor)
        texto = re.sub(r'[^A-Z0-9]+', ' ', texto.upper()).strip()
        return re.sub(r'\s+', ' ', texto)

    for cache_key, nc_atual in list(mapa_nc.items()):
        if re.sub(r'\D', '', str(nc_atual or '')) != nc_limpa:
            continue

        numero, codug_key, om_key = _cache_key_para_partes_pdr(cache_key)

        info_base = base_por_numero.get(numero, {})
        codug_numero = codug_key or normalizar_cod_ug(info_base.get('codug', ''))
        om_chave_numero = om_key or _om_chave_local(info_base.get('om_solicitante', ''))

        if numero in numeros_preservados:
            if codug_set_norm and codug_numero:
                if codug_numero in codug_set_norm:
                    continue
            elif om_chave_alvo and om_chave_numero and om_chave_numero == om_chave_alvo:
                continue
            elif not codug_set_norm and not om_chave_alvo:
                continue

        conflita = False
        if codug_set_norm and codug_numero:
            conflita = codug_numero not in codug_set_norm
        elif om_chave_alvo and om_chave_numero:
            conflita = om_chave_numero != om_chave_alvo

        if conflita:
            del mapa_nc[cache_key]
            removidos += 1

    return removidos


def _sanitizar_cache_nc_pdr_unicidade(mapa_nc, base_por_numero):
    def _om_chave_local(valor):
        texto = _normalizar_texto_planilha(valor)
        texto = re.sub(r'[^A-Z0-9]+', ' ', texto.upper()).strip()
        return re.sub(r'\s+', ' ', texto)

    grupos_por_nc = {}
    for cache_key, nc in list(mapa_nc.items()):
        numero_norm, codug_key, om_key = _cache_key_para_partes_pdr(cache_key)
        if not numero_norm:
            continue

        nc_limpa = re.sub(r'\D', '', str(nc or ''))
        if not nc_limpa:
            continue

        base_info = base_por_numero.get(numero_norm, {})
        codug = codug_key or normalizar_cod_ug(base_info.get('codug', ''))
        om_norm = om_key or _om_chave_local(base_info.get('om_solicitante', ''))
        chave_unicidade = f"CODUG:{codug}" if codug else f"OM:{om_norm}"

        grupos_por_nc.setdefault(nc_limpa, {}).setdefault(chave_unicidade, set()).add(cache_key)

    removidos = 0
    for nc, grupos in grupos_por_nc.items():
        if len(grupos) <= 1:
            continue
        for keys in grupos.values():
            for cache_key in keys:
                if cache_key in mapa_nc and re.sub(r'\D', '', str(mapa_nc.get(cache_key, '') or '')) == nc:
                    del mapa_nc[cache_key]
                    removidos += 1

    return removidos


def _normalizar_texto_planilha(valor):
    if pd.isna(valor):
        return ''
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def _carregar_solicitacoes_pdr_log_planilha():
    caminho = app.config.get('PDR_LOG_EXCEL_PATH', 'PLANILHA PDR LOG FINAL.xlsx')
    caminho_absoluto = caminho if os.path.isabs(caminho) else os.path.join(app.root_path, caminho)

    if not os.path.exists(caminho_absoluto):
        raise FileNotFoundError(f"Planilha não encontrada: {caminho_absoluto}")

    mtime = os.path.getmtime(caminho_absoluto)
    if PDR_LOG_CACHE['mtime'] == mtime and PDR_LOG_CACHE['data']:
        return PDR_LOG_CACHE['data']

    excel = pd.ExcelFile(caminho_absoluto)
    aba_bd = next((aba for aba in excel.sheet_names if str(aba).strip().lower() == 'bd'), None)
    if not aba_bd:
        raise ValueError("A aba 'BD' não foi encontrada na planilha PLANILHA PDR LOG FINAL.xlsx")

    df = pd.read_excel(caminho_absoluto, sheet_name=aba_bd)
    df.columns = [str(col).strip() for col in df.columns]

    colunas_esperadas = [
        'NR_SOL', 'DT_SOL', 'ESC_REGINAL', 'OM_SOLICITANTE', 'CLASSE',
        'COD_ASSUNTO', 'NOME_ASSUNTO', 'FINALIDADE_JUSTIFICATIVA',
        'VALOR_TOTAL', 'ATENDIDO_GESTOR', 'CODOM', 'CODUG', 'COD_ND'
    ]

    for coluna in colunas_esperadas:
        if coluna not in df.columns:
            df[coluna] = ''

    df['NR_SOL'] = df['NR_SOL'].apply(_normalizar_texto_planilha)

    df['VALOR_TOTAL'] = pd.to_numeric(df['VALOR_TOTAL'], errors='coerce').fillna(0)
    df['DT_SOL'] = pd.to_datetime(df['DT_SOL'], errors='coerce')

    df['data_solicitacao_fmt'] = df['DT_SOL'].dt.strftime('%d/%m/%Y')
    df['data_solicitacao_fmt'] = df['data_solicitacao_fmt'].fillna('--')

    mapa_nc_siafi = _carregar_cache_nc_pdr()
    # Indexa a tabela de OMs uma unica vez para evitar custo O(n*m) no carregamento.
    base_oms = _obter_oms_data()
    indice_om_exato = {}
    indice_om_normalizado = {}
    for om_item in base_oms:
        om_nome = str(om_item.get('OM', '') or '').strip()
        sigla_nome = str(om_item.get('SIGLA_UG', '') or '').strip()
        codom_ref = re.sub(r'\D', '', str(om_item.get('CODOM', '') or '').strip())
        codug_ref = re.sub(r'\D', '', str(om_item.get('CODUG', '') or '').strip())
        payload = {
            'codom': codom_ref,
            'codug': codug_ref,
            'sigla_ug': sigla_nome,
            'fonte': 'tabela_oms'
        }

        if om_nome:
            indice_om_exato.setdefault(om_nome.upper(), payload)
            chave = normalizar_om_solicitante_chave(om_nome)
            if chave:
                indice_om_normalizado.setdefault(chave, payload)
        if sigla_nome:
            indice_om_exato.setdefault(sigla_nome.upper(), payload)
            chave_sigla = normalizar_om_solicitante_chave(sigla_nome)
            if chave_sigla:
                indice_om_normalizado.setdefault(chave_sigla, payload)

    def _resolver_om_local(om_linha, codom_linha='', codug_linha=''):
        om_txt = str(om_linha or '').strip()
        if om_txt:
            ref_exata = indice_om_exato.get(om_txt.upper())
            if ref_exata:
                return ref_exata
            chave = normalizar_om_solicitante_chave(om_txt)
            if chave and chave in indice_om_normalizado:
                return indice_om_normalizado[chave]
        return {
            'codom': re.sub(r'\D', '', str(codom_linha or '').strip()),
            'codug': re.sub(r'\D', '', str(codug_linha or '').strip()),
            'sigla_ug': '',
            'fonte': 'form_fallback'
        }

    base_por_numero = {}
    contagem_por_numero = {}
    legado_para_chaves_novas = {}
    for _, linha in df.iterrows():
        numero_solicitacao = _normalizar_texto_planilha(linha.get('NR_SOL', ''))
        numero_chave = _normalizar_numero_solicitacao_pdr(numero_solicitacao)
        if not numero_chave:
            continue

        om_linha = _normalizar_texto_planilha(linha.get('OM_SOLICITANTE', ''))
        codom_linha = _normalizar_texto_planilha(linha.get('CODOM', ''))
        codug_linha = _normalizar_texto_planilha(linha.get('CODUG', ''))
        correspondencia_om = _resolver_om_local(om_linha, codom_linha, codug_linha)
        codom_corrigido = _normalizar_texto_planilha(correspondencia_om.get('codom', codom_linha))
        codug_corrigido = _normalizar_texto_planilha(correspondencia_om.get('codug', codug_linha))
        pi_mapeada = normalizar_pi(_mapear_pi_pdr_por_assunto(linha.get('NOME_ASSUNTO', '')))
        nd_mapeada = normalizar_nd(linha.get('COD_ND', ''))
        valor_total = linha.get('VALOR_TOTAL', 0)

        contagem_por_numero[numero_chave] = contagem_por_numero.get(numero_chave, 0) + 1
        base_por_numero[numero_chave] = {
            'om_solicitante': om_linha,
            'codom': codom_corrigido,
            'codug': codug_corrigido,
            'pi': pi_mapeada,
            'nd': nd_mapeada,
            'valor_total': valor_total
        }

        chave_nova = _chave_cache_nc_pdr(
            codom=codom_corrigido,
            codug=codug_corrigido,
            pi=pi_mapeada,
            nd=nd_mapeada,
            valor=valor_total
        )
        if chave_nova:
            chave_legada_linha = _chave_cache_nc_pdr(
                codom=codom_linha,
                codug=codug_corrigido,
                pi='',
                nd='',
                valor=valor_total
            )
            if chave_legada_linha:
                legado_para_chaves_novas.setdefault(chave_legada_linha, set()).add(chave_nova)

            chave_legada_corrigida = _chave_cache_nc_pdr(
                codom=codom_corrigido,
                codug=codug_corrigido,
                pi='',
                nd='',
                valor=valor_total
            )
            if chave_legada_corrigida:
                legado_para_chaves_novas.setdefault(chave_legada_corrigida, set()).add(chave_nova)

    cache_migrado = False
    for cache_key in list(mapa_nc_siafi.keys()):
        if '|' in str(cache_key):
            continue
        numero_legacy = _normalizar_numero_solicitacao_pdr(cache_key)
        if not numero_legacy:
            del mapa_nc_siafi[cache_key]
            cache_migrado = True
            continue

        if contagem_por_numero.get(numero_legacy, 0) == 1:
            info_numero = base_por_numero.get(numero_legacy, {})
            chave_nova = _chave_cache_nc_pdr(
                codom=info_numero.get('codom', ''),
                codug=info_numero.get('codug', ''),
                pi=info_numero.get('pi', ''),
                nd=info_numero.get('nd', ''),
                valor=info_numero.get('valor_total', 0)
            )
            if chave_nova:
                mapa_nc_siafi[chave_nova] = mapa_nc_siafi[cache_key]

        del mapa_nc_siafi[cache_key]
        cache_migrado = True

    for cache_key in list(mapa_nc_siafi.keys()):
        partes_cache = _decompor_cache_key_nc_pdr(cache_key)
        if partes_cache.get('tipo') != 'composite':
            continue
        if partes_cache.get('pi') or partes_cache.get('nd'):
            continue

        chaves_destino = legado_para_chaves_novas.get(str(cache_key), set())
        if len(chaves_destino) != 1:
            continue

        chave_destino = next(iter(chaves_destino))
        nc_valor = mapa_nc_siafi.get(cache_key, '')
        if chave_destino and nc_valor:
            mapa_nc_siafi[chave_destino] = nc_valor
            del mapa_nc_siafi[cache_key]
            cache_migrado = True

    removidos_conflito = _sanitizar_cache_nc_pdr_unicidade(mapa_nc_siafi, base_por_numero)
    if removidos_conflito or cache_migrado:
        _salvar_cache_nc_pdr(mapa_nc_siafi)

    dados = []
    for _, linha in df.iterrows():
        numero_solicitacao = _normalizar_texto_planilha(linha.get('NR_SOL', ''))
        numero_chave = _normalizar_numero_solicitacao_pdr(numero_solicitacao)
        status_original = _normalizar_texto_planilha(linha.get('ATENDIDO_GESTOR', ''))
        om_linha = _normalizar_texto_planilha(linha.get('OM_SOLICITANTE', ''))
        codom_linha = _normalizar_texto_planilha(linha.get('CODOM', ''))
        codug_linha = _normalizar_texto_planilha(linha.get('CODUG', ''))

        correspondencia_om = _resolver_om_local(om_linha, codom_linha, codug_linha)
        codom_corrigido = _normalizar_texto_planilha(correspondencia_om.get('codom', codom_linha))
        codug_corrigido = _normalizar_texto_planilha(correspondencia_om.get('codug', codug_linha))

        pi_chave_cache = normalizar_pi(_mapear_pi_pdr_por_assunto(linha.get('NOME_ASSUNTO', '')))
        nd_chave_cache = normalizar_nd(linha.get('COD_ND', ''))
        chave_cache = _chave_cache_nc_pdr(
            codom=codom_corrigido,
            codug=codug_corrigido,
            pi=pi_chave_cache,
            nd=nd_chave_cache,
            valor=linha.get('VALOR_TOTAL', 0)
        )
        nc_siafi_cache = mapa_nc_siafi.get(chave_cache, '')
        if not nc_siafi_cache and contagem_por_numero.get(numero_chave, 0) == 1:
            nc_siafi_cache = mapa_nc_siafi.get(numero_chave, '')
        if nc_siafi_cache:
            status_exibicao = 'NC SIAFI gerada'
        elif status_original.lower() == 'sim':
            status_exibicao = 'gerar NC'
        else:
            status_exibicao = status_original or 'Não informado'
        dados.append({
            'numero': numero_solicitacao,
            'data_solicitacao': linha.get('data_solicitacao_fmt', '--'),
            'orgao_demandante': _normalizar_texto_planilha(linha.get('ESC_REGINAL', '')),
            'om_solicitante': _normalizar_texto_planilha(linha.get('OM_SOLICITANTE', '')),
            'classe': _normalizar_texto_planilha(linha.get('CLASSE', '')),
            'cod_assunto': _normalizar_texto_planilha(linha.get('COD_ASSUNTO', '')),
            'nome_assunto': _normalizar_texto_planilha(linha.get('NOME_ASSUNTO', '')),
            'cod_nd': normalizar_nd(linha.get('COD_ND', '')),
            'codom': codom_corrigido,
            'codug': codug_corrigido,
            'descricao': _normalizar_texto_planilha(linha.get('FINALIDADE_JUSTIFICATIVA', '')),
            'valor_total': float(linha.get('VALOR_TOTAL', 0) or 0),
            'status': status_exibicao,
            'status_original': status_original,
            'nc_siafi': nc_siafi_cache,
        })

    dados.sort(key=lambda item: item['numero'], reverse=True)
    # Enriquecer com informação de 'Atendido' consultando NCs persistidas
    try:
        from models import SolicitacaoExtraPDRLOG
        # Construir mapa de nº solicitacao (normalizado) -> objeto DB
        solicitacoes_db = { _normalizar_numero_solicitacao_pdr(s.numero): s for s in SolicitacaoExtraPDRLOG.query.all() }
        for item in dados:
            numero_norm = _normalizar_numero_solicitacao_pdr(item.get('numero', ''))
            sol_db = solicitacoes_db.get(numero_norm)
            if not sol_db:
                continue
            # coletar NCs vinculadas que possuam ref_sisnc ou nc_siafi
            nc_siafi_vals = []
            ref_sisnc_vals = []
            for nc in getattr(sol_db, 'notas_credito', []) or []:
                if (nc.nc_siafi or '').strip():
                    nc_siafi_vals.append((nc.numero or '') + (f" ({nc.nc_siafi})" if nc.nc_siafi else ''))
                if (nc.ref_sisnc or '').strip():
                    ref_sisnc_vals.append((nc.numero or '') + (f" ({nc.ref_sisnc})" if nc.ref_sisnc else ''))

            if nc_siafi_vals or ref_sisnc_vals:
                # Marcar como Atendido e incluir informações para exibição
                detalhes = []
                if nc_siafi_vals:
                    detalhes.append('NC SIAFI: ' + ', '.join(nc_siafi_vals))
                if ref_sisnc_vals:
                    detalhes.append('REF SISNC: ' + ', '.join(ref_sisnc_vals))
                item['status'] = 'Atendido'
                item['status_detalhe'] = ' | '.join(detalhes)
    except Exception:
        # se houve erro ao consultar DB, deixar dados como estavam
        pass
    PDR_LOG_CACHE['mtime'] = mtime
    PDR_LOG_CACHE['data'] = dados
    return dados


def _mapear_pi_pdr_por_assunto(nome_assunto):
    assunto = str(nome_assunto or '')
    assunto_norm = unicodedata.normalize('NFD', assunto)
    assunto_norm = ''.join(ch for ch in assunto_norm if unicodedata.category(ch) != 'Mn').upper()
    assunto_norm = re.sub(r'\s+', ' ', assunto_norm)

    if 'PASA' in assunto_norm:
        return 'E6SUPLJA5PA'
    if ('MNT INSTALACOES ST APROV' in assunto_norm) or (('MNT' in assunto_norm or 'MANUTEN' in assunto_norm) and 'INSTAL' in assunto_norm and 'ST APROV' in assunto_norm):
        return 'E6SUPLJA7PA'
    if ('MNT INSTALACOES OP' in assunto_norm) or (('MNT' in assunto_norm or 'MANUTEN' in assunto_norm) and 'INSTAL' in assunto_norm and 'OP' in assunto_norm):
        return 'E6SUPLJA8OP'
    if 'MNT OP' in assunto_norm or (('MNT' in assunto_norm or 'MANUTEN' in assunto_norm) and 'OP' in assunto_norm):
        return 'E6SUPLJA6OP'
    return ''


def _agrupar_notas_credito_pdr(solicitacoes=None, incluir_solicitacoes=False, somente_status_sim=False):
    if solicitacoes is None:
        solicitacoes = _carregar_solicitacoes_pdr_log_planilha()

    agrupado = {}
    # Carrega tabela de correspondência OM x CODOM x CODUG
    tabela_oms = {str(om.get('CODOM', '')).strip(): om for om in _carregar_tabela_oms() if om.get('CODOM', '')}

    for solicitacao in solicitacoes:
        if somente_status_sim:
            status_original = _normalizar_texto_planilha(solicitacao.get('status_original', '')).strip().lower()
            if status_original != 'sim':
                continue

        pi_mapeada = _mapear_pi_pdr_por_assunto(solicitacao.get('nome_assunto', ''))
        nd_mapeada = normalizar_nd(solicitacao.get('cod_nd', ''))
        if not nd_mapeada or not pi_mapeada:
            continue

        om = str(solicitacao.get('om_solicitante', '') or '').strip()
        if not om:
            continue

        valor = float(solicitacao.get('valor_total', 0) or 0)
        nc_siafi = re.sub(r'\D', '', str(solicitacao.get('nc_siafi', '') or ''))
        codom = _normalizar_texto_planilha(solicitacao.get('codom', ''))
        # Buscar CODUG correto pela tabela de correspondência
        codug = ''
        if codom and codom in tabela_oms:
            codug = str(tabela_oms[codom].get('CODUG', '')).strip()
        else:
            codug = _normalizar_texto_planilha(solicitacao.get('codug', ''))

        chave = (om, pi_mapeada, nd_mapeada)
        item = agrupado.setdefault(chave, {
            'om_solicitante': om,
            'pi': pi_mapeada,
            'nd': nd_mapeada,
            'valor_total': 0.0,
            'quantidade_solicitacoes': 0,
            'ncs_siafi': set(),
            'codom_set': set(),
            'codug_set': set(),
            'solicitacoes': []
        })
        item['valor_total'] += valor
        item['quantidade_solicitacoes'] += 1
        if nc_siafi:
            item['ncs_siafi'].add(nc_siafi)
        if codom:
            item['codom_set'].add(codom)
        if codug:
            item['codug_set'].add(codug)
        if incluir_solicitacoes:
            item['solicitacoes'].append(solicitacao)

    grupos = []
    for item in agrupado.values():
        nc_siafi_consolidada = ', '.join(sorted(item['ncs_siafi'])) if item['ncs_siafi'] else ''
        # Determinar status preferindo dados persistidos em NotaCredito quando existirem
        status_determinado = 'NC SIAFI gerada' if nc_siafi_consolidada else 'gerar NC'
        try:
            from models import NotaCredito
            # Buscar qualquer NotaCredito que corresponda a este grupo (mesmo PI e ND)
            nc_db = db.session.query(NotaCredito).filter(
                NotaCredito.pi == item['pi'],
                NotaCredito.nd == item['nd']
            ).filter(
                (NotaCredito.nc_siafi != '') | (NotaCredito.ref_sisnc != '')
            ).first()

            if nc_db:
                if (nc_db.nc_siafi or '').strip():
                    status_determinado = 'NC SIAFI gerada'
                elif (nc_db.ref_sisnc or '').strip():
                    status_determinado = 'Cadastrada SISNC'
        except Exception:
            # Em caso de qualquer erro no acesso ao DB, manter o status calculado a partir da planilha
            pass

        grupos.append({
            'om_solicitante': item['om_solicitante'],
            'pi': item['pi'],
            'nd': item['nd'],
            'codom': ', '.join(sorted(item['codom_set'])) if item['codom_set'] else '--',
            'codug': ', '.join(sorted(item['codug_set'])) if item['codug_set'] else '--',
            'valor_total': round(item['valor_total'], 2),
            'quantidade_solicitacoes': item['quantidade_solicitacoes'],
            'nc_siafi': nc_siafi_consolidada,
            'status': status_determinado,
            'solicitacoes': item['solicitacoes'] if incluir_solicitacoes else []
        })

    return grupos


def _deduplicar_solicitacoes_para_nc_pdr(solicitacoes):
    """Remove duplicidades por OM+PI+ND+VALOR para montagem da lista de NCs."""
    if not solicitacoes:
        return [], {'total_entrada': 0, 'total_saida': 0, 'duplicadas_removidas': 0}

    melhores_por_chave = {}
    ordem_chaves = []

    for solicitacao in solicitacoes:
        om_display = str(solicitacao.get('om_solicitante', '') or '').strip()
        om_chave = normalizar_om_solicitante_chave(om_display)
        pi_mapeada = normalizar_pi(_mapear_pi_pdr_por_assunto(solicitacao.get('nome_assunto', '')))
        nd_mapeada = normalizar_nd(solicitacao.get('cod_nd', ''))
        valor = round(normalizar_valor(solicitacao.get('valor_total', 0)), 2)

        if not om_chave or not pi_mapeada or not nd_mapeada:
            # Sem chave completa, preserva sem deduplicar.
            chave = ('__sem_chave__', id(solicitacao))
            melhores_por_chave[chave] = solicitacao
            ordem_chaves.append(chave)
            continue

        chave = (om_chave, pi_mapeada, nd_mapeada, f"{valor:.2f}")
        numero_norm = _normalizar_numero_solicitacao_pdr(solicitacao.get('numero', ''))
        tem_numero = 1 if numero_norm else 0
        nc_siafi_limpa = re.sub(r'\D', '', str(solicitacao.get('nc_siafi', '') or ''))
        tem_nc_siafi = 1 if nc_siafi_limpa else 0
        prioridade = (tem_numero, tem_nc_siafi)

        atual = melhores_por_chave.get(chave)
        if atual is None:
            melhores_por_chave[chave] = solicitacao
            ordem_chaves.append(chave)
            continue

        numero_atual = _normalizar_numero_solicitacao_pdr(atual.get('numero', ''))
        tem_numero_atual = 1 if numero_atual else 0
        nc_atual_limpa = re.sub(r'\D', '', str(atual.get('nc_siafi', '') or ''))
        tem_nc_atual = 1 if nc_atual_limpa else 0
        prioridade_atual = (tem_numero_atual, tem_nc_atual)

        if prioridade > prioridade_atual:
            melhores_por_chave[chave] = solicitacao

    filtradas = [melhores_por_chave[chave] for chave in ordem_chaves]
    total_entrada = len(solicitacoes)
    total_saida = len(filtradas)
    return filtradas, {
        'total_entrada': total_entrada,
        'total_saida': total_saida,
        'duplicadas_removidas': max(total_entrada - total_saida, 0)
    }

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(Usuario, int(user_id))

# CONSTANTES
FINALIDADES = [
    'Quantitativo de Rancho (QR)',
    'Quantitativo de Subsistência (QS)', 
    'Reserva Regional (RR)',
    'Alimentação em Campanha',
    'PASA',
    'Manutenção de Órgão Provedor',
    'Hub Logístico',
    'Remanejamento',
    'Câmaras frigoríficas',
    'PASA-DEC',
    'Mnt OP-DEC',
    'Publicações',
    'Diárias',
    'Passagens',
    'Solenidades',
    'Outros'
]
ORGAOS_DEMANDANTES = [
    '1ª RM', '2ª RM', '3ª RM', '4ª RM', '5ª RM', '6ª RM', '7ª RM', '8ª RM', '9ª RM', '10ª RM', '11ª RM', '12ª RM',
    'DGP', 'DEC', 'DECEx', 'SEF', 'DCT', 'COLOG', 'COTER', 'GAB CMT EX', 'EME', 'SGEx', 'CIE', 'CCOMSEx',
    # Comandos Militares e Grupos de Logística adicionais
    'CMS', 'CML', 'CMO', 'CMSE', 'CMP', 'CMNE', 'CMN', 'CMA', '3º GPT LOG', '8º GPT LOG', '9º GPT LOG'
]
STATUS_SOLICITACAO = ['Aguardando Análise', 'Em Análise', 'Aguardando despacho', 'Aprovado Ch Sup', 'Aprovado Parcialmente', 'Negado Ch Sup', 'Arquivado', 'Devolvido para correções']
STATUS_NC = ['Não Conferida', 'Pendente', 'Cadastrada SISNC', 'Processada SIAFI', 'Cancelada']
ND_POSSIVEIS = [
    '33.90.00', '33.90.30', '33.90.39', '44.90.00', '44.90.52',
    '33.90.15', '33.90.33', '44.90.51', '44.90.30', '44.90.39', '33.91.39',
    '33.91.47', '33.90.47', '33.90.36'
]

MODALIDADES = ['crédito', 'PRO', 'Transferência Interna COLOG']
DESPACHOS_OPCOES = ['Aprovado integralmente pelo Ch Sup', 'Aprovado parcialmente pelo Ch Sup', 'Negado pelo Ch Sup']

# CONSTANTES ADICIONAIS
STATUS_PRO = ['Aguardando término do processo licitatório', 'Convertida em NC', 'Cancelada']

# Variantes de status com acentuação corrompida (para compatibilidade de filtros)
STATUS_PRO_VARIANTES = {
    'Aguardando t\ufffdrmino do processo licitat\ufffdrio': 'Aguardando término do processo licitatório'
    ,'Aguardando termino do processo licitatorio': 'Aguardando término do processo licitatório'
}

# DICIONÁRIO DE CORRESPONDÊNCIA PI
PI_POR_FINALIDADE = {
    'Quantitativo de Rancho (QR)': 'E6SUSOLA1QR',
    'Quantitativo de Subsistência (QS)': 'E6SUSOLA2QS', 
    'Reserva Regional (RR)': 'E6SUPLJA3RR',  # CORREÇÃO: E6SUPLJA3RR
    # Alimentação em Campanha agora aceita múltiplos PIs, tratar na view/template
    'Alimentação em Campanha': ['E6SUPLJA4QS', 'E6SUSOLA4QR'],
    'PASA': 'E6SUSOLA5PA',
    'Manutenção de Órgão Provedor': 'E6SUSOLA6OP',
    'Hub Logístico': 'E6SUSOLA2QS',
    'Remanejamento': '212BDETAL',
    'Câmaras frigoríficas': 'E6SUSOLA5CF',
    'PASA-DEC': 'E6SUSOLA7PA',
    'Mnt OP-DEC': 'E6SUSOLA8OP',
    'Publicações': 'E6SUSOLPUBL',
    'Diárias': 'E6SUSOLDIAR',
    'Passagens': 'E6SUSOLPASS',
    'Solenidades': 'E6SUSOLSOLE',
    'Outros': 'E6SUSOLOUTROS'
}

# DICIONÁRIO DE DESCRIÇÕES PARA NC BASEADO EM PI E ND
DESCRICOES_NC_POR_PI_ND = {
    'E6SUSOLA5PA': {
        '30': "C SUP-DIV SOL-PASA DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.406-01 (AQS MATERIAL DE CONSUMO ST APROV).",
        '39': "C SUP-DIV SOL-PASA DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.406-01 (SV MANUTENÇÃO ST APROV).",
        '52': "C SUP-DIV SOL-PASA DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.406-01 (AQS MATERIAL PERMANENTE ST APROV)."
    },
    'E6SUSOLA5CF': {
        '30': "C SUP-DIV SOL-PASA DSP EXTRAORDINÁRIA. OBS REGRAS DO PDRLOG-A5 (AQS MATERIAL DE CONSUMO CÂMARA FRIGO – Nº OPUS: {opus}).",
        '39': "C SUP-DIV SOL-PASA DSP EXTRAORDINÁRIA. OBS REGRAS DO PDRLOG-A5 (SV MANUTENÇÃO CÂMARA FRIGO – Nº OPUS: {opus}).",
        '52': "C SUP-DIV SOL-PASA DSP EXTRAORDINÁRIA. OBS REGRAS DO PDRLOG-A5 (AQS MATERIAL PERMANENTE CÂMARA FRIGO – Nº OPUS: {opus})."
    },
    'E6SUPLJA7PA': {
        '30': "C SUP-DIV PLJ-PASA DEC DSP PLANEJADA. OBS REGRAS DO BT30.406-01 (AQS MAT DE CONSUMO P/ ADEQUAÇÃO ST APROV – Nº OPUS: {opus}).",
        '39': "C SUP-DIV PLJ-PASA DEC DSP PLANEJADA. OBS REGRAS DO BT30.406-01 (CONTRATAÇÃO SV P/ ADEQUAÇÃO ST APROV – Nº OPUS: {opus})."
    },
    'E6SUSOLA7PA': {
        '30': "C SUP-DIV SOL-PASA DEC DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.406-01 (AQS MAT DE CONSUMO P/ ADEQUAÇÃO ST APROV – Nº OPUS: {opus}).",
        '39': "C SUP-DIV SOL-PASA DEC DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.406-01 (CONTRATAÇÃO SV P/ ADEQUAÇÃO ST APROV – Nº OPUS: {opus})."
    },
    'E6SUSOLPUBL': {
        '39': "C SUP-DIV SOL-ATENDER SV PUBLICAÇÕES OFICIAIS. OBS LEGISLAÇÃO VIGENTE. (SV DE PUBLICAÇÃO DE EDITAL DE LICITAÇÃO)."
    },
    'E6SUSOLA6OP': {
        '30': "C SUP-DIV SOL-MNT OP DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.409-01 (AQS MATERIAL DE CONSUMO MNT OP).",
        '39': "C SUP-DIV SOL-MNT OP DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.409-01 (SV MANUTENÇÃO MNT OP).",
        '52': "C SUP-DIV SOL-MNT OP DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.409-01 (AQS MATERIAL PERMANENTE MNT OP)."
    },
    'E6SUSOLA8OP': {
        '30': "C SUP-DIV SOL-PASA DEC DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.406-01 (AQS MAT DE CONSUMO P/ ADEQUAÇÃO OP – Nº OPUS: {opus}).",
        '39': "C SUP-DIV SOL-PASA DEC DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.406-01 (CONTRATAÇÃO SV P/ ADEQUAÇÃO OP – Nº OPUS: {opus})."
    },
    'E6SUSOLSOLE': {
        '30': "C SUP-DIV SOL-DPS DE ALIMENTAÇÃO EM CERIMONIAL MILITAR. OBSERVAR O BT30.410-01 (AQS GEN ALMT P/ {descricao_usuario}).",
        '39': "C SUP-DIV SOL-DPS DE ALIMENTAÇÃO EM CERIMONIAL MILITAR. OBSERVAR O BT30.410-01 (SV FORN ALMT P/ {descricao_usuario})."
    },
    'E6SUPLJA3RR': {
        'default': "C SUP-DIV SUBS-PLJ-RES REG DSP ORDINÁRIA. OBSERVAR REGRAS DO BT30.411-01. ({descricao_usuario})"
    },
    'E6SUSOLA1QR': {
        '30': "C SUP-DIV SOL-QUANTITATIVO DE RANCHO DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.403-01 (AQS MATERIAL DE CONSUMO PARA QR).",
        '39': "C SUP-DIV SOL-QUANTITATIVO DE RANCHO DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.403-01 (SV MANUTENÇÃO PARA QR)."
    },
    'E6SUSOLA2QS': {
        '30': "C SUP-DIV SOL-QUANTITATIVO DE SUBSISTÊNCIA DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.403-01 (AQS MATERIAL DE CONSUMO PARA QS).",
        '39': "C SUP-DIV SOL-QUANTITATIVO DE SUBSISTÊNCIA DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.403-01 (SV MANUTENÇÃO PARA QS)."
    },
    'E6SUSOLA4QR': {
        '30': "C SUP-DIV SOL-ALIMENTAÇÃO EM CAMPANHA DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.404-01 (AQS MATERIAL DE CONSUMO PARA AC).",
        '39': "C SUP-DIV SOL-ALIMENTAÇÃO EM CAMPANHA DSP EXTRAORDINÁRIA. OBS REGRAS DO BT30.404-01 (SV MANUTENÇÃO PARA AC)."
    },
    # ADICIONAR DESCRIÇÕES PARA O NOVO ND 33.91.39
    'default': {
        '30': "C SUP-DIV SOL-ATENDER SV PUBLICAÇÕES OFICIAIS. OBS LEGISLAÇÃO VIGENTE. (SV DE PUBLICAÇÃO DE EDITAL DE LICITAÇÃO).",
        '39': "C SUP-DIV SOL-ATENDER SV PUBLICAÇÕES OFICIAIS. OBS LEGISLAÇÃO VIGENTE. (SV DE PUBLICAÇÃO DE EDITAL DE LICITAÇÃO).",
        
    }
}
def criar_usuarios_iniciais():
    """Cria usuários iniciais para o sistema"""
    try:
        print("👤 Criando usuários iniciais...")
        
        # Verificar se já existem usuários
        if Usuario.query.first():
            print("✅ Usuários já existem, pulando criação...")
            return
        
        # Criar usuário administrador (acesso total)
        admin = Usuario(
            nome="Administrador",
            email="admin@sistema.com", 
            departamento="COLOG",
            nivel_acesso="admin"
        )
        admin.set_password("admin123")
        
        # Criar usuário padrão (acesso total)
        usuario = Usuario(
            nome="Usuário Padrão",
            email="usuario@sistema.com",
            departamento="COLOG",
            nivel_acesso="usuario"
        )
        usuario.set_password("user123")
        
        # Criar usuário apenas para Notas de Crédito
        nc_user = Usuario(
            nome="Usuário NC",
            email="ncuser@sistema.com",
            departamento="COLOG",
            nivel_acesso="nc_only"
        )
        nc_user.set_password("nc123")
        
        db.session.add(admin)
        db.session.add(usuario)
        db.session.add(nc_user)
        db.session.commit()
        
        print("✅ Usuários iniciais criados com sucesso!")
        print(f"   👑 Admin: admin@sistema.com / admin123")
        print(f"   👤 Usuário: usuario@sistema.com / user123")
        print(f"   📋 NC Only: ncuser@sistema.com / nc123")
        
    except Exception as e:
        db.session.rollback()
        print(f"❌ Erro ao criar usuários iniciais: {e}")
        raise
def init_database():
    """Inicializa o banco de dados de forma segura, mantendo dados existentes"""
    with app.app_context():
        try:
            import os
            from sqlalchemy import inspect, text
            
            print("🚀 Iniciando inicialização segura do banco...")
            db_path = db.engine.url.database
            if db_path and not os.path.isabs(db_path):
                db_path = os.path.abspath(db_path)
            print(f"🗄️ Banco configurado: {db_path or 'memória/indefinido'}")

            # Verificar se o banco existe
            if not db_path or not os.path.exists(db_path):
                print("📝 Banco não existe. Criando novo banco...")
                if db_path:
                    db_dir = os.path.dirname(db_path)
                    if db_dir:
                        os.makedirs(db_dir, exist_ok=True)
                db.create_all()
                criar_usuarios_iniciais()
                print("🎉 Novo banco criado com sucesso!")
            else:
                print("📝 Banco existente detectado. Verificando estrutura...")
                
                # Criar tabelas se não existirem
                db.create_all()
                
                # Verificar e atualizar estrutura
                inspector = inspect(db.engine)
                
                # Verificar colunas da tabela solicitacao_extra_pdrlog
                try:
                    columns_solicitacao = [col['name'] for col in inspector.get_columns('solicitacao_extra_pdrlog')]
                    
                    # Adicionar coluna tem_fsv se não existir
                    if 'tem_fsv' not in columns_solicitacao:
                        print("⚠️  Adicionando coluna 'tem_fsv' à tabela solicitacao_extra_pdrlog...")
                        with db.engine.connect() as conn:
                            conn.execute(text('ALTER TABLE solicitacao_extra_pdrlog ADD COLUMN tem_fsv VARCHAR(10)'))
                        print("✅ Coluna 'tem_fsv' adicionada")
                    
                    # Adicionar coluna diex_dom se não existir
                    if 'diex_dom' not in columns_solicitacao:
                        print("⚠️  Adicionando coluna 'diex_dom' à tabela solicitacao_extra_pdrlog...")
                        with db.engine.connect() as conn:
                            conn.execute(text('ALTER TABLE solicitacao_extra_pdrlog ADD COLUMN diex_dom VARCHAR(100)'))
                        print("✅ Coluna 'diex_dom' adicionada")

                    if 'nr_opus' not in columns_solicitacao:
                        print("⚠️  Adicionando coluna 'nr_opus' à tabela solicitacao_extra_pdrlog...")
                        with db.engine.connect() as conn:
                            conn.execute(text('ALTER TABLE solicitacao_extra_pdrlog ADD COLUMN nr_opus VARCHAR(100)'))
                        print("✅ Coluna 'nr_opus' adicionada")

                    if 'destinatario_colog' not in columns_solicitacao:
                        print("⚠️  Adicionando coluna 'destinatario_colog' à tabela solicitacao_extra_pdrlog...")
                        with db.engine.connect() as conn:
                            conn.execute(text('ALTER TABLE solicitacao_extra_pdrlog ADD COLUMN destinatario_colog VARCHAR(50)'))
                        print("✅ Coluna 'destinatario_colog' adicionada")
                        
                except Exception as e:
                    print(f"⚠️  Erro ao verificar colunas da solicitacao_extra_pdrlog: {e}")

                # Verificar colunas da tabela pedido_solicitacao
                try:
                    columns_pedido = [col['name'] for col in inspector.get_columns('pedido_solicitacao')]

                    if 'descricao_om' not in columns_pedido:
                        print("⚠️  Adicionando coluna 'descricao_om' à tabela pedido_solicitacao...")
                        with db.engine.connect() as conn:
                            conn.execute(text('ALTER TABLE pedido_solicitacao ADD COLUMN descricao_om TEXT'))
                        print("✅ Coluna 'descricao_om' adicionada")
                except Exception as e:
                    print(f"⚠️  Erro ao verificar colunas da pedido_solicitacao: {e}")

                # Verificar colunas da tabela item_pedido
                try:
                    columns_item = [col['name'] for col in inspector.get_columns('item_pedido')]

                    if 'valor_restante' not in columns_item:
                        print("⚠️  Adicionando coluna 'valor_restante' à tabela item_pedido...")
                        with db.engine.connect() as conn:
                            conn.execute(text('ALTER TABLE item_pedido ADD COLUMN valor_restante FLOAT'))
                        print("✅ Coluna 'valor_restante' adicionada")
                except Exception as e:
                    print(f"⚠️  Erro ao verificar colunas da item_pedido: {e}")

                # Verificar colunas da tabela nota_credito
                try:
                    columns_nc = [col['name'] for col in inspector.get_columns('nota_credito')]

                    if 'item_id' not in columns_nc:
                        print("⚠️  Adicionando coluna 'item_id' à tabela nota_credito...")
                        with db.engine.connect() as conn:
                            conn.execute(text('ALTER TABLE nota_credito ADD COLUMN item_id INTEGER'))
                        print("✅ Coluna 'item_id' adicionada")

                    if 'diex_credito' not in columns_nc:
                        print("⚠️  Adicionando coluna 'diex_credito' à tabela nota_credito...")
                        with db.engine.connect() as conn:
                            conn.execute(text('ALTER TABLE nota_credito ADD COLUMN diex_credito VARCHAR(100)'))
                        print("✅ Coluna 'diex_credito' adicionada")
                except Exception as e:
                    print(f"⚠️  Erro ao verificar colunas da nota_credito: {e}")
                
                # Verificar se existem usuários, se não, criar
                if not Usuario.query.first():
                    criar_usuarios_iniciais()
                    print("✅ Usuários iniciais criados")
                
                print("🎉 Banco verificado e atualizado (dados preservados)!")
            
        except Exception as e:
            print(f"❌ Erro na inicialização do banco: {e}")
            import traceback
            traceback.print_exc()

def construir_descricao_nc(pi, nd, diex, data, om, codom, descricao_usuario="", numero_pro="", descricao_existente="", descricao_om=""):
    """Constrói a descrição da NC baseada em PI e ND, mas preserva descrições editadas manualmente"""
    try:
        # Se já existe uma descrição e não é a padrão, preservar
        if descricao_existente and not descricao_existente.startswith("C SUP-DIVSUBS-SOL DSP EXTRAORDINÁRIA"):
            print(f"📝 Preservando descrição editada pelo usuário")
            return descricao_existente
        
        # Extrair o código do ND (últimos dois dígitos)
        nd_codigo = ""
        if nd and '.' in nd:
            nd_codigo = nd.split('.')[-1]  # Pega os últimos dígitos após o último ponto
            
        print(f"📝 Gerando descrição para PI: {pi}, ND: {nd}, Código ND: {nd_codigo}")
        
        # Verificar se temos descrição específica para este PI e ND
        if pi in DESCRICOES_NC_POR_PI_ND:
            pi_descricoes = DESCRICOES_NC_POR_PI_ND[pi]
            
            # Primeiro tentar com o código específico do ND
            if nd_codigo in pi_descricoes:
                descricao_template = pi_descricoes[nd_codigo]
                print(f"✅ Encontrada descrição específica para PI {pi} e ND {nd_codigo}")
            elif 'default' in pi_descricoes:
                descricao_template = pi_descricoes['default']
                print(f"⚠️  Usando descrição default para PI {pi}")
            else:
                # Se não encontrou específica nem default, verificar no dicionário geral
                if 'default' in DESCRICOES_NC_POR_PI_ND and nd_codigo in DESCRICOES_NC_POR_PI_ND['default']:
                    descricao_template = DESCRICOES_NC_POR_PI_ND['default'][nd_codigo]
                    print(f"⚠️  Usando descrição default geral para ND {nd_codigo}")
                else:
                    descricao_template = "C SUP-DIVSUBS-SOL DSP EXTRAORDINÁRIA. OBSERVAR REGRAS."
                    print(f"⚠️  Nenhuma descrição encontrada, usando padrão")
        else:
            # Verificar no dicionário geral
            if 'default' in DESCRICOES_NC_POR_PI_ND and nd_codigo in DESCRICOES_NC_POR_PI_ND['default']:
                descricao_template = DESCRICOES_NC_POR_PI_ND['default'][nd_codigo]
                print(f"⚠️  Usando descrição default geral para ND {nd_codigo}")
            else:
                descricao_template = "C SUP-DIVSUBS-SOL DSP EXTRAORDINÁRIA. OBSERVAR REGRAS."
                print(f"⚠️  PI {pi} não encontrado, usando descrição padrão")
        
        # Substituir variáveis na descrição
        descricao_final = descricao_template
        
        # Substituir {opus} se houver descrição do usuário para PIs específicos
        # Se for PASA-DEC, Mnt OP-DEC ou Câmaras frigoríficas, e houver OPUS, sempre usar OPUS no parêntesis
        finalidade_especial = pi in ['E6SUPLJA7PA', 'E6SUSOLA7PA', 'E6SUSOLA8OP', 'E6SUSOLA5CF']
        if '{opus}' in descricao_template:
            # Sempre priorizar o campo nr_opus da solicitação se existir
            opus_valor = None
            import inspect
            frame = inspect.currentframe()
            nr_opus = None
            try:
                outer = frame.f_back
                if outer and 'nr_opus' in outer.f_locals:
                    nr_opus = outer.f_locals['nr_opus']
            finally:
                del frame
            if nr_opus and str(nr_opus).strip():
                opus_valor = str(nr_opus).strip()
            elif descricao_om and 'Nr OPUS' in descricao_om:
                match = re.search(r'Nr OPUS[:\s]*([\d\w\-/]+)', descricao_om, re.IGNORECASE)
                if match:
                    opus_valor = match.group(1).strip()
                else:
                    opus_valor = descricao_om.strip()
            elif descricao_usuario and descricao_usuario.strip() and descricao_usuario.strip().lower() != 'a ser informado pelo usuário':
                opus_valor = descricao_usuario.strip()
            elif descricao_om and descricao_om.strip():
                opus_valor = descricao_om.strip()
            # Remover duplicidade: sempre remover "Nr OPUS: <valor>" de descricao_om antes de inserir
            if opus_valor:
                # Remove todas as variantes de OPUS do texto base (incluindo duplicatas) ANTES de substituir {opus}
                padrao_opus_com_valor = r'(–?\s*)?(Nº|Nr) OPUS:?\s*'+re.escape(opus_valor)
                padrao_opus_sem_valor = r'(–?\s*)?(Nº|Nr) OPUS:?\s*'
                while re.search(padrao_opus_com_valor, descricao_final, flags=re.IGNORECASE):
                    descricao_final = re.sub(padrao_opus_com_valor, '', descricao_final, flags=re.IGNORECASE)
                # Remove qualquer 'Nº OPUS:' ou 'Nr OPUS:' sem valor
                descricao_final = re.sub(padrao_opus_sem_valor, '', descricao_final, flags=re.IGNORECASE)
                descricao_om_sem_opus = descricao_om
                if descricao_om:
                    while re.search(padrao_opus_com_valor, descricao_om_sem_opus, flags=re.IGNORECASE):
                        descricao_om_sem_opus = re.sub(padrao_opus_com_valor, '', descricao_om_sem_opus, flags=re.IGNORECASE).strip()
                    descricao_om_sem_opus = re.sub(padrao_opus_sem_valor, '', descricao_om_sem_opus, flags=re.IGNORECASE).strip()
                # Sempre usar formato 'Nr OPUS: <valor>'
                descricao_final = descricao_final.replace('{opus}', f'Nr OPUS:{opus_valor}')
            else:
                descricao_final = descricao_final.replace('{opus}', 'A SER INFORMADO PELO USUÁRIO')
        # Substituir {descricao_usuario} se houver
        if '{descricao_usuario}' in descricao_final and descricao_usuario:
            descricao_final = descricao_final.replace('{descricao_usuario}', descricao_usuario)
        elif '{descricao_usuario}' in descricao_final:
            descricao_final = descricao_final.replace('{descricao_usuario}', 'A SER INFORMADO PELO USUÁRIO')
        
        # Prefixo exigido: "(CODOM-OM) C SUP-DIV SUBS " (COD/OM primeiro)
        if codom and om:
            prefixo = f"({codom}-{om}) C SUP-DIV SUBS "
        elif om:
            prefixo = f"({om}) C SUP-DIV SUBS "
        elif codom:
            prefixo = f"({codom}) C SUP-DIV SUBS "
        else:
            prefixo = "C SUP-DIV SUBS "

        # Remover prefixos antigos "C SUP-DIV" ou "C SUP-DIVSUBS" do template
        desc_sem_prefixo = descricao_final
        if descricao_final.upper().startswith("C SUP-DIVSUBS"):
            desc_sem_prefixo = descricao_final[len("C SUP-DIVSUBS"):].lstrip()
        elif descricao_final.upper().startswith("C SUP-DIV"):
            desc_sem_prefixo = descricao_final[len("C SUP-DIV"):].lstrip()

        # Inserir descrição adicional no parêntese da descrição do serviço (não no prefixo da OM)
        if descricao_om and descricao_om.strip():
            descricao_extra = descricao_om.strip()
            descricao_extra = re.sub(r'(?i)nr\s*opus\s*:?', 'Nr OPUS:', descricao_extra).strip()

            parenteses_desc = list(re.finditer(r'\(([^()]*)\)', desc_sem_prefixo))
            if parenteses_desc:
                match_parenteses = parenteses_desc[-1]
                conteudo_parenteses = (match_parenteses.group(1) or '').strip()
                if descricao_extra.lower() not in conteudo_parenteses.lower():
                    conteudo_novo = f"{conteudo_parenteses} – {descricao_extra}" if conteudo_parenteses else descricao_extra
                    desc_sem_prefixo = (
                        desc_sem_prefixo[:match_parenteses.start()]
                        + f"({conteudo_novo})"
                        + desc_sem_prefixo[match_parenteses.end():]
                    )
            else:
                desc_sem_prefixo = f"{desc_sem_prefixo} ({descricao_extra})"

        descricao_final = prefixo + desc_sem_prefixo

        # Adicionar referência à PRO se existir
        referencia_pro = f" REF '{numero_pro}'." if numero_pro else ""
        
        # Adicionar referência ao DIEx e data
        referencia_diex = f" REF '{diex}' DE '{data}'." if diex and data else ""
        
        # Adicionar informações padrão
        descricao_completa = f"{descricao_final}{referencia_pro}{referencia_diex} EMPH ATÉ 48H. PRB DETAORC."
        
        # Remover aspas simples e duplas da descrição
        descricao_sem_aspas = descricao_completa.replace("'", "").replace('"', "")
        print(f"✅ Descrição gerada (sem aspas): {descricao_sem_aspas[:100]}...")
        return descricao_sem_aspas
        
    except Exception as e:
        print(f"❌ Erro ao construir descrição da NC: {str(e)}")
        import traceback
        traceback.print_exc()
        referencia_pro = f" REF '{numero_pro}'." if numero_pro else ""
        referencia_diex = f" REF '{diex}' DE '{data}'." if diex and data else ""
        return f"C SUP-DIVSUBS-SOL DSP EXTRAORDINÁRIA. OBSERVAR REGRAS.{referencia_pro}{referencia_diex} EMPH ATÉ 48H. PRB DETAORC."


def normalizar_prefixo_descricao(descricao_atual, om, codom):
    """Normaliza descrição existente aplicando o prefixo padrão (CODOM-OM) C SUP-DIV SUBS."""
    prefixo = "C SUP-DIV SUBS "
    if codom and om:
        prefixo = f"({codom}-{om}) C SUP-DIV SUBS "
    elif om:
        prefixo = f"({om}) C SUP-DIV SUBS "
    elif codom:
        prefixo = f"({codom}) C SUP-DIV SUBS "

    # Strip antigos prefixos para evitar duplicação
    texto = (descricao_atual or "").strip()
    texto = re.sub(r'^\([^)]*\)\s*C\s*SUP-DIV\s*SUBS\s*', '', texto, flags=re.IGNORECASE)
    texto = re.sub(r'^C\s*SUP-DIVSUBS\s*', '', texto, flags=re.IGNORECASE)
    texto = re.sub(r'^C\s*SUP-DIV\s*', '', texto, flags=re.IGNORECASE)

    return prefixo + texto.strip()


def normalizar_descricao_legada_nc(descricao_atual, om, codom, nr_opus=''):
    """Normaliza descrições antigas de NC para o padrão atual (prefixo, REF e posição do Nr OPUS)."""
    texto = normalizar_prefixo_descricao(descricao_atual or '', om, codom)

    # Padronizar referências antigas
    texto = re.sub(r"\bREF\s+PRO\s+'([^']+)'\.", r"REF '\1'.", texto, flags=re.IGNORECASE)
    texto = re.sub(r"\bREF\s+DIEX\s+'([^']+)'\s+DE\s+'([^']+)'\.", r"REF '\1' DE '\2'.", texto, flags=re.IGNORECASE)

    # Capturar Nr OPUS já existente no texto (ou usar o da solicitação)
    opus_encontrado = re.search(r'Nr\s*OPUS\s*:?\s*([^\)\.;]+)', texto, flags=re.IGNORECASE)
    valor_opus = (opus_encontrado.group(1).strip() if opus_encontrado else '')
    if not valor_opus and nr_opus:
        valor_opus = str(nr_opus).strip()

    # Remover ocorrências antigas de Nr OPUS para reposicionar corretamente
    texto = re.sub(r'\s*[–-]?\s*Nr\s*OPUS\s*:?\s*[^\)\.;]+', '', texto, flags=re.IGNORECASE)
    texto = re.sub(r'\(\s*\)', '', texto)

    # Reposicionar Nr OPUS no último parêntese (descrição do serviço)
    if valor_opus:
        opus_texto = f"Nr OPUS: {valor_opus}"
        parenteses = list(re.finditer(r'\(([^()]*)\)', texto))
        if parenteses:
            alvo = parenteses[-1]
            conteudo = (alvo.group(1) or '').strip()
            if opus_texto.lower() not in conteudo.lower():
                novo_conteudo = f"{conteudo} – {opus_texto}" if conteudo else opus_texto
                texto = texto[:alvo.start()] + f"({novo_conteudo})" + texto[alvo.end():]
        else:
            texto = f"{texto} ({opus_texto})"

    return texto.strip()

def valor_por_extenso(valor):
    """Converte valor numérico para extenso em português de forma correta"""
    if valor == 0:
        return "zero"
    
    # Dicionários para conversão
    unidades = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove']
    dez_a_dezenove = ['dez', 'onze', 'doze', 'treze', 'catorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove']
    dezenas = ['', '', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa']
    centenas = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos', 'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos']
    
    def converter_centenas(n):
        """Converte números de 0 a 999"""
        if n == 0:
            return ''
        elif n == 100:
            return 'cem'
        
        texto = ''
        c = n // 100
        resto = n % 100
        d = resto // 10
        u = resto % 10
        
        if c > 0:
            texto += centenas[c]
            if resto > 0:
                texto += ' e '
        
        if resto > 0:
            if d == 1:
                texto += dez_a_dezenove[u]
            else:
                if d > 1:
                    texto += dezenas[d]
                    if u > 0:
                        texto += ' e '
                if u > 0 and d != 1:
                    texto += unidades[u]
        
        return texto
    
    # Separar parte inteira e centavos
    parte_inteira = int(valor)
    centavos = round((valor - parte_inteira) * 100)
    
    # Converter parte inteira
    if parte_inteira == 0:
        texto_inteiro = "zero"
    elif parte_inteira == 1:
        texto_inteiro = "um"
    else:
        # Converter milhões, milhares e unidades
        milhões = parte_inteira // 1000000
        resto = parte_inteira % 1000000
        milhares = resto // 1000
        unidades_num = resto % 1000
        
        partes = []
        
        if milhões > 0:
            if milhões == 1:
                partes.append("um milhão")
            else:
                partes.append(converter_centenas(milhões) + " milhões")
        
        if milhares > 0:
            if milhares == 1:
                partes.append("mil")
            else:
                partes.append(converter_centenas(milhares) + " mil")
        
        if unidades_num > 0:
            partes.append(converter_centenas(unidades_num))
        
        texto_inteiro = ' '.join(partes)
    
    # Adicionar "real" ou "reais"
    if parte_inteira == 1:
        texto_completo = texto_inteiro + " real"
    else:
        texto_completo = texto_inteiro + " reais"
    
    # Adicionar centavos se houver
    if centavos > 0:
        if centavos == 1:
            texto_completo += " e um centavo"
        else:
            texto_centavos = converter_centenas(centavos)
            texto_completo += f" e {texto_centavos} centavos"
    
    return texto_completo

def construir_texto_pro(solicitacao):
    """Constrói o texto da PRO conforme o formato especificado, incluindo todas as OMs"""
    try:
        print(f"📋 Construindo PRO para solicitação com {len(solicitacao.pedidos)} OMs")
        
        # Agrupar itens por OM
        itens_por_om = {}
        for pedido in solicitacao.pedidos:
            om_key = pedido.om
            if om_key not in itens_por_om:
                itens_por_om[om_key] = {
                    'codom': pedido.codom,
                    'codug': pedido.codug,
                    'sigla_ug': pedido.sigla_ug,
                    'itens': []
                }
            
            for item in pedido.itens:
                if item.valor_aprovado > 0:
                    itens_por_om[om_key]['itens'].append(item)
        
        # Construir partes do texto
        partes = []
        total_geral = 0
        
        print(f"📊 Processando {len(itens_por_om)} OMs na PRO")
        
        for om_nome, dados_om in itens_por_om.items():
            print(f"  📝 Processando OM: {om_nome}")
            
            # Agrupar itens por ND para esta OM
            itens_por_nd = {}
            for item in dados_om['itens']:
                if item.nd not in itens_por_nd:
                    itens_por_nd[item.nd] = []
                itens_por_nd[item.nd].append(item)
            
            for nd, itens in itens_por_nd.items():
                total_nd = sum(item.valor_aprovado for item in itens)
                if total_nd > 0:
                    total_geral += total_nd
                    
                    # Formatar valor por extenso
                    valor_extenso = valor_por_extenso(total_nd)
                    
                    # Descrição baseada no primeiro item
                    descricao = itens[0].finalidade if itens and itens[0].finalidade else "aquisição"
                    
                    # Formatar valor com separador de milhares
                    valor_formatado = f"R$ {total_nd:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    
                    parte = f"no valor de {valor_formatado} ({valor_extenso}), na ND '{nd}', para '{descricao}' na {om_nome}"
                    partes.append(parte)
                    
                    print(f"    ✅ ND {nd}: {valor_formatado} - {valor_extenso}")
        
        # Juntar todas as partes
        if partes:
            texto_itens = '; '.join(partes)
            
            # Buscar COD UG (usar o primeiro pedido como referência)
            pedido_principal = solicitacao.pedidos[0] if solicitacao.pedidos else None
            cod_ug = pedido_principal.codug if pedido_principal else ''
            
            # Construir texto completo
            texto_completo = f"""1. Em atenção ao constante do documento da referência, informo ao senhor que há Previsão de Recurso Orçamentário (PRO), {texto_itens} (CODUG '{cod_ug}'). 
2. Solicito que, tão logo o processo licitatório esteja concluído, possibilitando o empenho imediato, sejam encaminhados os valores atualizados para provisão de recursos.
3. Para outros esclarecimentos, coloco à disposição a Divisão de Subsistência/ C Sup, por meio do telefone (61) 3415-5229, RITEx 860-5229 ou e-mail sglsubs.liab@gmail.com.
Por ordem do Comandante Logístico"""
            
            print(f"✅ PRO gerada com sucesso para {len(itens_por_om)} OMs")
            print(f"💰 Total geral: R$ {total_geral:,.2f}")
            
            return texto_completo
        else:
            print("⚠️  Nenhum item com valor aprovado > 0 encontrado")
            return "Texto da PRO não pôde ser gerado (nenhum valor aprovado)."
        
    except Exception as e:
        print(f"❌ Erro ao construir texto da PRO: {str(e)}")
        import traceback
        traceback.print_exc()
        return "Texto da PRO não pôde ser gerado automaticamente."

def gerar_pro_automatica(solicitacao):
    """Gera uma PRO automaticamente após aprovação da solicitação com modalidade PRO"""
    try:
        # Verificar se já existe PRO para esta solicitação
        pro_existente = Pro.query.filter_by(solicitacao_pdrlog_id=solicitacao.id).first()
        if pro_existente:
            print(f"⚠️  PRO já existe para esta solicitação: {pro_existente.numero}")
            return pro_existente
        
        # Gerar número da PRO
        ano = datetime.now().year
        ultima_pro = Pro.query.filter(
            Pro.numero.like(f'PRO-{ano}-%')
        ).order_by(Pro.id.desc()).first()
        
        if ultima_pro:
            ultimo_numero = int(ultima_pro.numero.split('-')[-1])
            novo_numero = f"PRO-{ano}-{ultimo_numero + 1:03d}"
        else:
            novo_numero = f"PRO-{ano}-001"
        
        print(f"✅ Número da PRO gerado: {novo_numero}")
        print(f"📊 Gerando PRO para solicitação com {len(solicitacao.pedidos)} OMs:")
        
        for i, pedido in enumerate(solicitacao.pedidos):
            total_om = sum(item.valor_aprovado for item in pedido.itens)
            print(f"   • OM {i+1}: {pedido.om} - Total aprovado: R$ {total_om:,.2f}")
        
        # Construir texto da PRO
        texto_pro = construir_texto_pro(solicitacao)
        
        # Buscar dados da primeira OM para os campos básicos
        pedido_principal = solicitacao.pedidos[0] if solicitacao.pedidos else None
        if not pedido_principal:
            print("❌ Nenhum pedido encontrado na solicitação")
            return None
        
        # Criar a PRO - NOVO STATUS
        valor_total = sum(
            item.valor_aprovado for pedido in solicitacao.pedidos for item in pedido.itens
        )
        pro = Pro(
            numero=novo_numero,
            status='Aguardando término do processo licitatório',  # NOVO STATUS
            usuario_id=current_user.id,
            solicitacao_pdrlog_id=solicitacao.id,
            cod_ug=pedido_principal.codug or '',
            sigla_ug=pedido_principal.sigla_ug or '',
            descricao=texto_pro,
            valor_total=valor_total,
            valor_restante=valor_total
        )
        
        db.session.add(pro)
        db.session.commit()
        
        print(f"✅ PRO gerada automaticamente: {novo_numero}")
        print(f"📝 PRO inclui {len(solicitacao.pedidos)} OMs da solicitação")
        return pro
        
    except Exception as e:
        print(f"❌ Erro ao gerar PRO automática: {str(e)}")
        import traceback
        traceback.print_exc()
        return None
    
# DECORATOR PARA CONTROLE DE ACESSO
def acesso_requerido(*niveis_permitidos):
    """Decorator para controlar acesso baseado no nível do usuário"""
    def decorator(f):
        @wraps(f)
        @login_required
        def decorated_function(*args, **kwargs):
            if current_user.nivel_acesso not in niveis_permitidos:
                if current_user.nivel_acesso == 'nc_only':
                    # Usuário NC Only é redirecionado para lista de NCs
                    flash('Acesso permitido apenas para Notas de Crédito.', 'warning')
                    return redirect(url_for('listar_ncs'))
                else:
                    flash('Acesso não autorizado para esta funcionalidade.', 'error')
                    return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return decorated_function
    return decorator

@app.route('/pro/<int:id>/pdf')
@acesso_requerido('admin', 'usuario')
def gerar_pdf_pro(id):
    """Gera PDF da PRO no formato especificado usando reportlab - VERSÃO FINAL AJUSTADA"""
    try:
        # Buscar a PRO com todos os relacionamentos
        pro = Pro.query\
            .options(
                db.joinedload(Pro.solicitacao_origem)
                    .joinedload(SolicitacaoExtraPDRLOG.pedidos)
                    .joinedload(PedidoSolicitacao.itens)
            )\
            .filter(Pro.id == id)\
            .first()
        
        if not pro:
            flash('PRO não encontrada!', 'error')
            return redirect(url_for('listar_pros'))

        if pro.status in STATUS_PRO_VARIANTES:
            pro.status = STATUS_PRO_VARIANTES[pro.status]
        
        if not pro.solicitacao_origem:
            flash('Solicitação relacionada não encontrada!', 'error')
            return redirect(url_for('detalhes_pro', id=id))
        
        solicitacao = pro.solicitacao_origem
        
        # Determinar finalidade principal - AGORA VERIFICA SE TEM DESCRIÇÃO ESPECÍFICA
        finalidade_principal = "Extra PDRLOG"
        
        # Verificar se há descrição específica na solicitação
        if solicitacao.descricao and "DESCRIÇÃO PARA NC:" in solicitacao.descricao:
            # Extrair a descrição específica
            descricao_parts = solicitacao.descricao.split("DESCRIÇÃO PARA NC:")
            if len(descricao_parts) > 1:
                finalidade_principal = descricao_parts[1].strip()
                print(f"✅ Usando descrição específica para NC: {finalidade_principal}")
        else:
            # Se não tem descrição específica, usar a finalidade mais comum dos itens
            finalidades = []
            for pedido in solicitacao.pedidos:
                for item in pedido.itens:
                    if item.valor_aprovado > 0 and item.finalidade:
                        finalidades.append(item.finalidade)
            
            if finalidades:
                finalidade_principal = max(set(finalidades), key=finalidades.count)
                print(f"✅ Usando finalidade mais comum: {finalidade_principal}")
        
        # O restante da função permanece igual...
        # Criar buffer para o PDF
        buffer = BytesIO()
        
        # Criar documento com margens ajustadas
        doc = SimpleDocTemplate(buffer, pagesize=A4,
                               rightMargin=1.5*cm, leftMargin=1.5*cm,
                               topMargin=1.5*cm, bottomMargin=2*cm)
        
        # Estilos (mantidos iguais)
        styles = getSampleStyleSheet()
        
        # Estilo personalizado para cabeçalho
        header_style = ParagraphStyle(
            'Header',
            parent=styles['Normal'],
            fontSize=12,
            alignment=1,
            spaceAfter=6,
            fontName='Helvetica-Bold'
        )
        
        title_style = ParagraphStyle(
            'Title',
            parent=styles['Normal'],
            fontSize=14,
            alignment=1,
            spaceAfter=8,
            textDecoration='underline',
            fontName='Helvetica-Bold'
        )
        
        bold_style = ParagraphStyle(
            'Bold',
            parent=styles['Normal'],
            fontSize=11,
            fontName='Helvetica-Bold'
        )
        
        normal_style = ParagraphStyle(
            'NormalJustified',
            parent=styles['Normal'],
            fontSize=11,
            alignment=4,
            spaceAfter=6,
            fontName='Helvetica'
        )
        
        # Estilo para tabela com quebra de texto e centralização
        table_style = ParagraphStyle(
            'TableStyle',
            parent=styles['Normal'],
            fontSize=9,
            fontName='Helvetica',
            wordWrap='CJK',
            alignment=1
        )
        
        # Estilo para valor (alinhado à direita)
        valor_style = ParagraphStyle(
            'ValorStyle',
            parent=styles['Normal'],
            fontSize=9,
            fontName='Courier',
            alignment=2
        )
        
        # Conteúdo do PDF
        story = []
        
        # ADICIONAR LOGOTIPO (código mantido igual)
        try:
            caminhos_logo = [
                os.path.join('static', 'img', 'brasao.png'),
                os.path.join('img', 'brasao.png'),
                'brasao.png'
            ]
            logo_encontrado = False
            for caminho in caminhos_logo:
                if os.path.exists(caminho):
                    from reportlab.platypus import Image
                    # reduzir ligeiramente o brasão para ganhar espaço vertical
                    logo = Image(caminho, width=2.2*cm, height=2.2*cm)
                    logo.hAlign = 'CENTER'
                    story.append(logo)
                    # espaçamento menor após o brasão para compensar altura reduzida
                    story.append(Spacer(1, 0.15*cm))
                    logo_encontrado = True
                    print(f"✅ Logotipo carregado de: {caminho}")
                    break
            if not logo_encontrado:
                print(f"⚠️ Logotipo não encontrado em nenhum dos caminhos: {caminhos_logo}")
        except Exception as e:
            print(f"⚠️ Erro ao carregar logotipo: {e}")
        
        # Cabeçalho (código mantido igual)
        story.append(Paragraph('MINISTÉRIO DA DEFESA', header_style))
        story.append(Paragraph('EXÉRCITO BRASILEIRO', header_style))
        story.append(Paragraph('COMANDO LOGÍSTICO', header_style))
        story.append(Paragraph('DEPARTAMENTO MARECHAL FALCONIERI', header_style))
        story.append(Spacer(1, 0.5*cm))
        
        # Título e número na mesma linha (mais espaço entre eles)
        # Formatar número exibido: remover 'PRO-' e inverter para '<seq>-<ano>' (ex: 121-2026)
        display_num = pro.numero
        try:
            parts = (pro.numero or '').split('-')
            if len(parts) >= 3 and parts[0].upper() == 'PRO':
                display_num = f"{parts[2]}-{parts[1]}"
        except Exception:
            display_num = pro.numero

        # reduzir distância visível entre título e número
        titulo_com_num = f'Previsão de Recurso Orçamentário&nbsp;&nbsp;&nbsp;Nr {display_num}'
        story.append(Paragraph(titulo_com_num, title_style))
        story.append(Spacer(1, 0.5*cm))
        
        # Unidade Gestora (não exibida no PDF; variável mantida para uso em nome de arquivo)
        nome_ug = pro.sigla_ug or "UG"
        
        # Item 1 - USAR FINALIDADE PRINCIPAL (que agora pode ser a descrição do usuário)
        story.append(Paragraph(
            f'<b>1.</b> O Comando Logístico conta com a previsão de recurso orçamentário da Ação 212B - Plano Orçamentário: 0006 Finalidade: <b>"{finalidade_principal}"</b>.',
            normal_style
        ))
        
        # O restante do código da função permanece igual...
        # Item 2
        story.append(Paragraph('<b>2.</b> Deverão ser alocados créditos conforme o quadro abaixo:', normal_style))
        story.append(Spacer(1, 0.3*cm))
        
        # Preparar dados da tabela com Plano Interno (PI)
        table_data = []
        
        # Cabeçalho da tabela - COM PI - AZUL MUITO CLARO
        header_row = [
            Paragraph('Unidade Gestora<br/>Executora', table_style),
            Paragraph('CODUG', table_style),
            Paragraph('Finalidade', table_style),
            Paragraph('ND', table_style),
            Paragraph('PI', table_style),
            Paragraph('Valor<br/>(R$)', table_style)
        ]
        table_data.append(header_row)
        
        total_geral = 0
        valor_total_pdf = float(pro.valor_total or 0)
        
        for pedido in solicitacao.pedidos:
            # Filtrar apenas itens com valor aprovado > 0
            itens_com_valor = [item for item in pedido.itens if item.valor_aprovado > 0]
            
            if itens_com_valor:
                # Agrupar por ND e PI
                itens_agrupados = {}
                for item in itens_com_valor:
                    chave = f"{item.nd}_{item.pi}"
                    if chave not in itens_agrupados:
                        itens_agrupados[chave] = {
                            'nd': item.nd,
                            'pi': item.pi or '--',
                            'valor': 0,
                            'finalidade': item.finalidade
                        }
                    itens_agrupados[chave]['valor'] += item.valor_aprovado
                
                # Determinar finalidade para esta OM
                finalidades_om = [item.finalidade for item in itens_com_valor if item.finalidade]
                finalidade_om = max(set(finalidades_om), key=finalidades_om.count) if finalidades_om else finalidade_principal
                
                # Adicionar linhas para cada item agrupado
                itens_lista = list(itens_agrupados.values())
                
                for idx, item_data in enumerate(itens_lista):
                    # Criar textos com quebra automática
                    om_text = pedido.om if idx == 0 else ""
                    codug_text = pedido.codug or "" if idx == 0 else ""

                    finalidade_text = ""
                    if idx == 0:
                        finalidade_partes = []

                        descricao_om_nc = (pedido.descricao_om or '').strip()
                        if descricao_om_nc:
                            finalidade_partes.append(descricao_om_nc)

                        nr_opus = (solicitacao.nr_opus or '').strip()
                        if nr_opus:
                            finalidade_partes.append(f"Nr OPUS: {nr_opus}")

                        finalidade_text = '<br/>'.join([parte for parte in finalidade_partes if parte])
                    
                    table_row = [
                        Paragraph(om_text, table_style),
                        Paragraph(codug_text, table_style),
                        Paragraph(finalidade_text, table_style),
                        Paragraph(item_data['nd'], table_style),
                        Paragraph(item_data['pi'], table_style),
                        Paragraph(f'R$ {item_data["valor"]:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'), valor_style)
                    ]
                    table_data.append(table_row)
        
        # Adicionar linha de total
        total_row = [
            Paragraph('', table_style),
            Paragraph('', table_style),
            Paragraph('', table_style),
            Paragraph('', table_style),
            Paragraph('<b>TOTAL:</b>', table_style),
            Paragraph(f'<b>R$ {valor_total_pdf:,.2f}</b>'.replace(',', 'X').replace('.', ',').replace('X', '.'), valor_style)
        ]
        table_data.append(total_row)
        
        # Criar tabela com larguras ajustadas
        col_widths = [3.5*cm, 1.5*cm, 5*cm, 2*cm, 2*cm, 3*cm]
        
        # Criar tabela
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        # Estilo da tabela - CABEÇALHO AZUL MUITO CLARO, TODOS ITENS CENTRALIZADOS
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#87CEEB')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('TOPPADDING', (0, 0), (-1, 0), 8),
            
            ('BACKGROUND', (0, 1), (-1, -2), colors.HexColor('#f8f9fa')),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#dee2e6')),
            
            ('ALIGN', (0, 1), (4, -2), 'CENTER'),
            ('ALIGN', (5, 1), (5, -2), 'RIGHT'),
            
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#e9ecef')),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('ALIGN', (0, -1), (4, -1), 'CENTER'),
            ('ALIGN', (5, -1), (5, -1), 'RIGHT'),
            ('LINEABOVE', (0, -1), (-1, -1), 1, colors.black),
            
            ('WORDWRAP', (0, 0), (-1, -1), True),
            
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        
        story.append(table)
        story.append(Spacer(1, 0.8*cm))
        
        # Itens 3-6
        story.append(Paragraph(
            '<b>3.</b> Em consequência, essa Unidade Gestora deverá dar início aos processos licitatórios de acordo com a legislação em vigência.',
            normal_style
        ))
        
        story.append(Paragraph(
            f'<b>4.</b> Ao final do processo licitatório, o resultado deverá ser informado por intermédio de DIEx, vinculando o DIEx que solicitou a PRO (<b>{solicitacao.diex}</b>) para que o respectivo crédito seja liberado.',
            normal_style
        ))
        
        story.append(Paragraph(
            '<b>5.</b> A UG terá o prazo de até 120 dias para informar o resultado do processo de licitação. Findo este prazo, sem manifestação da UG, a previsão de recurso orçamentário deverá ser anulada.',
            normal_style
        ))
        
        story.append(Paragraph(
            '<b>6.</b> Estão autorizados os procedimentos necessários à fase interna e externa da licitação, em A-1 (ano anterior à vigência da Lei Orçamentária Anual - LOA), baseados no Projeto de Lei Orçamentária Anual (PLOA), condicionando em A (ano de vigência da LOA) a homologação do certame e a correspondente contratação, com a efetiva disponibilidade orçamentária.',
            normal_style
        ))
        
        if solicitacao.diex_dom:
            data_formatada = solicitacao.data_solicitacao.strftime("%d %b %y").upper() if solicitacao.data_solicitacao else ""
            story.append(Paragraph(
                f'(DIEx nº {solicitacao.diex_dom}, {data_formatada})',
                ParagraphStyle('Italic', parent=styles['Normal'], fontSize=9, fontName='Helvetica-Oblique')
            ))
        # Inserir local e data por extenso (não será afetado pela redução automática de Spacers)
        meses_pt = {
            1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril', 5: 'maio', 6: 'junho',
            7: 'julho', 8: 'agosto', 9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'
        }
        hoje = datetime.now()
        data_extenso = f"{hoje.day} de {meses_pt[hoje.month]} de {hoje.year}"
        # aumentar o espaçamento entre data e assinatura sem usar Spacer (ParagraphStyle preservado)
        date_style = ParagraphStyle('DateLine', parent=styles['Normal'], fontSize=11, alignment=1, spaceBefore=18, spaceAfter=14)
        story.append(Paragraph(f'Brasília-DF, {data_extenso}.', date_style))

        # Assinatura (linha e identificação) - usar espaçoBefore maior para afastar da data
        signature_line_style = ParagraphStyle('SigLine', parent=styles['Normal'], alignment=1, spaceBefore=22)
        signature_name_style = ParagraphStyle('SigName', parent=styles['Normal'], alignment=1, spaceBefore=10)
        signature_role_style = ParagraphStyle('SigRole', parent=styles['Normal'], alignment=1, spaceAfter=22)

        # Assinatura: apenas nome e cargo (removida a linha traçada)
        # aumentar espaço antes do nome para dar área para assinatura manuscrita
        signature_name_style = ParagraphStyle('SigName', parent=styles['Normal'], alignment=1, spaceBefore=48, fontSize=11)
        signature_role_style = ParagraphStyle('SigRole', parent=styles['Normal'], alignment=1, spaceAfter=30, fontSize=11)
        story.append(Paragraph('Gen Bda ERNESTO ISAACODETE DUTRA PEREIRA BATISTA LOPES', signature_name_style))
        story.append(Paragraph('Chefe de Suprimento', signature_role_style))

        # Forçar o conteúdo a caber em uma página
        from reportlab.platypus import PageBreak
        # Remove todos os PageBreaks (não há, mas por garantia)
        story = [el for el in story if not isinstance(el, PageBreak)]
        # Reduzir espaçamentos para garantir que caiba em uma página
        for i, el in enumerate(story):
            if isinstance(el, Spacer):
                if el.height > 0.5*cm:
                    story[i] = Spacer(1, 0.3*cm)

        # Construir PDF em uma página
        doc.build(story, onFirstPage=lambda canvas, doc: None, onLaterPages=lambda canvas, doc: None)
        
        buffer.seek(0)
        
        # Nome do arquivo formatado: PRO Nr 'xxx'-'Nome UG'
        nome_ug_safe = ''.join(c for c in nome_ug if c.isalnum() or c in (' ', '-', '_')).strip()
        nome_ug_safe = nome_ug_safe.replace(' ', '_')
        
        nome_arquivo = f"PRO_Nr_{pro.numero.replace('-', '_')}_{nome_ug_safe}.pdf"
        
        # Criar resposta
        response = make_response(buffer.read())
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'attachment; filename="{nome_arquivo}"'
        
        return response
        
    except Exception as e:
        print(f"❌ Erro ao gerar PDF da PRO: {str(e)}")
        import traceback
        traceback.print_exc()
        flash(f'Erro ao gerar PDF: {str(e)}', 'error')
        return redirect(url_for('detalhes_pro', id=id))

def gerar_nc_automatica(solicitacao, descricao_usuario="", numero_pro=""):
    """Gera Notas de Crédito automaticamente - UMA NC POR ND, com referência à PRO se fornecida"""
    try:
        # CORREÇÃO: Carregar relacionamentos necessários
        solicitacao_completa = SolicitacaoExtraPDRLOG.query\
            .options(
                db.joinedload(SolicitacaoExtraPDRLOG.pedidos).joinedload(PedidoSolicitacao.itens)
            )\
            .filter(SolicitacaoExtraPDRLOG.id == solicitacao.id)\
            .first()
        
        if not solicitacao_completa:
            print("❌ Solicitação não encontrada")
            return None
        
        ncs_existentes = NotaCredito.query.filter_by(solicitacao_pdrlog_id=solicitacao_completa.id).all()
        if ncs_existentes:
            print(f"⚠️  NCs já existem para esta solicitação: {[nc.numero for nc in ncs_existentes]}")
            return ncs_existentes[0]
        
        ncs_geradas = []
        nc_counter = 0
        
        # GERAR UMA NC PARA CADA ND COM SUA FINALIDADE ESPECÍFICA
        for pedido in solicitacao_completa.pedidos:
            for item in pedido.itens:
                if item.valor_aprovado > 0:
                    try:
                        # Gerar número da NC
                        ano = datetime.now().year
                        ultima_nc = NotaCredito.query.filter(
                            NotaCredito.numero.like(f'NC-{ano}-%')
                        ).order_by(NotaCredito.id.desc()).first()
                        
                        if ultima_nc:
                            ultimo_numero = int(ultima_nc.numero.split('-')[-1])
                            novo_numero = f"NC-{ano}-{ultimo_numero + 1 + nc_counter:03d}"
                        else:
                            novo_numero = f"NC-{ano}-{nc_counter + 1:03d}"
                        
                        # Usar PI do item (já gerado automaticamente baseado na finalidade do ND)
                        pi = item.pi
                        
                        # Formatar data
                        data_diex_formatada = solicitacao_completa.data_solicitacao.strftime('%d/%m/%Y') if solicitacao_completa.data_solicitacao else ''
                        
                         # DEBUG: Verificar dados para descrição
                        print(f"🔍 Dados para descrição NC:")
                        print(f"   • PI: {pi}")
                        print(f"   • ND: {item.nd}")
                        print(f"   • DIEx: {solicitacao_completa.diex}")
                        print(f"   • Data: {data_diex_formatada}")
                        print(f"   • OM: {pedido.om}")
                        print(f"   • Descrição usuário: {descricao_usuario}")
                        print(f"   • Número PRO: {numero_pro}")

                        # Construir descrição automática INCLUINDO REFERÊNCIA À PRO
                        descricao_adicional = (pedido.descricao_om or '').strip()
                        nr_opus = (solicitacao_completa.nr_opus or '').strip()
                        if nr_opus:
                            opus_texto = f"Nr OPUS {nr_opus}"
                            if descricao_adicional:
                                if opus_texto.lower() not in descricao_adicional.lower():
                                    descricao_adicional = f"{descricao_adicional} - {opus_texto}"
                            else:
                                descricao_adicional = opus_texto

                        descricao_nc = construir_descricao_nc(
                            pi,
                            item.nd,
                            solicitacao_completa.diex,
                            data_diex_formatada,
                            pedido.om,
                            pedido.codom,
                            descricao_usuario,
                            numero_pro,
                            descricao_om=descricao_adicional
                        )
                        print(f"✅ Descrição gerada: {descricao_nc[:150]}...")
                        
                        # Criar a Nota de Crédito COM STATUS INICIAL "Não Conferida"
                        nc = NotaCredito(
                            numero=novo_numero,
                            status='Não Conferida',  # STATUS INICIAL CORRIGIDO
                            usuario_id=current_user.id,
                            solicitacao_pdrlog_id=solicitacao_completa.id,
                            cod_ug=pedido.codug or '',
                            sigla_ug=pedido.sigla_ug or '',
                            pi=pi,
                            nd=item.nd,
                            valor=item.valor_aprovado,
                            descricao=descricao_nc,  # DESCRIÇÃO GERADA AUTOMATICAMENTE
                            ref_sisnc='',
                            nc_siafi=''
                        )
                        
                        db.session.add(nc)
                        ncs_geradas.append(nc)
                        nc_counter += 1
                        
                        print(f"✅ Nota de Crédito gerada: {novo_numero}")
                        print(f"📋 Detalhes da NC:")
                        print(f"   • Status inicial: Não Conferida")
                        print(f"   • Descrição gerada: {descricao_nc[:100]}...")
                        print(f"   • OM: {pedido.om}")
                        print(f"   • ND: {item.nd}")
                        print(f"   • Finalidade: {item.finalidade}")
                        print(f"   • PI: {pi}")
                        print(f"   • Valor: R$ {item.valor_aprovado:,.2f}")
                        if numero_pro:
                            print(f"   • Referência PRO: {numero_pro}")
                        
                    except Exception as e:
                        print(f"❌ Erro ao gerar NC para ND {item.nd}: {str(e)}")
                        continue
        
        db.session.commit()
        
        if ncs_geradas:
            print(f"🎉 {len(ncs_geradas)} Nota(s) de Crédito gerada(s) automaticamente!")
            return ncs_geradas[0]
        else:
            print("❌ Nenhuma Nota de Crédito foi gerada (nenhum item com valor aprovado > 0)")
            return None
        
    except Exception as e:
        print(f"❌ Erro ao gerar Nota de Crédito automática: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

# FUNÇÃO PARA FILTRO NL2BR
def nl2br(value):
    """Converte quebras de linha em tags <br> para HTML"""
    if value:
        return value.replace('\n', '<br>')
    return value

# REGISTRAR O FILTRO NO JINJA2
app.jinja_env.filters['nl2br'] = nl2br
def extrair_descricao_nc(descricao_completa):
    """Extrai a descrição para NC da descrição completa da solicitação"""
    if not descricao_completa:
        return ""
    
    if "DESCRIÇÃO PARA NC:" in descricao_completa:
        partes = descricao_completa.split("DESCRIÇÃO PARA NC:")
        if len(partes) > 1:
            descricao_nc = partes[1].strip()
            # Remover linhas em branco extras
            descricao_nc = "\n".join([linha.strip() for linha in descricao_nc.split("\n") if linha.strip()])
            return descricao_nc
    
    return ""

# Registrar o filtro no Jinja2
app.jinja_env.filters['extrair_descricao_nc'] = extrair_descricao_nc
# ===== ROTAS DE AUTENTICAÇÃO =====

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Página de login"""
    if current_user.is_authenticated:
        if current_user.nivel_acesso == 'nc_only':
            return redirect(url_for('listar_ncs'))
        return redirect(url_for('painel_subsistencia'))
    
    if request.method == 'POST':
        email = request.form.get('email')
        senha = request.form.get('senha')
        
        usuario = Usuario.query.filter_by(email=email).first()
        
        if usuario and usuario.check_password(senha):
            login_user(usuario)
            flash(f'Bem-vindo, {usuario.nome}!', 'success')
            next_page = request.args.get('next')
            
            # Redirecionar usuário nc_only direto para NCs
            if usuario.nivel_acesso == 'nc_only':
                return redirect(url_for('listar_ncs'))
            
            return redirect(next_page or url_for('painel_subsistencia'))
        else:
            flash('Email ou senha incorretos!', 'error')
    
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    """Logout do usuário"""
    logout_user()
    flash('Logout realizado com sucesso!', 'success')
    return redirect(url_for('login'))

# ===== ROTAS PRINCIPAIS COM CONTROLE DE ACESSO =====

@app.route('/painel-subsistencia', methods=['GET', 'POST'])
@login_required
def painel_subsistencia():
    # Caminho para persistência dos dados CSV dentro do `instance/`
    instance_dir = app.instance_path
    os.makedirs(instance_dir, exist_ok=True)
    csv_path = os.path.join(instance_dir, 'Extra_PDRLOG_cache.csv')
    auditoria_upload_info = _carregar_meta_upload_auditoria('tg')
    sisnc_upload_info = _carregar_meta_upload_auditoria('sisnc')

    if request.method == 'POST':
        if current_user.nivel_acesso != 'admin':
            flash('Apenas administradores podem enviar planilhas.', 'error')
            return redirect(url_for('painel_subsistencia'))

        auditoria_file = request.files.get('auditoria_file')
        if auditoria_file and auditoria_file.filename:
            try:
                resumo_upload = _processar_upload_auditoria_unificada(auditoria_file)
                flash(
                    'Base unificada da auditoria carregada com sucesso. '
                    f"Total: {resumo_upload['total']} | "
                    f"PDR Log (PI com PLJ): {resumo_upload['pdr']} | "
                    f"Extra PDR Log (PI com SOL): {resumo_upload['extra']}.",
                    'success'
                )
            except Exception as e:
                flash(f'Erro ao processar upload unificado da auditoria: {str(e)}', 'error')
            return redirect(url_for('painel_subsistencia'))

        sisnc_file = request.files.get('sisnc_file')
        if sisnc_file and sisnc_file.filename:
            try:
                resumo_sisnc = _processar_upload_auditoria_sisnc(sisnc_file)
                flash(
                    'Base SISNC da auditoria carregada com sucesso. '
                    f"Total: {resumo_sisnc['total']} | "
                    f"PDR Log (PI com PLJ): {resumo_sisnc['pdr']} | "
                    f"Extra PDR Log (PI com SOL): {resumo_sisnc['extra']}.",
                    'success'
                )
            except Exception as e:
                flash(f'Erro ao processar upload SISNC da auditoria: {str(e)}', 'error')
            return redirect(url_for('painel_subsistencia'))

        bi_file = request.files.get('bi_file')
        if bi_file and bi_file.filename:
            try:
                df_bi, info_bi = _carregar_dataframe_bi_upload(bi_file)
                df_bi.to_csv(csv_path, index=False, encoding='utf-8-sig')
                flash(
                    f"Planilha BI carregada com sucesso (cabeçalho na linha {info_bi.get('header_row', 1)}).",
                    'success'
                )
            except Exception as e:
                flash(f'Erro ao processar planilha BI: {str(e)}', 'error')
            return redirect(url_for('painel_subsistencia'))

        flash('Nenhuma planilha válida foi selecionada para upload.', 'warning')
        return redirect(url_for('painel_subsistencia'))

    # Carregar DataFrame do cache se existir
    df = None
    if os.path.exists(csv_path):
        try:
            df = pd.read_csv(csv_path)
            # Detect possible multi-line header where second row contains
            # human-readable column names (e.g. 'CREDITO DISPONIVEL') which
            # end up as the first data row and cause object dtype concatenation
            # during groupby sums. If detected, re-read skipping that row.
            try:
                first_row_vals = df.iloc[0].astype(str).str.upper().tolist()
                header_like = any('CREDITO' in v or 'CREDITO DISPONIVEL' in v or 'DESPESAS EMPENHADAS' in v or 'DESPESAS LIQUIDADAS' in v for v in first_row_vals)
            except Exception:
                header_like = False
            if header_like:
                df = pd.read_csv(csv_path, skiprows=[1])
        except Exception as e:
            print(f"[painel_subsistencia] Erro ao ler CSV de cache: {e}")
    if df is None or df.empty:
        print("[painel_subsistencia] Nenhum dado disponível para exibir no painel.")
        return render_template(
            'painel_subsistencia.html',
            bi_totals=None,
            relacao_gnd_pi=None,
            relacao_pi_total=[],
            auditoria_upload_info=auditoria_upload_info,
            sisnc_upload_info=sisnc_upload_info
        )

    # Mapeamento flexível de nomes de colunas
    def find_col(possibilities):
        for col in df.columns:
            norm_col = ''.join(c for c in str(col).lower() if c.isalnum())
            for p in possibilities:
                norm_p = ''.join(c for c in p.lower() if c.isalnum())
                if norm_col == norm_p:
                    return col
        return None

    def coluna_valida(coluna):
        if coluna is None:
            return False
        try:
            if pd.isna(coluna):
                return False
        except Exception:
            pass
        return coluna in df.columns

    def detectar_coluna_por_conteudo(tipo):
        melhor_coluna = None
        melhor_score = 0
        for coluna in df.columns:
            try:
                serie = df[coluna].fillna('').astype(str).str.strip()
            except Exception:
                continue

            if tipo == 'pi':
                score = int(serie.str.contains(r'\bE\d[A-Z0-9]{6,}\b', case=False, regex=True, na=False).sum())
            elif tipo == 'ug':
                score_puro = int(serie.str.fullmatch(r'\d{6}(?:\.0+)?', na=False).sum())
                score_misto = int(serie.str.contains(r'(?<!\d)16\d{4}(?!\d)', regex=True, na=False).sum())
                score = score_puro + score_misto
            elif tipo == 'gnd':
                score_puro = int(serie.str.fullmatch(r'[3-9](?:\.0+)?', na=False).sum())
                score_texto = int(serie.str.contains(r'\bGND\s*[3-9]\b', case=False, regex=True, na=False).sum())
                score = score_puro + score_texto
            else:
                score = 0

            if score > melhor_score:
                melhor_score = score
                melhor_coluna = coluna

        # Evita falso positivo em colunas sem representatividade
        minimo = max(3, int(len(df) * 0.01))
        if melhor_score < minimo:
            return None
        return melhor_coluna

    relacao_gnd_pi = None
    if current_user.nivel_acesso == 'nc_only':
        return redirect(url_for('listar_ncs'))

    col_gnd = find_col(['Grupo Despesa', 'GRUPO DE DESPESA', 'GND'])
    col_pi = find_col(['PI', 'PROJETO/ATIVIDADE'])
    col_ug = find_col(['UG Executora', 'UG', 'UG_EXECUTORA', 'UGEXECUTORA'])
    col_credito = find_col(['19', 'CREDITO DISPONIVEL', 'CRÉDITO DISPONÍVEL'])
    col_empenhado = find_col(['23', 'DESPESAS EMPENHADAS', 'EMPENHADO'])
    col_liquidado = find_col(['25', 'DESPESAS LIQUIDADAS', 'LIQUIDADO'])
    col_pago = find_col(['28', 'DESPESAS PAGAS', 'PAGO'])

    # Fallback por conteúdo para casos de cabeçalho genérico (ex.: COLUNA, COLUNA_2...)
    if not coluna_valida(col_pi):
        col_pi = detectar_coluna_por_conteudo('pi')
    if not coluna_valida(col_ug):
        col_ug = detectar_coluna_por_conteudo('ug')
    if not coluna_valida(col_gnd):
        col_gnd = detectar_coluna_por_conteudo('gnd')

    # Normalizar e limpar colunas-chave para evitar concatenação de strings
    # e diferenças por espaços/maiúsculas que impedem consolidação por PI.
    if coluna_valida(col_pi):
        df[col_pi] = df[col_pi].astype(str).str.strip().replace({'': pd.NA})
    if coluna_valida(col_gnd):
        df[col_gnd] = df[col_gnd].astype(str).str.strip().replace({'': pd.NA})

    # Converter colunas numéricas para dtype numérico de forma robusta.
    numeric_cols = [c for c in (col_credito, col_empenhado, col_liquidado, col_pago) if coluna_valida(c)]
    for c in numeric_cols:
        if c in df.columns:
            # Remover caracteres inesperados, normalizar vírgula decimal para ponto
            df[c] = df[c].astype(str).str.replace(r'[^0-9\-,.]', '', regex=True).str.replace(',', '.', regex=False)
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)

    # Normalizar UG e criar colunas auxiliares para crédito UGR/UGE
    if coluna_valida(col_ug) and coluna_valida(col_credito):
        def norm_ug(v):
            try:
                if pd.isna(v):
                    return None
            except Exception:
                pass
            s = str(v).strip()
            candidatos = re.findall(r'(?<!\d)\d{6}(?!\d)', s)
            if candidatos:
                return candidatos[0]
            s_limpo = re.sub(r'\D', '', s)
            if len(s_limpo) >= 6:
                return s_limpo[:6]
            return s
        df['_ug_norm'] = df[col_ug].apply(norm_ug).ffill()
        # marcar créditos pertencentes à UG 160504
        df['_credito_ugr'] = df.apply(lambda r: float(r[col_credito]) if str(r['_ug_norm']) == '160504' else 0.0, axis=1)
        df['_credito_uge'] = df[col_credito] - df['_credito_ugr']

    try:
        bi_totals = {
            'credito_disponivel': float(pd.to_numeric(df[col_credito], errors='coerce').sum() or 0.0) if coluna_valida(col_credito) else 0.0,
            'empenhado': float(pd.to_numeric(df[col_empenhado], errors='coerce').sum() or 0.0) if coluna_valida(col_empenhado) else 0.0,
            'liquidado': float(pd.to_numeric(df[col_liquidado], errors='coerce').sum() or 0.0) if coluna_valida(col_liquidado) else 0.0,
            'pago': float(pd.to_numeric(df[col_pago], errors='coerce').sum() or 0.0) if coluna_valida(col_pago) else 0.0,
        }
        # total descentralizado: soma do crédito disponível (UGE restantes) + empenhado + liquidado + pago
        # campos adicionais poderão ser preenchidos abaixo (e.g., credito_uge_restantes)
    except Exception as e:
        print(f"[painel_subsistencia] Erro ao calcular totais: {e}")
        bi_totals = None

    # Se tivermos coluna de UG, separar crédito disponível entre UGR (UG 160504)
    # e UGE (restantes). Normalizamos a coluna antes da filtragem.
    try:
        if bi_totals is not None and coluna_valida(col_ug) and coluna_valida(col_credito):
            # Normalizar UG: tentar converter para inteiro quando possível
            def norm_ug(v):
                try:
                    if pd.isna(v):
                        return None
                except Exception:
                    pass
                s = str(v).strip()
                candidatos = re.findall(r'(?<!\d)\d{6}(?!\d)', s)
                if candidatos:
                    return candidatos[0]
                s_limpo = re.sub(r'\D', '', s)
                if len(s_limpo) >= 6:
                    return s_limpo[:6]
                return s

            ug_series = df[col_ug].apply(norm_ug)
            # Preencher valores ausentes copiando a UG do registro anterior (o CSV organiza blocos)
            ug_series = ug_series.ffill()
            # máscara para UG 160504
            mask_160504 = ug_series == '160504'
            credito_total = float(df[col_credito].sum())
            credito_ugr = float(df.loc[mask_160504, col_credito].sum()) if mask_160504.any() else 0.0
            credito_uge = credito_total - credito_ugr
            bi_totals['credito_ugr_160504'] = credito_ugr
            bi_totals['credito_uge_restantes'] = credito_uge
            # calcular Total Descentralizado: Crédito disponível (UGE restantes) + Empenhado
            try:
                bi_totals['total_descentralizado'] = (
                    float(bi_totals.get('credito_uge_restantes', 0.0) or 0.0)
                    + float(bi_totals.get('empenhado', 0.0) or 0.0)
                )
            except Exception:
                bi_totals['total_descentralizado'] = 0.0
    except Exception as e:
        print(f"[painel_subsistencia] Erro ao calcular quebra de UG: {e}")

    # Agrupamento por GND+PI
    relacao_gnd_pi = None
    if coluna_valida(col_gnd) and coluna_valida(col_pi):
        try:
            agg_gnd_pi = {}
            if coluna_valida(col_credito):
                agg_gnd_pi[col_credito] = 'sum'
            if coluna_valida(col_empenhado):
                agg_gnd_pi[col_empenhado] = 'sum'
            if coluna_valida(col_liquidado):
                agg_gnd_pi[col_liquidado] = 'sum'
            if coluna_valida(col_pago):
                agg_gnd_pi[col_pago] = 'sum'

            if agg_gnd_pi:
                agrupado = df.groupby([col_gnd, col_pi], dropna=False).agg(agg_gnd_pi).reset_index()
            else:
                agrupado = df[[col_gnd, col_pi]].drop_duplicates().reset_index(drop=True)

            relacao_gnd_pi = {}
            for _, row in agrupado.iterrows():
                gnd = row[col_gnd]
                pi = row[col_pi]
                def tofloat(val):
                    try:
                        return float(val)
                    except Exception:
                        return 0.0
                if pd.isna(gnd) or pd.isna(pi):
                    continue
                if gnd not in relacao_gnd_pi:
                    relacao_gnd_pi[gnd] = {}
                if pi not in relacao_gnd_pi[gnd]:
                    relacao_gnd_pi[gnd][pi] = {
                        'gnd': gnd,
                        'pi': pi,
                        'credito_disponivel': 0.0,
                        'empenhado': 0.0,
                        'liquidado': 0.0,
                        'pago': 0.0
                    }
                relacao_gnd_pi[gnd][pi]['credito_disponivel'] += tofloat(row[col_credito]) if coluna_valida(col_credito) and col_credito in agrupado.columns else 0.0
                relacao_gnd_pi[gnd][pi]['empenhado'] += tofloat(row[col_empenhado]) if coluna_valida(col_empenhado) and col_empenhado in agrupado.columns else 0.0
                relacao_gnd_pi[gnd][pi]['liquidado'] += tofloat(row[col_liquidado]) if coluna_valida(col_liquidado) and col_liquidado in agrupado.columns else 0.0
                relacao_gnd_pi[gnd][pi]['pago'] += tofloat(row[col_pago]) if coluna_valida(col_pago) and col_pago in agrupado.columns else 0.0
            # Converter para lista para compatibilidade com o template
            for gnd in relacao_gnd_pi:
                relacao_gnd_pi[gnd] = list(relacao_gnd_pi[gnd].values())
            print(f"[painel_subsistencia] relacao_gnd_pi gerada: {relacao_gnd_pi}")
        except KeyError as e:
            print(f"[painel_subsistencia] KeyError no agrupamento GND+PI: {e}")
            flash('A planilha BI foi carregada, mas o layout não permitiu consolidar GND/PI automaticamente.', 'warning')
            relacao_gnd_pi = None

    # Agrupamento por PI (total)
    relacao_pi_total = []
    if coluna_valida(col_pi):
        try:
            # Agrupar por PI e incluir crédito UGR/UGE se colunas auxiliares existirem
            agg_map = {}
            if coluna_valida(col_empenhado):
                agg_map[col_empenhado] = 'sum'
            if coluna_valida(col_liquidado):
                agg_map[col_liquidado] = 'sum'
            if coluna_valida(col_pago):
                agg_map[col_pago] = 'sum'
            if coluna_valida(col_credito):
                agg_map[col_credito] = 'sum'
            if '_credito_ugr' in df.columns:
                agg_map['_credito_ugr'] = 'sum'
            if '_credito_uge' in df.columns:
                agg_map['_credito_uge'] = 'sum'
            if agg_map:
                agrupado_pi = df.groupby([col_pi], dropna=False).agg(agg_map).reset_index()
            else:
                agrupado_pi = df[[col_pi]].drop_duplicates().reset_index(drop=True)
            for _, row in agrupado_pi.iterrows():
                pi = row[col_pi]
                def tofloat(val):
                    try:
                        return float(val)
                    except Exception:
                        return 0.0
                if pd.isna(pi):
                    continue
                relacao_pi_total.append({
                    'pi': pi,
                    'credito_disponivel_ugr': tofloat(row['_credito_ugr']) if '_credito_ugr' in agrupado_pi.columns else (tofloat(row[col_credito]) if coluna_valida(col_credito) and col_credito in agrupado_pi.columns else 0.0),
                    'credito_disponivel_uge': tofloat(row['_credito_uge']) if '_credito_uge' in agrupado_pi.columns else 0.0,
                    'empenhado': tofloat(row[col_empenhado]) if coluna_valida(col_empenhado) and col_empenhado in agrupado_pi.columns else 0.0,
                    'liquidado': tofloat(row[col_liquidado]) if coluna_valida(col_liquidado) and col_liquidado in agrupado_pi.columns else 0.0,
                    'pago': tofloat(row[col_pago]) if coluna_valida(col_pago) and col_pago in agrupado_pi.columns else 0.0
                })
                # adicionar Total Descentralizado por PI: Crédito disponível (UGE restantes) + Empenhado
                relacao_pi_total[-1]['total_descentralizado'] = (
                    relacao_pi_total[-1].get('credito_disponivel_uge', 0.0)
                    + relacao_pi_total[-1].get('empenhado', 0.0)
                )
        except KeyError as e:
            print(f"[painel_subsistencia] KeyError no agrupamento por PI: {e}")
            flash('A planilha BI foi carregada, mas o layout não permitiu consolidar totais por PI automaticamente.', 'warning')
            relacao_pi_total = []
    else:
        relacao_pi_total = []

    print(f"[painel_subsistencia] relacao_gnd_pi FINAL: {relacao_gnd_pi}")
    print(f"[painel_subsistencia] relacao_pi_total FINAL: {relacao_pi_total}")
    if relacao_pi_total is not None and len(relacao_pi_total) == 0:
        print("[painel_subsistencia] AVISO: relacao_pi_total está vazia após processamento!")

    return render_template(
        'painel_subsistencia.html',
        bi_totals=bi_totals,
        relacao_gnd_pi=relacao_gnd_pi,
        relacao_pi_total=relacao_pi_total,
        auditoria_upload_info=auditoria_upload_info,
        sisnc_upload_info=sisnc_upload_info
    )


@app.route('/painel-subsistencia/tabela-oms', methods=['GET', 'POST'])
@app.route('/painel-subsistencia/tabela-oms/', methods=['GET', 'POST'])
@app.route('/painel_subsistencia/tabela_oms', methods=['GET', 'POST'])
@app.route('/painel_subsistencia/tabela_oms/', methods=['GET', 'POST'])
@app.route('/tabela_oms_subsistencia', methods=['GET', 'POST'])
@app.route('/tabela_oms_subsistencia/', methods=['GET', 'POST'])
@acesso_requerido('admin', 'usuario')
def tabela_oms_subsistencia():
    if request.method == 'POST':
        if current_user.nivel_acesso != 'admin':
            flash('Apenas administradores podem alterar a tabela de OMs.', 'error')
            return redirect(url_for('tabela_oms_subsistencia'))

        acao = (request.form.get('acao') or '').strip().lower()

        try:
            if acao == 'importar_padrao':
                lista = _extrair_oms_da_extra_pdrlog_xlsx()
                if not lista:
                    flash('Nenhum registro de OM foi encontrado na planilha Extra PDRLOG.xlsx.', 'warning')
                else:
                    _salvar_tabela_oms(lista)
                    _recarregar_dados_planilha_globais()
                    flash(f'Tabela de OMs importada com sucesso da Extra PDRLOG.xlsx ({len(lista)} registros).', 'success')

            elif acao == 'importar_upload':
                arquivo = request.files.get('arquivo_oms')
                if not arquivo or not arquivo.filename:
                    flash('Selecione um arquivo para importar a tabela de OMs.', 'warning')
                else:
                    lista = _extrair_oms_da_extra_pdrlog_xlsx(arquivo_upload=arquivo)
                    if not lista:
                        flash('Nenhum registro de OM válido foi encontrado no arquivo enviado.', 'warning')
                    else:
                        _salvar_tabela_oms(lista)
                        _recarregar_dados_planilha_globais()
                        flash(f'Tabela de OMs importada com sucesso do arquivo enviado ({len(lista)} registros).', 'success')

            elif acao == 'inserir':
                rm = request.form.get('rm', '')
                om = request.form.get('om', '')
                codom = request.form.get('codom', '')
                codug = request.form.get('codug', '')
                sigla_ug = request.form.get('sigla_ug', '')

                nova = _normalizar_linha_om(rm=rm, om=om, codom=codom, codug=codug, sigla_ug=sigla_ug)
                if not nova['OM']:
                    flash('O campo OM é obrigatório para inserção.', 'warning')
                else:
                    tabela = _carregar_tabela_oms()
                    mapa = {item['OM'].upper(): item for item in tabela}
                    mapa[nova['OM'].upper()] = nova
                    _salvar_tabela_oms(sorted(mapa.values(), key=lambda x: x.get('OM', '').upper()))
                    _recarregar_dados_planilha_globais()
                    flash(f'OM "{nova["OM"]}" salva com sucesso na tabela.', 'success')

            elif acao == 'atualizar':
                om_original = request.form.get('om_original', '')
                rm = request.form.get('rm', '')
                om = request.form.get('om', '')
                codom = request.form.get('codom', '')
                codug = request.form.get('codug', '')
                sigla_ug = request.form.get('sigla_ug', '')

                if not om_original:
                    flash('Registro de OM inválido para atualização.', 'warning')
                else:
                    tabela = _carregar_tabela_oms()
                    chave_antiga = om_original.strip().upper()
                    atualizada = _normalizar_linha_om(rm=rm, om=om, codom=codom, codug=codug, sigla_ug=sigla_ug)
                    if not atualizada['OM']:
                        flash('O campo OM é obrigatório para atualização.', 'warning')
                    else:
                        nova_tabela = []
                        for item in tabela:
                            if item.get('OM', '').strip().upper() == chave_antiga:
                                continue
                            nova_tabela.append(item)
                        nova_tabela.append(atualizada)
                        _salvar_tabela_oms(sorted(nova_tabela, key=lambda x: x.get('OM', '').upper()))
                        _recarregar_dados_planilha_globais()
                        flash(f'OM "{atualizada["OM"]}" atualizada com sucesso.', 'success')

            elif acao == 'excluir':
                om_original = request.form.get('om_original', '')
                if not om_original:
                    flash('Registro de OM inválido para exclusão.', 'warning')
                else:
                    chave = om_original.strip().upper()
                    tabela = _carregar_tabela_oms()
                    nova_tabela = [item for item in tabela if item.get('OM', '').strip().upper() != chave]
                    _salvar_tabela_oms(nova_tabela)
                    _recarregar_dados_planilha_globais()
                    flash(f'OM "{om_original}" excluída da tabela.', 'success')

            else:
                flash('Ação inválida para a tabela de OMs.', 'warning')

        except Exception as e:
            flash(f'Erro ao processar tabela de OMs: {str(e)}', 'error')

        return redirect(url_for('tabela_oms_subsistencia'))

    registros = _carregar_tabela_oms()
    # Filtros GET
    filtro_om = request.args.get('filtro_om', '').strip().lower()
    filtro_codom = request.args.get('filtro_codom', '').strip().lower()
    filtro_codug = request.args.get('filtro_codug', '').strip().lower()
    filtro_sigla_ug = request.args.get('filtro_sigla_ug', '').strip().lower()

    def match_filtro(valor, filtro):
        return filtro in (valor or '').strip().lower() if filtro else True

    registros_filtrados = [r for r in registros if
        match_filtro(r.get('OM', ''), filtro_om) and
        match_filtro(r.get('CODOM', ''), filtro_codom) and
        match_filtro(r.get('CODUG', ''), filtro_codug) and
        match_filtro(r.get('SIGLA_UG', ''), filtro_sigla_ug)
    ]

    return render_template(
        'tabela_oms_subsistencia.html',
        registros=registros_filtrados,
        total_registros=len(registros_filtrados)
    )


@app.route('/exportar_painel_subsistencia_excel')
@acesso_requerido('admin', 'usuario')
def exportar_painel_subsistencia_excel():
    """Exporta os dados do painel subsistência para um arquivo Excel.
    Gera três abas: Resumo LOA, Totais por PI e Relação GND-PI.
    """
    instance_dir = app.instance_path
    csv_path = os.path.join(instance_dir, 'Extra_PDRLOG_cache.csv')
    if not os.path.exists(csv_path):
        flash('Arquivo de cache BI não encontrado para exportação.', 'error')
        return redirect(url_for('painel_subsistencia'))

    df = pd.read_csv(csv_path)

    # Reuse the same find_col logic as painel_subsistencia
    def find_col_local(df, possibilities):
        for col in df.columns:
            norm_col = ''.join(c for c in str(col).lower() if c.isalnum())
            for p in possibilities:
                norm_p = ''.join(c for c in p.lower() if c.isalnum())
                if norm_col == norm_p:
                    return col
        return None

    col_pi = find_col_local(df, ['PI', 'PROJETO/ATIVIDADE'])
    col_ug = find_col_local(df, ['UG Executora', 'UG', 'UG_EXECUTORA', 'UGEXECUTORA'])
    col_credito = find_col_local(df, ['19', 'CREDITO DISPONIVEL', 'CRÉDITO DISPONÍVEL'])
    col_empenhado = find_col_local(df, ['23', 'DESPESAS EMPENHADAS', 'EMPENHADO'])
    col_liquidado = find_col_local(df, ['25', 'DESPESAS LIQUIDADAS', 'LIQUIDADO'])
    col_pago = find_col_local(df, ['28', 'DESPESAS PAGAS', 'PAGO'])

    # Normalize and coerce numeric
    if col_pi in df.columns:
        df[col_pi] = df[col_pi].astype(str).str.strip().replace({'': pd.NA})
    numeric_cols = [c for c in (col_credito, col_empenhado, col_liquidado, col_pago) if c]
    for c in numeric_cols:
        if c in df.columns:
            # remover quaisquer caracteres que não sejam dígitos, vírgula, ponto ou sinal negativo
            df[c] = df[c].astype(str).str.replace(r'[^0-9,\.\-]', '', regex=True).str.replace(',', '.', regex=False)
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)

    # UG normalized and per-row split
    if col_ug in df.columns and col_credito in df.columns:
        def norm_ug(v):
            try:
                if pd.isna(v):
                    return None
            except Exception:
                pass
            s = str(v).strip()
            if s.endswith('.0'):
                s = s[:-2]
            return s
        df['_ug_norm'] = df[col_ug].apply(norm_ug).ffill()
        df['_credito_ugr'] = df.apply(lambda r: float(r[col_credito]) if str(r['_ug_norm']) == '160504' else 0.0, axis=1)
        df['_credito_uge'] = df[col_credito] - df['_credito_ugr']

    # bi_totals
    bi_totals = {
        'credito_disponivel': float(df[col_credito].sum()) if col_credito in df.columns else 0.0,
        'empenhado': float(df[col_empenhado].sum()) if col_empenhado in df.columns else 0.0,
        'liquidado': float(df[col_liquidado].sum()) if col_liquidado in df.columns else 0.0,
        'pago': float(df[col_pago].sum()) if col_pago in df.columns else 0.0,
    }
    if '_credito_ugr' in df.columns:
        bi_totals['credito_ugr_160504'] = float(df['_credito_ugr'].sum())
        bi_totals['credito_uge_restantes'] = float(df['_credito_uge'].sum())
        bi_totals['total_descentralizado'] = bi_totals['credito_uge_restantes'] + bi_totals['empenhado']

    # Totais por PI
    rel_pi_df = pd.DataFrame()
    if col_pi and col_credito:
        agg_map = {}
        agg_map[col_credito] = 'sum'
        if col_empenhado: agg_map[col_empenhado] = 'sum'
        if col_liquidado: agg_map[col_liquidado] = 'sum'
        if col_pago: agg_map[col_pago] = 'sum'
        if '_credito_ugr' in df.columns: agg_map['_credito_ugr'] = 'sum'
        if '_credito_uge' in df.columns: agg_map['_credito_uge'] = 'sum'
        rel_pi_df = df.groupby([col_pi], dropna=False).agg(agg_map).reset_index()
        # rename columns for clarity
        rename_map = {}
        if '_credito_ugr' in rel_pi_df.columns:
            rename_map['_credito_ugr'] = 'Credito_UGR'
        if '_credito_uge' in rel_pi_df.columns:
            rename_map['_credito_uge'] = 'Credito_UGE'
        if col_credito in rel_pi_df.columns:
            rename_map[col_credito] = 'Credito_Total'
        if col_empenhado in rel_pi_df.columns:
            rename_map[col_empenhado] = 'Empenhado'
        if col_liquidado in rel_pi_df.columns:
            rename_map[col_liquidado] = 'Liquidado'
        if col_pago in rel_pi_df.columns:
            rename_map[col_pago] = 'Pago'
        rel_pi_df = rel_pi_df.rename(columns=rename_map)
        if 'Credito_UGE' in rel_pi_df.columns and 'Empenhado' in rel_pi_df.columns:
            rel_pi_df['Total_Descentralizado'] = rel_pi_df['Credito_UGE'] + rel_pi_df['Empenhado']

    # Relação GND-PI sheet (if available)
    rel_gnd_pi_df = pd.DataFrame()
    col_gnd = find_col_local(df, ['Grupo Despesa', 'GRUPO DE DESPESA', 'GND'])
    if col_gnd and col_pi:
        agg_map2 = {}
        if col_credito: agg_map2[col_credito] = 'sum'
        if col_empenhado: agg_map2[col_empenhado] = 'sum'
        if col_liquidado: agg_map2[col_liquidado] = 'sum'
        if col_pago: agg_map2[col_pago] = 'sum'
        if '_credito_ugr' in df.columns: agg_map2['_credito_ugr'] = 'sum'
        if '_credito_uge' in df.columns: agg_map2['_credito_uge'] = 'sum'
        rel_gnd_pi_df = df.groupby([col_gnd, col_pi], dropna=False).agg(agg_map2).reset_index()

    # Build Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Resumo LOA
        resumo_df = pd.DataFrame([bi_totals])
        resumo_df.to_excel(writer, sheet_name='Resumo LOA', index=False)
        # Totais por PI
        if not rel_pi_df.empty:
            rel_pi_df.to_excel(writer, sheet_name='Totais por PI', index=False)
        # GND-PI
        if not rel_gnd_pi_df.empty:
            rel_gnd_pi_df.to_excel(writer, sheet_name='GND-PI', index=False)
        # Ajustar larguras (simples)
        for sheet in writer.sheets:
            ws = writer.sheets[sheet]
            for col_cells in ws.columns:
                try:
                    max_length = max((len(str(cell.value)) for cell in col_cells)) + 2
                except Exception:
                    max_length = 10
                col_letter = col_cells[0].column_letter
                ws.column_dimensions[col_letter].width = max_length
    output.seek(0)
    return send_file(
        output,
        download_name=f'painel_subsistencia_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )



@app.route('/pdr-log/solicitacoes')
@acesso_requerido('admin', 'usuario')
def listar_solicitacoes_pdr_log():
    """Lista de solicitações do PDR Log baseada na aba BD da planilha oficial."""
    definir_modulo_menu(MENU_MODULO_PDR)
    status_filter = request.args.get('status', '').strip()
    classe_filter = request.args.get('classe', '').strip()
    cod_assunto_filter = request.args.get('cod_assunto', '').strip()
    om_filter = request.args.get('om', '').strip()
    numero_filter = request.args.get('numero', '').strip()

    try:
        solicitacoes = _carregar_solicitacoes_pdr_log_planilha()
    except Exception as e:
        flash(f"Erro ao carregar planilha do PDR Log: {e}", 'error')
        return render_template(
            'solicitacoes_pdr_log.html',
            solicitacoes=[],
            total_solicitacoes=0,
            total_filtradas=0,
            status_filter=status_filter,
            classe_filter=classe_filter,
            cod_assunto_filter=cod_assunto_filter,
            om_filter=om_filter,
            numero_filter=numero_filter,
            status_opcoes=[],
            classes_opcoes=[],
            assuntos_opcoes=[]
        )

    status_opcoes = sorted({s['status'] for s in solicitacoes if s['status']})
    classes_opcoes = sorted({s['classe'] for s in solicitacoes if s['classe']})
    assuntos_opcoes = sorted({s['cod_assunto'] for s in solicitacoes if s['cod_assunto']})

    filtradas = solicitacoes
    if status_filter:
        filtradas = [s for s in filtradas if s['status'] == status_filter]
    if classe_filter:
        filtradas = [s for s in filtradas if s['classe'] == classe_filter]
    if cod_assunto_filter:
        filtradas = [s for s in filtradas if s['cod_assunto'] == cod_assunto_filter]
    if om_filter:
        termo_om = om_filter.lower()
        filtradas = [s for s in filtradas if termo_om in s['om_solicitante'].lower()]
    if numero_filter:
        termo_numero = numero_filter.lower()
        filtradas = [s for s in filtradas if termo_numero in s['numero'].lower()]

    return render_template(
        'solicitacoes_pdr_log.html',
        solicitacoes=filtradas,
        total_solicitacoes=len(solicitacoes),
        total_filtradas=len(filtradas),
        status_filter=status_filter,
        classe_filter=classe_filter,
        cod_assunto_filter=cod_assunto_filter,
        om_filter=om_filter,
        numero_filter=numero_filter,
        status_opcoes=status_opcoes,
        classes_opcoes=classes_opcoes,
        assuntos_opcoes=assuntos_opcoes
    )


@app.route('/pdr-log/auditoria-solicitacoes')
@acesso_requerido('admin', 'usuario')
def auditoria_solicitacoes_pdr_log():
    """Auditoria de duplicidades de solicitações por OM+PI+ND+VALOR."""
    definir_modulo_menu(MENU_MODULO_PDR)

    om_filter = request.args.get('om', '').strip().lower()
    pi_filter = normalizar_pi(request.args.get('pi', '').strip())
    nd_filter = normalizar_nd(request.args.get('nd', '').strip())

    try:
        solicitacoes = _carregar_solicitacoes_pdr_log_planilha()
    except Exception as e:
        flash(f"Erro ao carregar planilha do PDR Log: {e}", 'error')
        return render_template(
            'auditoria_solicitacoes_pdr_log.html',
            duplicidades=[],
            total_solicitacoes=0,
            total_grupos_duplicados=0,
            total_registros_duplicados=0,
            om_filter=om_filter,
            pi_filter=pi_filter,
            nd_filter=nd_filter
        )

    grupos = {}
    total_sem_numero = 0
    total_com_numero = 0
    for s in solicitacoes:
        numero_raw = str(s.get('numero', '') or '').strip()
        numero_norm = _normalizar_numero_solicitacao_pdr(numero_raw)
        if not numero_norm:
            total_sem_numero += 1
        else:
            total_com_numero += 1

        om_display = str(s.get('om_solicitante', '') or '').strip()
        om_chave = normalizar_om_solicitante_chave(om_display)
        pi = normalizar_pi(_mapear_pi_pdr_por_assunto(s.get('nome_assunto', '')))
        nd = normalizar_nd(s.get('cod_nd', ''))
        valor = round(normalizar_valor(s.get('valor_total', 0)), 2)
        valor_chave = f"{valor:.2f}"

        if not om_chave or not pi or not nd:
            continue

        chave = (om_chave, pi, nd, valor_chave)
        grupo = grupos.setdefault(chave, {
            'om_display': om_display or '--',
            'pi': pi,
            'nd': nd,
            'valor': valor,
            'solicitacoes': []
        })
        grupo['solicitacoes'].append({
            'numero': numero_raw or '--',
            'numero_ordem': int(numero_norm) if numero_norm.isdigit() else 0,
            'data_solicitacao': s.get('data_solicitacao', '--') or '--',
            'status': s.get('status', '--') or '--',
            'codom': s.get('codom', '--') or '--',
            'codug': s.get('codug', '--') or '--',
            'om_solicitante': om_display or '--'
        })

    duplicidades = []
    for _, grupo in grupos.items():
        if len(grupo['solicitacoes']) <= 1:
            continue
        duplicidades.append({
            'om_solicitante': grupo['om_display'],
            'pi': grupo['pi'],
            'nd': grupo['nd'],
            'valor': grupo['valor'],
            'quantidade': len(grupo['solicitacoes']),
            'solicitacoes': sorted(grupo['solicitacoes'], key=lambda x: x.get('numero_ordem', 0))
        })

    if om_filter:
        duplicidades = [d for d in duplicidades if om_filter in str(d.get('om_solicitante', '')).lower()]
    if pi_filter:
        duplicidades = [d for d in duplicidades if d.get('pi', '') == pi_filter]
    if nd_filter:
        duplicidades = [d for d in duplicidades if d.get('nd', '') == nd_filter]

    duplicidades.sort(key=lambda d: (d.get('quantidade', 0), d.get('valor', 0)), reverse=True)

    total_registros_duplicados = sum(item['quantidade'] for item in duplicidades)
    total_valor_registros_duplicados = round(
        sum((item.get('valor', 0.0) or 0.0) * (item.get('quantidade', 0) or 0) for item in duplicidades),
        2
    )
    total_valor_excedente_duplicidades = round(
        sum((item.get('valor', 0.0) or 0.0) * max((item.get('quantidade', 0) or 0) - 1, 0) for item in duplicidades),
        2
    )

    return render_template(
        'auditoria_solicitacoes_pdr_log.html',
        duplicidades=duplicidades,
        total_solicitacoes=len(solicitacoes),
        total_solicitacoes_com_numero=total_com_numero,
        total_solicitacoes_sem_numero=total_sem_numero,
        total_grupos_duplicados=len(duplicidades),
        total_registros_duplicados=total_registros_duplicados,
        total_valor_registros_duplicados=total_valor_registros_duplicados,
        total_valor_excedente_duplicidades=total_valor_excedente_duplicidades,
        om_filter=request.args.get('om', '').strip(),
        pi_filter=request.args.get('pi', '').strip(),
        nd_filter=request.args.get('nd', '').strip()
    )


@app.route('/pdr-log/dashboard')
@acesso_requerido('admin', 'usuario')
def dashboard_pdr_log():
    """Dashboard do módulo PDR Log com base exclusiva na aba BD da planilha oficial."""
    definir_modulo_menu(MENU_MODULO_PDR)

    status_filter = request.args.get('status', '').strip()
    pi_filter = normalizar_pi(request.args.get('pi', '').strip())
    nd_filter = normalizar_nd(request.args.get('nd', '').strip())
    om_filter = request.args.get('om', '').strip()

    try:
        solicitacoes = _carregar_solicitacoes_pdr_log_planilha()
    except Exception as e:
        flash(f"Erro ao carregar dados do dashboard PDR Log: {e}", 'error')
        solicitacoes = []

    status_opcoes = sorted({s['status'] for s in solicitacoes if s['status']})
    pi_opcoes = sorted({normalizar_pi(_mapear_pi_pdr_por_assunto(s.get('nome_assunto', ''))) for s in solicitacoes if normalizar_pi(_mapear_pi_pdr_por_assunto(s.get('nome_assunto', '')))} )
    nd_opcoes = sorted({normalizar_nd(s.get('cod_nd', '')) for s in solicitacoes if normalizar_nd(s.get('cod_nd', ''))})

    filtradas = solicitacoes
    if status_filter:
        filtradas = [s for s in filtradas if s['status'] == status_filter]
    if pi_filter:
        filtradas = [
            s for s in filtradas
            if normalizar_pi(_mapear_pi_pdr_por_assunto(s.get('nome_assunto', ''))) == pi_filter
        ]
    if nd_filter:
        filtradas = [s for s in filtradas if normalizar_nd(s.get('cod_nd', '')) == nd_filter]
    if om_filter:
        termo_om = om_filter.lower()
        filtradas = [s for s in filtradas if termo_om in s['om_solicitante'].lower()]

    total_solicitacoes = len(filtradas)
    total_valor = sum(float(s.get('valor_total', 0) or 0) for s in filtradas)

    status_contagem = {}
    for registro in filtradas:
        chave = registro.get('status') or 'Não informado'
        status_contagem[chave] = status_contagem.get(chave, 0) + 1

    pi_soma = {}
    for registro in filtradas:
        chave = normalizar_pi(_mapear_pi_pdr_por_assunto(registro.get('nome_assunto', ''))) or 'Não informado'
        pi_soma[chave] = pi_soma.get(chave, 0) + float(registro.get('valor_total', 0) or 0)

    om_soma = {}
    for registro in filtradas:
        chave = registro.get('om_solicitante') or 'Não informada'
        om_soma[chave] = om_soma.get(chave, 0) + float(registro.get('valor_total', 0) or 0)

    pis_ordenados = sorted(pi_soma.items(), key=lambda item: item[1], reverse=True)[:10]
    oms_ordenadas = sorted(om_soma.items(), key=lambda item: item[1], reverse=True)[:10]

    def _parse_data_br(valor):
        try:
            return datetime.strptime(valor, '%d/%m/%Y')
        except Exception:
            return datetime.min

    solicitacoes_recentes = sorted(
        filtradas,
        key=lambda item: _parse_data_br(item.get('data_solicitacao', '--')),
        reverse=True
    )[:10]

    solicitacoes_recentes_exibir = []
    for item in solicitacoes_recentes:
        item_view = dict(item)
        item_view['pi'] = normalizar_pi(_mapear_pi_pdr_por_assunto(item.get('nome_assunto', '')))
        item_view['nd'] = normalizar_nd(item.get('cod_nd', ''))
        solicitacoes_recentes_exibir.append(item_view)

    return render_template(
        'dashboard_pdr_log.html',
        total_solicitacoes=total_solicitacoes,
        total_valor=total_valor,
        status_filter=status_filter,
        pi_filter=request.args.get('pi', '').strip(),
        nd_filter=request.args.get('nd', '').strip(),
        om_filter=om_filter,
        status_opcoes=status_opcoes,
        pi_opcoes=pi_opcoes,
        nd_opcoes=nd_opcoes,
        status_labels=list(status_contagem.keys()),
        status_valores=list(status_contagem.values()),
        pi_labels=[item[0] for item in pis_ordenados],
        pi_valores=[item[1] for item in pis_ordenados],
        om_labels=[item[0] for item in oms_ordenadas],
        om_valores=[item[1] for item in oms_ordenadas],
        solicitacoes_recentes=solicitacoes_recentes_exibir
    )

@app.route('/')
@login_required
def dashboard_root():
    return redirect(url_for('painel_subsistencia'))
    nc_query = nc_query.join(SolicitacaoExtraPDRLOG, NotaCredito.solicitacao_pdrlog_id == SolicitacaoExtraPDRLOG.id)
    if finalidade_filter:
                nc_query = (
                    nc_query
                    .join(PedidoSolicitacao, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
                    .join(ItemPedido, ItemPedido.pedido_id == PedidoSolicitacao.id)
                    .filter(ItemPedido.finalidade == finalidade_filter)
                    .distinct()
                )

        # Subquery para reuso
    solicitacao_subq = solicitacao_query.with_entities(SolicitacaoExtraPDRLOG.id).subquery()

    def soma_solicitacoes(status_val=None):
            q = db.session.query(func.coalesce(func.sum(ItemPedido.valor_solicitado), 0))\
                .join(PedidoSolicitacao, ItemPedido.pedido_id == PedidoSolicitacao.id)\
                .join(SolicitacaoExtraPDRLOG, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)\
                .filter(SolicitacaoExtraPDRLOG.id.in_(select(solicitacao_subq)))
            if status_val:
                q = q.filter(SolicitacaoExtraPDRLOG.status == status_val)
            return q.scalar() or 0

    def conta_solicitacoes(status_val=None):
            q = solicitacao_query
            if status_val:
                q = q.filter(SolicitacaoExtraPDRLOG.status == status_val)
            return q.count()

    total_solicitacoes = conta_solicitacoes()
    total_solicitacoes_valor = soma_solicitacoes()

    solicitacoes_aguardando = conta_solicitacoes('Aguardando Análise')
    solicitacoes_aguardando_valor = soma_solicitacoes('Aguardando Análise')

    solicitacoes_analise = conta_solicitacoes('Em Análise')
    solicitacoes_analise_valor = soma_solicitacoes('Em Análise')

    solicitacoes_aprovadas = conta_solicitacoes('Aprovado Ch Sup')
    solicitacoes_aprovadas_valor = soma_solicitacoes('Aprovado Ch Sup')

    solicitacoes_negadas = conta_solicitacoes('Negado Ch Sup')
    solicitacoes_negadas_valor = soma_solicitacoes('Negado Ch Sup')

    solicitacoes_arquivadas = conta_solicitacoes('Arquivado')
    solicitacoes_arquivadas_valor = soma_solicitacoes('Arquivado')

        # NCs com os mesmos filtros (via solicitação origem)
    nc_query = NotaCredito.query
    if finalidade_filter or orgao_filter or descricao_filter:
            nc_query = nc_query.join(SolicitacaoExtraPDRLOG, NotaCredito.solicitacao_pdrlog_id == SolicitacaoExtraPDRLOG.id)
            if finalidade_filter:
                nc_query = (
                    nc_query
                    .join(PedidoSolicitacao, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
                    .join(ItemPedido, ItemPedido.pedido_id == PedidoSolicitacao.id)
                    .filter(ItemPedido.finalidade == finalidade_filter)
                    .distinct()
                )
            if orgao_filter:
                nc_query = nc_query.filter(SolicitacaoExtraPDRLOG.orgao_demandante == orgao_filter)
            if descricao_filter:
                nc_query = nc_query.filter(SolicitacaoExtraPDRLOG.descricao.ilike(f'%{descricao_filter}%'))

    total_ncs = nc_query.count()
    total_ncs_valor = (nc_query.with_entities(func.coalesce(func.sum(NotaCredito.valor), 0)).scalar() or 0)

    ncs_pendentes = nc_query.filter(NotaCredito.status == 'Pendente').count()
    ncs_pendentes_valor = (nc_query.filter(NotaCredito.status == 'Pendente')
                                      .with_entities(func.coalesce(func.sum(NotaCredito.valor), 0)).scalar() or 0)

        # PROs com filtros na solicitação de origem
    pro_query = Pro.query
    if finalidade_filter or orgao_filter or descricao_filter:
            pro_query = pro_query.join(SolicitacaoExtraPDRLOG, Pro.solicitacao_pdrlog_id == SolicitacaoExtraPDRLOG.id)
            if finalidade_filter:
                pro_query = (
                    pro_query
                    .join(PedidoSolicitacao, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
                    .join(ItemPedido, ItemPedido.pedido_id == PedidoSolicitacao.id)
                    .filter(ItemPedido.finalidade == finalidade_filter)
                    .distinct()
                )
            if orgao_filter:
                pro_query = pro_query.filter(SolicitacaoExtraPDRLOG.orgao_demandante == orgao_filter)
            if descricao_filter:
                pro_query = pro_query.filter(SolicitacaoExtraPDRLOG.descricao.ilike(f'%{descricao_filter}%'))

    total_pros = pro_query.count()
        # Valor de PRO baseado no total solicitado da solicitação de origem
    total_pros_valor = (db.session.query(func.coalesce(func.sum(ItemPedido.valor_solicitado), 0))
            .join(PedidoSolicitacao, ItemPedido.pedido_id == PedidoSolicitacao.id)
            .join(SolicitacaoExtraPDRLOG, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
            .join(Pro, Pro.solicitacao_pdrlog_id == SolicitacaoExtraPDRLOG.id)
            .filter(Pro.id.in_(pro_query.with_entities(Pro.id)))
            .scalar() or 0)

    pros_pendentes = pro_query.filter(Pro.status == 'Pendente').count()
    pros_pendentes_valor = (db.session.query(func.coalesce(func.sum(ItemPedido.valor_solicitado), 0))
            .join(PedidoSolicitacao, ItemPedido.pedido_id == PedidoSolicitacao.id)
            .join(SolicitacaoExtraPDRLOG, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
            .join(Pro, Pro.solicitacao_pdrlog_id == SolicitacaoExtraPDRLOG.id)
            .filter(Pro.id.in_(pro_query.filter(Pro.status == 'Pendente').with_entities(Pro.id)))
            .scalar() or 0)

        # Séries para gráficos (aplicam os mesmos filtros)
        # Somatório por finalidade (valor solicitado)
    soma_por_finalidade = {f: 0 for f in FINALIDADES}
    finais_rows = (
            db.session.query(
                ItemPedido.finalidade,
                func.coalesce(func.sum(ItemPedido.valor_solicitado), 0)
            )
            .join(PedidoSolicitacao, ItemPedido.pedido_id == PedidoSolicitacao.id)
            .join(SolicitacaoExtraPDRLOG, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
            .filter(SolicitacaoExtraPDRLOG.id.in_(select(solicitacao_subq.c.id)))
            .group_by(ItemPedido.finalidade)
            .all()
        )

    for fin, total_val in finais_rows:
            if fin in soma_por_finalidade:
                soma_por_finalidade[fin] = total_val or 0

    solicitacoes_por_finalidade = [soma_por_finalidade[f] for f in FINALIDADES]

        # Somatório por órgão demandante (valor solicitado)
    orgaos_labels = []
    orgaos_valores = []
    orgaos_rows = (
            db.session.query(
                SolicitacaoExtraPDRLOG.orgao_demandante,
                               func.coalesce(func.sum(ItemPedido.valor_solicitado), 0)
            )
            .join(PedidoSolicitacao, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)
            .join(ItemPedido, ItemPedido.pedido_id == PedidoSolicitacao.id)
            .filter(SolicitacaoExtraPDRLOG.id.in_(select(solicitacao_subq.c.id)))
            .group_by(SolicitacaoExtraPDRLOG.orgao_demandante)
            .order_by(func.coalesce(func.sum(ItemPedido.valor_solicitado), 0).desc())
            .all()
        )

    for orgao, total_val in orgaos_rows:
            orgaos_labels.append(orgao or '---')
            orgaos_valores.append(total_val or 0)

    solicitacoes_por_status = {}
    for status in STATUS_SOLICITACAO:
            count = solicitacao_query.filter(SolicitacaoExtraPDRLOG.status == status).count()
            solicitacoes_por_status[status] = count

        # Solicitações por mês (últimos 6 meses) por valor solicitado
    meses_labels = []
    meses_valores = []
    now = datetime.now()
    for back in range(5, -1, -1):
            idx = (now.year * 12 + now.month - 1) - back
            year = idx // 12
            month = idx % 12 + 1
            start = datetime(year, month, 1)
            if month == 12:
                end = datetime(year + 1, 1, 1)
            else:
                end = datetime(year, month + 1, 1)

            total_mes = db.session.query(func.coalesce(func.sum(ItemPedido.valor_solicitado), 0))\
                .join(PedidoSolicitacao, ItemPedido.pedido_id == PedidoSolicitacao.id)\
                .join(SolicitacaoExtraPDRLOG, PedidoSolicitacao.solicitacao_id == SolicitacaoExtraPDRLOG.id)\
                .filter(SolicitacaoExtraPDRLOG.id.in_(select(solicitacao_subq.c.id)))\
                .filter(SolicitacaoExtraPDRLOG.data_solicitacao >= start)\
                .filter(SolicitacaoExtraPDRLOG.data_solicitacao < end)\
                .scalar() or 0

            meses_labels.append(start.strftime('%b'))
            meses_valores.append(total_mes)


# ===== ROTAS DE SOLICITAÇÕES (APENAS ADMIN E USUARIO) =====

@app.route('/solicitacoes')
@acesso_requerido('admin', 'usuario')
def redirecionar_solicitacoes_para_pdr_log():
    """Mantém compatibilidade de URL antiga, redirecionando para a listagem oficial do PDR Log."""
    definir_modulo_menu(MENU_MODULO_PDR)
    return redirect(url_for('listar_solicitacoes_pdr_log', **request.args.to_dict()))


@app.route('/menu/pdr-log')
@login_required
def entrar_menu_pdr_log():
    if current_user.nivel_acesso == 'nc_only':
        return redirect(url_for('listar_ncs'))
    definir_modulo_menu(MENU_MODULO_PDR)
    return redirect(url_for('dashboard_pdr_log'))


@app.route('/menu/extra-pdr-log')
@login_required
def entrar_menu_extra_pdr_log():
    if current_user.nivel_acesso == 'nc_only':
        return redirect(url_for('listar_ncs'))
    definir_modulo_menu(MENU_MODULO_EXTRA)
    return redirect(url_for('dashboard'))


@app.route('/solicitacoes-extra')
@acesso_requerido('admin', 'usuario')
def listar_solicitacoes():
    """Lista todas as solicitações"""
    definir_modulo_menu(MENU_MODULO_EXTRA)
    try:
        status_filter = request.args.get('status', '')
        finalidade_filter = request.args.get('finalidade', '')
        modalidade_filter = request.args.get('modalidade', '')
        diex_filter = request.args.get('diex', '').strip()
        om_filter = request.args.get('om', '').strip()

        print(
            f"🔍 Filtros aplicados - Status: '{status_filter}', Finalidade: '{finalidade_filter}', Modalidade: '{modalidade_filter}', "
            f"DIEx: '{diex_filter}', OM: '{om_filter}'"
        )
        
        query = SolicitacaoExtraPDRLOG.query

        if status_filter:
            query = query.filter_by(status=status_filter)
            print(f"✅ Filtro de status aplicado: {status_filter}")

        pedidos_joined = False
        if om_filter or finalidade_filter:
            query = query.join(PedidoSolicitacao)
            pedidos_joined = True
            print("🔗 Join com PedidoSolicitacao aplicado para filtros de OM/Finalidade")

        if finalidade_filter:
            query = query.join(ItemPedido).filter(ItemPedido.finalidade == finalidade_filter)
            print(f"✅ Filtro de finalidade aplicado: {finalidade_filter}")

        if om_filter:
            if not pedidos_joined:
                query = query.join(PedidoSolicitacao)
                print("🔗 Join com PedidoSolicitacao aplicado para filtro de OM")
            query = query.filter(PedidoSolicitacao.om.ilike(f"%{om_filter}%"))
            print(f"✅ Filtro de OM aplicado: {om_filter}")

        if diex_filter:
            query = query.filter(SolicitacaoExtraPDRLOG.diex.ilike(f"%{diex_filter}%"))
            print(f"✅ Filtro de DIEx aplicado: {diex_filter}")


        if modalidade_filter:
            query = query.filter(SolicitacaoExtraPDRLOG.modalidade == modalidade_filter)
            print(f"✅ Filtro de modalidade aplicado: {modalidade_filter}")

        if finalidade_filter or om_filter:
            query = query.distinct()

        solicitacoes = query.order_by(SolicitacaoExtraPDRLOG.data_criacao.desc()).all()
        
        print(f"📊 Total de solicitações encontradas: {len(solicitacoes)}")
        
        return render_template('solicitacoes.html',
                    solicitacoes=solicitacoes,
                    status_filter=status_filter,
                    finalidade_filter=finalidade_filter,
                    modalidade_filter=modalidade_filter,
                    diex_filter=diex_filter,
                    om_filter=om_filter,
                    finalidades=FINALIDADES,
                    modalidades=MODALIDADES,
                    status_list=STATUS_SOLICITACAO)
    except Exception as e:
        print(f"❌ Erro ao carregar solicitações: {str(e)}")
        import traceback
        traceback.print_exc()
        flash(f'Erro ao carregar solicitações: {str(e)}', 'error')
        return render_template('solicitacoes.html',
                    solicitacoes=[],
                    status_filter='',
                    finalidade_filter='',
                    modalidade_filter='',
                    diex_filter='',
                    om_filter='',
                    finalidades=FINALIDADES,
                    modalidades=MODALIDADES,
                    status_list=STATUS_SOLICITACAO)

@app.route('/solicitacao/nova', methods=['GET', 'POST'])
@acesso_requerido('admin', 'usuario')
def nova_solicitacao():
    """Nova solicitação"""
    if request.method == 'POST':
        try:
            converter_valor = globals().get('moeda_para_float') or globals().get('normalizar_valor') or (lambda x: float(re.sub(r'[^0-9,\.\-]', '', str(x)).replace(',', '.') ) if x not in (None, '') else 0.0)
            # Debug: verificar dados recebidos
            print("=== DADOS RECEBIDOS NO FORMULÁRIO ===")
            for key in request.form:
                print(f"{key}: {request.form.getlist(key)}")
            
            # ========== CORREÇÃO: GERAR NÚMERO DA SOLICITAÇÃO PRIMEIRO ==========
            # Gerar número da solicitação
            ano = datetime.now().year
            ultima_solicitacao = SolicitacaoExtraPDRLOG.query.filter(
                SolicitacaoExtraPDRLOG.numero.like(f'PDRLOG-{ano}-%')
            ).order_by(SolicitacaoExtraPDRLOG.id.desc()).first()
            
            if ultima_solicitacao:
                ultimo_numero = int(ultima_solicitacao.numero.split('-')[-1])
                novo_numero = f"PDRLOG-{ano}-{ultimo_numero + 1:03d}"
            else:
                novo_numero = f"PDRLOG-{ano}-001"
            
            print(f"✅ Número da solicitação gerado: {novo_numero}")
            # ========== FIM DA CORREÇÃO ==========

            # Converter data
            data_solicitacao_str = request.form.get('data_solicitacao')
            data_solicitacao = datetime.strptime(data_solicitacao_str, '%Y-%m-%d') if data_solicitacao_str else datetime.now()
            
            # Criar solicitação principal (sem finalidade geral)
            solicitacao = SolicitacaoExtraPDRLOG(
                numero=novo_numero,  # AGORA A VARIÁVEL ESTÁ DEFINIDA
                diex=request.form.get('diex'),
                orgao_demandante=request.form.get('orgao_demandante'),
                data_solicitacao=data_solicitacao,
                descricao=request.form.get('descricao'),
                parecer_analise=request.form.get('parecer_analise', ''),
                despacho=request.form.get('despacho', ''),
                status='Aguardando Análise',
                finalidade='',  # Finalidade geral removida
                modalidade=request.form.get('modalidade'),
                usuario_id=current_user.id,
                # NOVOS CAMPOS PARA PASA-DEC e Mnt OP-DEC
                tem_fsv=request.form.get('tem_fsv', ''),
                diex_dom=request.form.get('diex_dom', ''),
                nr_opus=request.form.get('nr_opus', ''),
                destinatario_colog=request.form.get('destinatario_colog', '')
            )
            
            db.session.add(solicitacao)
            db.session.flush()
            print(f"✅ Solicitação principal criada: {solicitacao.numero}")
            
            # Processar múltiplos pedidos (OMs)
            oms = request.form.getlist('om[]')
            codoms = request.form.getlist('codom[]')
            codugs = request.form.getlist('codug[]')
            siglas_ug = request.form.getlist('sigla_ug[]')
            descricoes_om = request.form.getlist('descricao_om[]')
            descricoes_om = request.form.getlist('descricao_om[]')
            
            total_pedidos = 0
            total_itens = 0
            
            for i, om in enumerate(oms):
                if om and om.strip():
                    correspondencia_om = _resolver_correspondencia_om(
                        om,
                        codoms[i] if i < len(codoms) else '',
                        codugs[i] if i < len(codugs) else '',
                        siglas_ug[i] if i < len(siglas_ug) else ''
                    )
                    pedido = PedidoSolicitacao(
                        solicitacao_id=solicitacao.id,
                        om=om.strip(),
                        codom=correspondencia_om['codom'],
                        codug=correspondencia_om['codug'],
                        sigla_ug=correspondencia_om['sigla_ug'],
                        descricao_om=descricoes_om[i] if i < len(descricoes_om) else ''
                    )
                    db.session.add(pedido)
                    db.session.flush()
                    total_pedidos += 1
                    
                    print(f"✅ Pedido {i+1} criado: {pedido.om} | CODOM={pedido.codom} CODUG={pedido.codug} FONTE={correspondencia_om['fonte']}")
                    
                    # Processar NDs com finalidades específicas
                    nds_key = f'nd[{i}][]'
                    finalidades_key = f'finalidade_nd[{i}][]'
                    valores_solicitado_key = f'valor_solicitado[{i}][]'
                    valores_aprovado_key = f'valor_aprovado[{i}][]'

                    nds = request.form.getlist(nds_key)
                    finalidades_nd = request.form.getlist(finalidades_key)
                    valores_solicitado = request.form.getlist(valores_solicitado_key)
                    valores_aprovado = request.form.getlist(valores_aprovado_key)

                    print(f"🔄 Processando OM {i}: {om}")
                    print(f"📦 NDs encontrados: {nds}")
                    print(f"🎯 Finalidades por ND: {finalidades_nd}")

                    # Processar os itens (NDs) para este pedido
                    for j, nd in enumerate(nds):
                        if nd and nd.strip() and j < len(valores_solicitado) and j < len(valores_aprovado):
                            valor_solicitado_str = valores_solicitado[j]
                            valor_aprovado_str = valores_aprovado[j]
                            try:
                                valor_sol_float = converter_valor(valor_solicitado_str)
                                valor_apr_float = converter_valor(valor_aprovado_str)

                                # Obter finalidade específica para este ND
                                finalidade_nd = finalidades_nd[j] if j < len(finalidades_nd) and finalidades_nd[j] else ''

                                # Determinar PI baseado na finalidade do ND
                                pi = PI_POR_FINALIDADE.get(finalidade_nd, '')

                                item = ItemPedido(
                                    pedido_id=pedido.id,
                                    nd=nds[j].strip(),
                                    finalidade=finalidade_nd,
                                    pi=pi,
                                    valor_solicitado=valor_sol_float,
                                    valor_aprovado=valor_apr_float
                                )
                                db.session.add(item)
                                total_itens += 1
                                print(f"✅ Item {j+1} adicionado: ND {nds[j]}, Finalidade: {finalidade_nd}, PI: {pi}, Valor Solicitado R$ {valor_sol_float:.2f}, Valor Aprovado R$ {valor_apr_float:.2f}")
                            except ValueError as e:
                                print(f"❌ Erro ao converter valores: {e}")
                                continue
            
            db.session.commit()
            print(f"🎉 Solicitação finalizada: {total_pedidos} pedidos, {total_itens} itens")
            
            flash(f'Solicitação {novo_numero} cadastrada com sucesso! {total_pedidos} OM(s) e {total_itens} ND(s)', 'success')
            return redirect(url_for('detalhes_solicitacao', id=solicitacao.id))
            
        except Exception as e:
            db.session.rollback()
            print(f"❌ Erro detalhado no cadastro: {str(e)}")
            import traceback
            traceback.print_exc()
            flash(f'Erro ao cadastrar solicitação: {str(e)}', 'error')
    
    return render_template('cadastro_solicitacao.html', 
                         finalidades=FINALIDADES,
                         nd_possiveis=ND_POSSIVEIS,
                         orgaos_demandantes=ORGAOS_DEMANDANTES,
                         oms_data=_carregar_tabela_oms(),
                         modalidades=MODALIDADES,
                         despachos_opcoes=DESPACHOS_OPCOES,
                         pi_por_finalidade=PI_POR_FINALIDADE)

@app.route('/solicitacao/<int:id>/editar', methods=['POST'])
@acesso_requerido('admin', 'usuario')
def editar_solicitacao(id):
    """Editar solicitação"""
    solicitacao = db.session.get(SolicitacaoExtraPDRLOG, id)
    if not solicitacao:
        flash('Solicitação não encontrada!', 'error')
        return redirect(url_for('listar_solicitacoes'))
    
    if request.method == 'POST':
        try:
            converter_valor = globals().get('moeda_para_float') or globals().get('normalizar_valor') or (lambda x: float(re.sub(r'[^0-9,\.\-]', '', str(x)).replace(',', '.') ) if x not in (None, '') else 0.0)
            print("=== INICIANDO EDIÇÃO DA SOLICITAÇÃO ===")
            
            # Salvar status anterior para verificar mudanças
            status_anterior = solicitacao.status
            
            # Atualizar dados básicos da solicitação
            solicitacao.diex = request.form.get('diex')
            solicitacao.orgao_demandante = request.form.get('orgao_demandante')
            solicitacao.descricao = request.form.get('descricao')
            solicitacao.parecer_analise = request.form.get('parecer_analise', '')
            solicitacao.despacho = request.form.get('despacho', '')
            solicitacao.modalidade = request.form.get('modalidade')
            
            # NOVO: Capturar descrição para NC do usuário
            descricao_nc_usuario = request.form.get('descricao_nc_usuario', '').strip()
            if descricao_nc_usuario:
                print(f"✅ Descrição para NC capturada: {descricao_nc_usuario}")
                # Salvar na descrição da solicitação com marcador
                if solicitacao.descricao:
                    # Verificar se já tem descrição NC
                    if "DESCRIÇÃO PARA NC:" in solicitacao.descricao:
                        # Substituir descrição NC existente
                        partes = solicitacao.descricao.split("DESCRIÇÃO PARA NC:")
                        nova_descricao = partes[0].strip() + f"\n\nDESCRIÇÃO PARA NC: {descricao_nc_usuario}"
                        solicitacao.descricao = nova_descricao
                    else:
                        # Adicionar nova descrição NC
                        solicitacao.descricao += f"\n\nDESCRIÇÃO PARA NC: {descricao_nc_usuario}"
                else:
                    # Criar nova descrição com NC
                    solicitacao.descricao = f"DESCRIÇÃO PARA NC: {descricao_nc_usuario}"
            
            # NOVOS CAMPOS PARA PASA-DEC e Mnt OP-DEC
            solicitacao.tem_fsv = request.form.get('tem_fsv', '')
            solicitacao.diex_dom = request.form.get('diex_dom', '')
            solicitacao.nr_opus = request.form.get('nr_opus', '')
            solicitacao.destinatario_colog = request.form.get('destinatario_colog', '')
            
            # ATUALIZAÇÃO AUTOMÁTICA DO STATUS - CORREÇÃO
            parecer_preenchido = solicitacao.parecer_analise and solicitacao.parecer_analise.strip()
            despacho_preenchido = solicitacao.despacho and solicitacao.despacho.strip()
            
            # Verificar se foi devolvido para correções
            if parecer_preenchido and 'devolvido para a om para correções' in solicitacao.parecer_analise.lower():
                solicitacao.status = 'Devolvido para correções'
            elif not parecer_preenchido:
                # Se não tem parecer, volta para aguardando análise
                solicitacao.status = 'Aguardando Análise'
            elif parecer_preenchido and not despacho_preenchido:
                # Se tem parecer mas não tem despacho, vai para aguardando despacho
                solicitacao.status = 'Aguardando despacho'
            elif parecer_preenchido and despacho_preenchido:
                # Se tem ambos, analisa o conteúdo do despacho
                despacho_lower = solicitacao.despacho.lower()
                if 'aprovado integralmente' in despacho_lower:
                    solicitacao.status = 'Aprovado Ch Sup'
                elif 'aprovado parcialmente' in despacho_lower:
                    solicitacao.status = 'Aprovado Parcialmente'
                elif 'negado' in despacho_lower:
                    solicitacao.status = 'Negado Ch Sup'
                else:
                    # Se não identificou, mantém como aguardando despacho
                    solicitacao.status = 'Aguardando despacho'
            
            print(f"🔄 Status atualizado: {status_anterior} -> {solicitacao.status}")
            
            # VERIFICAR SE DEVE GERAR NC OU PRO AUTOMATICAMENTE
            deve_gerar_documento = (
                status_anterior != 'Aprovado Ch Sup' and
                solicitacao.status in ['Aprovado Ch Sup', 'Aprovado Parcialmente'] and
                despacho_preenchido
            )
            
            if deve_gerar_documento:
                print("🎯 Condição para gerar documento atendida! Iniciando geração...")
            
            # DEBUG: Verificar dados recebidos
            print("📋 DADOS RECEBIDOS NO FORMULÁRIO:")
            for key in request.form:
                if key.startswith('om') or key.startswith('nd') or key.startswith('finalidade') or key.startswith('valor') or key == 'descricao_nc_usuario':
                    print(f"  {key}: {request.form.get(key)}")
            
            # PRIMEIRO: Limpar todos os pedidos e itens existentes
            for pedido in solicitacao.pedidos:
                for item in pedido.itens:
                    db.session.delete(item)
                db.session.delete(pedido)
            
            db.session.flush()
            
            # SEGUNDO: Recriar os pedidos e itens - CORREÇÃO PARA MÚLTIPLAS OMs
            oms = request.form.getlist('om[]')
            codoms = request.form.getlist('codom[]')
            codugs = request.form.getlist('codug[]')
            siglas_ug = request.form.getlist('sigla_ug[]')
            descricoes_om = request.form.getlist('descricao_om[]')
            
            print(f"📊 OMs a processar: {oms}")
            print(f"📊 Total de OMs: {len(oms)}")
            
            # CORREÇÃO: Processar cada OM individualmente com seus próprios NDs
            for i, om in enumerate(oms):
                if om and om.strip():
                    correspondencia_om = _resolver_correspondencia_om(
                        om,
                        codoms[i] if i < len(codoms) else '',
                        codugs[i] if i < len(codugs) else '',
                        siglas_ug[i] if i < len(siglas_ug) else ''
                    )
                    pedido = PedidoSolicitacao(
                        solicitacao_id=id,
                        om=om.strip(),
                        codom=correspondencia_om['codom'],
                        codug=correspondencia_om['codug'],
                        sigla_ug=correspondencia_om['sigla_ug'],
                        descricao_om=descricoes_om[i] if i < len(descricoes_om) else ''
                    )
                    db.session.add(pedido)
                    db.session.flush()  # Para obter o ID do pedido
                    
                    print(f"✅ Pedido criado: {pedido.om} (índice {i}) | CODOM={pedido.codom} CODUG={pedido.codug} FONTE={correspondencia_om['fonte']}")
                    
                    # CORREÇÃO: Processar NDs específicos para cada OM
                    # Usar arrays simples e distribuir igualmente entre as OMs
                    nds = request.form.getlist('nd[]')
                    finalidades_nd = request.form.getlist('finalidade_nd[]')
                    valores_solicitados = request.form.getlist('valor_solicitado[]')
                    valores_aprovados = request.form.getlist('valor_aprovado[]')
                    
                    print(f"📊 Total de NDs recebidos: {len(nds)}")
                    print(f"📊 Total de finalidades: {len(finalidades_nd)}")
                    print(f"📊 Total de valores solicitados: {len(valores_solicitados)}")
                    print(f"📊 Total de valores aprovados: {len(valores_aprovados)}")
                    
                    # CORREÇÃO: Distribuir NDs igualmente entre as OMs
                    total_oms = len(oms)
                    total_nds = len(nds)
                    
                    if total_oms > 0 and total_nds > 0:
                        # Calcular quantos NDs cada OM deve ter
                        nds_por_om = total_nds // total_oms
                        resto = total_nds % total_oms
                        
                        # Calcular índices de início e fim para esta OM
                        inicio = i * nds_por_om + min(i, resto)
                        fim = inicio + nds_por_om + (1 if i < resto else 0)
                        
                        print(f"🔄 Processando NDs {inicio} a {fim-1} para OM {i}")
                        
                        # Processar NDs para esta OM
                        for j in range(inicio, fim):
                            if j < len(nds) and nds[j] and nds[j].strip():
                                try:
                                    finalidade_nd = finalidades_nd[j] if j < len(finalidades_nd) and finalidades_nd[j] else ''
                                    valor_solicitado_str = valores_solicitados[j] if j < len(valores_solicitados) and valores_solicitados[j] else '0'
                                    valor_aprovado_str = valores_aprovados[j] if j < len(valores_aprovados) and valores_aprovados[j] else '0'
                                    
                                    print(f"  🎯 Processando ND {j} para OM {i}: {nds[j]}")
                                    print(f"  📝 Finalidade: {finalidade_nd}")
                                    print(f"  💰 Valor solicitado: {valor_solicitado_str}")
                                    print(f"  💰 Valor aprovado: {valor_aprovado_str}")
                                    
                                    # Remover formatação de moeda
                                    valor_sol_float = converter_valor(valor_solicitado_str)
                                    valor_apr_float = converter_valor(valor_aprovado_str)
                                    
                                    # SE O DESPACHO FOR "APROVADO INTEGRALMENTE", COPIAR VALOR SOLICITADO PARA APROVADO
                                    if solicitacao.despacho and 'aprovado integralmente' in solicitacao.despacho.lower():
                                        valor_apr_float = valor_sol_float
                                        print(f"  🔄 Aprovado integralmente - valor aprovado definido como: {valor_apr_float}")
                                    # SE O DESPACHO FOR "NEGADO", DEFINIR VALOR APROVADO COMO 0
                                    elif solicitacao.despacho and 'negado' in solicitacao.despacho.lower():
                                        valor_apr_float = 0.0
                                        print(f"  🔄 Negado - valor aprovado definido como: 0")
                                    
                                    # Determinar PI baseado na finalidade do ND
                                    pi = PI_POR_FINALIDADE.get(finalidade_nd, '')
                                    
                                    item = ItemPedido(
                                        pedido_id=pedido.id,
                                        nd=nds[j].strip(),
                                        finalidade=finalidade_nd,
                                        pi=pi,
                                        valor_solicitado=valor_sol_float,
                                        valor_aprovado=valor_apr_float
                                    )
                                    db.session.add(item)
                                    print(f"  ✅ Item criado para OM {i}: ND {nds[j]}, Finalidade: {finalidade_nd}, PI: {pi}, Solicitado: {valor_sol_float}, Aprovado: {valor_apr_float}")
                                    
                                except ValueError as e:
                                    print(f"❌ Erro ao processar item {j} para OM {i}: {e}")
                                    continue
                            else:
                                print(f"⚠️  ND {j} inválido ou vazio para OM {i}")
            
            db.session.commit()
            print("✅ Solicitação atualizada com sucesso!")
            
            # GERAR NC OU PRO APÓS COMMIT (para garantir que todos os dados estão salvos)
            if deve_gerar_documento:
                if solicitacao.modalidade == 'crédito':
                    print("🚀 Iniciando geração automática de Nota de Crédito...")
                    nc_gerada = gerar_nc_automatica(solicitacao, descricao_nc_usuario)
                    if nc_gerada:
                        flash(f'Solicitação atualizada com sucesso! Nota de Crédito {nc_gerada.numero} gerada automaticamente.', 'success')
                    else:
                        flash('Solicitação atualizada com sucesso! Erro ao gerar Nota de Crédito automática.', 'warning')
                elif solicitacao.modalidade == 'PRO':
                    print("🚀 Iniciando geração automática de PRO...")
                    pro_gerada = gerar_pro_automatica(solicitacao)
                    if pro_gerada:
                        flash(f'Solicitação atualizada com sucesso! PRO {pro_gerada.numero} gerada automaticamente.', 'success')
                    else:
                        flash('Solicitação atualizada com sucesso! Erro ao gerar PRO automática.', 'warning')
                elif solicitacao.modalidade == 'Transferência Interna COLOG':
                    print("🚀 Iniciando geração automática de Transferência Interna COLOG...")
                    # Aqui você pode adicionar lógica específica se desejar criar um registro/modelo, ou apenas exibir o template
                    return render_template('transferencia_interna_colog.html', solicitacao=solicitacao)
            else:
                flash('Solicitação atualizada com sucesso! Status atualizado automaticamente.', 'success')
            
        except Exception as e:
            db.session.rollback()
            print(f"❌ Erro detalhado na edição: {str(e)}")
            import traceback
            traceback.print_exc()
            flash(f'Erro ao atualizar solicitação: {str(e)}', 'error')
    
    return redirect(url_for('detalhes_solicitacao', id=id))

@app.route('/solicitacao/<int:id>')
@acesso_requerido('admin', 'usuario')
def detalhes_solicitacao(id):
    """Detalhes da solicitação"""
    try:
        solicitacao = db.session.get(SolicitacaoExtraPDRLOG, id)
        if not solicitacao:
            flash('Solicitação não encontrada!', 'error')
            return redirect(url_for('listar_solicitacoes'))
            
        # Buscar PROs relacionadas
        pros_relacionadas = Pro.query.filter_by(solicitacao_pdrlog_id=id).all()

        # IDs dos PROs relacionados
        pro_ids = [pro.id for pro in pros_relacionadas]

        # Buscar NCs diretamente relacionadas e via PROs
        ncs_diretas = NotaCredito.query.filter_by(solicitacao_pdrlog_id=id).all()
        ncs_via_pro = NotaCredito.query.filter(NotaCredito.pro_id.in_(pro_ids)).all() if pro_ids else []

        # Combinar e remover duplicatas
        ncs_relacionadas = list({nc.id: nc for nc in ncs_diretas + ncs_via_pro}.values())

        # Converter valores dos itens para string com ponto decimal
        for pedido in solicitacao.pedidos:
            for item in pedido.itens:
                if item.valor_solicitado is not None:
                    item.valor_solicitado_str = f'{item.valor_solicitado:.2f}'
                else:
                    item.valor_solicitado_str = '0.00'
                if item.valor_aprovado is not None:
                    item.valor_aprovado_str = f'{item.valor_aprovado:.2f}'
                else:
                    item.valor_aprovado_str = '0.00'

        # Verificar se possui finalidades PASA-DEC ou Mnt OP-DEC
        possui_finalidade_especial = False
        for pedido in solicitacao.pedidos:
            for item in pedido.itens:
                if item.finalidade in ['PASA-DEC', 'Mnt OP-DEC', 'Câmaras frigoríficas']:
                    possui_finalidade_especial = True
                    break
            if possui_finalidade_especial:
                break

        # Verificar condições para PASA-DEC/Mnt OP-DEC/Câmaras frigoríficas
        condicoes_especiais_atendidas = True
        if possui_finalidade_especial:
            if not solicitacao.tem_fsv or (solicitacao.tem_fsv == 'sim' and (not solicitacao.diex_dom or not solicitacao.nr_opus)):
                condicoes_especiais_atendidas = False

        return render_template('detalhes_solicitacao.html', 
                    solicitacao=solicitacao, 
                    ncs_relacionadas=ncs_relacionadas,
                    pros_relacionadas=pros_relacionadas,
                    finalidades=FINALIDADES,
                    status_list=STATUS_SOLICITACAO,
                    nd_possiveis=ND_POSSIVEIS,
                    orgaos_demandantes=ORGAOS_DEMANDANTES,
                    oms_data=_obter_oms_data(),
                    modalidades=MODALIDADES,
                    despachos_opcoes=DESPACHOS_OPCOES,
                    pi_por_finalidade=PI_POR_FINALIDADE,
                    possui_finalidade_especial=possui_finalidade_especial,
                    condicoes_especiais_atendidas=condicoes_especiais_atendidas)
    except Exception as e:
        flash(f'Erro ao carregar solicitação: {str(e)}', 'error')
        return redirect(url_for('listar_solicitacoes'))

@app.route('/solicitacao/<int:id>/excluir', methods=['POST'])
@acesso_requerido('admin')
def excluir_solicitacao(id):
    """Exclui uma solicitação - apenas admin"""
    try:
        solicitacao = db.session.get(SolicitacaoExtraPDRLOG, id)
        if not solicitacao:
            flash('Solicitação não encontrada!', 'error')
            return redirect(url_for('listar_solicitacoes'))
            
        numero_solicitacao = solicitacao.numero
        
        print(f"🗑️  Iniciando exclusão da solicitação {numero_solicitacao}")
        
        # Verificar se há NC relacionada
        nc_relacionada = NotaCredito.query.filter_by(solicitacao_pdrlog_id=id).first()
        if nc_relacionada:
            flash('Não é possível excluir a solicitação pois existe uma Nota de Crédito relacionada!', 'error')
            return redirect(url_for('detalhes_solicitacao', id=id))
        
        # Verificar se há PRO relacionada
        pro_relacionada = Pro.query.filter_by(solicitacao_pdrlog_id=id).first()
        if pro_relacionada:
            flash('Não é possível excluir a solicitação pois existe uma PRO relacionada!', 'error')
            return redirect(url_for('detalhes_solicitacao', id=id))
        
        # SIMPLES: Deletar a solicitação (o cascade deve cuidar do resto)
        db.session.delete(solicitacao)
        db.session.commit()
        
        print(f"✅ Solicitação {numero_solicitacao} excluída com sucesso!")
        flash(f'Solicitação {numero_solicitacao} excluída com sucesso!', 'success')
        
    except Exception as e:
        db.session.rollback()
        print(f"❌ Erro ao excluir solicitação: {str(e)}")
        import traceback
        traceback.print_exc()
        flash(f'Erro ao excluir solicitação: {str(e)}', 'error')
    
    return redirect(url_for('listar_solicitacoes'))

# ===== ROTAS DE NOTAS DE CRÉDITO (ACESSO LIVRE PARA TODOS) =====

@app.route('/notas_credito')
@login_required
def listar_ncs():
    global DADOS_PLANILHA
    """Lista todas as Notas de Crédito - todos os usuários"""
    if obter_modulo_menu() == MENU_MODULO_PDR:
        try:
            solicitacoes = _carregar_solicitacoes_pdr_log_planilha()
            incluir_duplicadas = str(request.args.get('incluir_duplicadas', '0') or '0').strip().lower() in ('1', 'true', 'on', 'sim')
            resumo_dedup = {'total_entrada': len(solicitacoes), 'total_saida': len(solicitacoes), 'duplicadas_removidas': 0}
            if not incluir_duplicadas:
                solicitacoes, resumo_dedup = _deduplicar_solicitacoes_para_nc_pdr(solicitacoes)
            om_filter = request.args.get('om', '').strip()
            nd_filter = normalizar_nd(request.args.get('nd', '').strip())
            pi_filter = normalizar_pi(request.args.get('pi', '').strip())
            nc_siafi_filter = re.sub(r'\D', '', request.args.get('nc_siafi', '').strip())
            codug_filter = re.sub(r'\D', '', request.args.get('codug', '').strip())
            codom_filter = re.sub(r'\D', '', request.args.get('codom', '').strip())
            status_filter = request.args.get('status', '').strip()
            filtro_vazios = str(request.args.get('filtro_vazios', '0') or '0').strip().lower() in ('1', 'true', 'on', 'sim')

            def _campo_codigos_contem(campo, codigo):
                if not codigo:
                    return True
                tokens = [re.sub(r'\D', '', str(parte or '')) for parte in str(campo or '').split(',')]
                return any(token == codigo for token in tokens if token)

            ncs_agrupadas_raw = _agrupar_notas_credito_pdr(solicitacoes=solicitacoes, incluir_solicitacoes=False)
            codug_opcoes = set()
            codom_opcoes = set()
            pi_opcoes = set()
            nd_opcoes = set()
            for item in ncs_agrupadas_raw:
                pi_item = normalizar_pi(item.get('pi', ''))
                nd_item = normalizar_nd(item.get('nd', ''))
                if pi_item:
                    pi_opcoes.add(pi_item)
                if nd_item:
                    nd_opcoes.add(nd_item)
                for parte in str(item.get('codug', '') or '').split(','):
                    codigo = re.sub(r'\D', '', parte)
                    if codigo:
                        codug_opcoes.add(codigo)
                for parte in str(item.get('codom', '') or '').split(','):
                    codigo = re.sub(r'\D', '', parte)
                    if codigo:
                        codom_opcoes.add(codigo)

            ncs_agrupadas = []
            for item in ncs_agrupadas_raw:
                if om_filter and om_filter.lower() not in item['om_solicitante'].lower():
                    continue
                if nd_filter and normalizar_nd(item['nd']) != nd_filter:
                    continue
                if pi_filter and normalizar_pi(item['pi']) != pi_filter:
                    continue
                if nc_siafi_filter and nc_siafi_filter not in item.get('nc_siafi', ''):
                    continue
                if codug_filter and not _campo_codigos_contem(item.get('codug', ''), codug_filter):
                    continue
                if codom_filter and not _campo_codigos_contem(item.get('codom', ''), codom_filter):
                    continue
                if status_filter and item.get('status', '') != status_filter:
                    continue
                if filtro_vazios:
                    if (item.get('ref_sisnc') or '').strip() or (item.get('nc_siafi') or '').strip():
                        continue
                ncs_agrupadas.append(item)

            ncs_agrupadas.sort(key=lambda x: (x['status'] != 'NC SIAFI gerada', x['om_solicitante']))

            if (not incluir_duplicadas) and resumo_dedup.get('duplicadas_removidas', 0) > 0:
                flash(
                    'Lista de NC PDR Log atualizada sem duplicidades: '
                    f"{resumo_dedup.get('duplicadas_removidas', 0)} solicitação(ões) duplicada(s) excluída(s).",
                    'info'
                )

            return render_template(
                'notas_credito_pdr_log.html',
                ncs=ncs_agrupadas,
                om_filter=om_filter,
                nd_filter=request.args.get('nd', '').strip(),
                pi_filter=request.args.get('pi', '').strip(),
                nc_siafi_filter=request.args.get('nc_siafi', '').strip(),
                codug_filter=request.args.get('codug', '').strip(),
                codom_filter=request.args.get('codom', '').strip(),
                status_filter=status_filter,
                incluir_duplicadas=incluir_duplicadas,
                codug_opcoes=sorted(codug_opcoes),
                codom_opcoes=sorted(codom_opcoes),
                pi_opcoes=sorted(pi_opcoes),
                nd_opcoes=sorted(nd_opcoes),
                filtro_vazios=filtro_vazios
            )
        except Exception as e:
            flash(f'Erro ao carregar Notas de Crédito do PDR Log: {str(e)}', 'error')
            return render_template(
                'notas_credito_pdr_log.html',
                ncs=[],
                om_filter='',
                nd_filter='',
                pi_filter='',
                nc_siafi_filter='',
                codug_filter='',
                codom_filter='',
                status_filter='',
                incluir_duplicadas=False,
                codug_opcoes=[],
                codom_opcoes=[],
                pi_opcoes=[],
                nd_opcoes=[],
                filtro_vazios=False
            )

    try:
        status_filter = request.args.get('status', '')
        diex_filter = request.args.get('diex', '')
        orgao_filter = request.args.get('orgao', '')
        pi_filter = request.args.get('pi', '')
        nd_filter = request.args.get('nd', '')
        sigla_ug_filter = request.args.get('sigla_ug', '')
        om_filter = request.args.get('om', '')
        ref_sisnc_filter = request.args.get('ref_sisnc', '')
        nc_siafi_filter = request.args.get('nc_siafi', '')
        cod_ug_filter = request.args.get('cod_ug', '')

        query = NotaCredito.query.options(
            db.joinedload(NotaCredito.usuario),
            db.joinedload(NotaCredito.solicitacao_origem)
        )

        if status_filter:
            query = query.filter_by(status=status_filter)
        if diex_filter:
            query = query.join(SolicitacaoExtraPDRLOG).filter(
                SolicitacaoExtraPDRLOG.diex.ilike(f'%{diex_filter}%')
            )
        if orgao_filter:
            query = query.join(SolicitacaoExtraPDRLOG).filter(
                SolicitacaoExtraPDRLOG.orgao_demandante.ilike(f'%{orgao_filter}%')
            )
        if pi_filter:
            query = query.filter(NotaCredito.pi.ilike(f'%{pi_filter}%'))
        if nd_filter:
            query = query.filter(NotaCredito.nd.ilike(f'%{nd_filter}%'))
        if sigla_ug_filter:
            query = query.join(SolicitacaoExtraPDRLOG, NotaCredito.solicitacao_pdrlog_id == SolicitacaoExtraPDRLOG.id)
            query = query.filter(
                SolicitacaoExtraPDRLOG.id.in_(
                    db.session.query(PedidoSolicitacao.solicitacao_id)
                    .filter(PedidoSolicitacao.sigla_ug.ilike(f'%{sigla_ug_filter}%'))
                )
            )
        if om_filter:
            query = query.join(SolicitacaoExtraPDRLOG, NotaCredito.solicitacao_pdrlog_id == SolicitacaoExtraPDRLOG.id)
            query = query.filter(
                SolicitacaoExtraPDRLOG.id.in_(
                    db.session.query(PedidoSolicitacao.solicitacao_id)
                    .filter(PedidoSolicitacao.om.ilike(f'%{om_filter}%'))
                )
            )
        if ref_sisnc_filter:
            query = query.filter(NotaCredito.ref_sisnc.ilike(f'%{ref_sisnc_filter}%'))
        if nc_siafi_filter:
            query = query.filter(NotaCredito.nc_siafi.ilike(f'%{nc_siafi_filter}%'))
        if cod_ug_filter:
            query = query.filter(NotaCredito.cod_ug.ilike(f'%{cod_ug_filter}%'))

        ncs = query.order_by(NotaCredito.data_criacao.desc()).all()

        return render_template('notas_credito.html',
            ncs=ncs,
            status_filter=status_filter,
            diex_filter=diex_filter,
            orgao_filter=orgao_filter,
            pi_filter=pi_filter,
            nd_filter=nd_filter,
            sigla_ug_filter=sigla_ug_filter,
            om_filter=om_filter,
            ref_sisnc_filter=ref_sisnc_filter,
            nc_siafi_filter=nc_siafi_filter,
            cod_ug_filter=cod_ug_filter,
            status_list=STATUS_NC,
            pis_disponiveis=list(PI_POR_FINALIDADE.values()),
            DADOS_PLANILHA=globals().get('DADOS_PLANILHA', {'OMs': []})
        )

    except Exception as e:
        flash(f'Erro ao carregar Notas de Crédito: {str(e)}', 'error')
        return render_template('notas_credito.html',
            ncs=[],
            status_filter='',
            diex_filter='',
            orgao_filter='',
            pi_filter='',
            nd_filter='',
            sigla_ug_filter='',
            om_filter='',
            ref_sisnc_filter='',
            nc_siafi_filter='',
            cod_ug_filter='',
            status_list=STATUS_NC,
            pis_disponiveis=[],
            DADOS_PLANILHA=globals().get('DADOS_PLANILHA', {'OMs': []})
        )


@app.route('/pdr-log/notas-credito/grupo')
@login_required
def detalhes_ncs_pdr_grupo():
    if obter_modulo_menu() != MENU_MODULO_PDR:
        return redirect(url_for('listar_ncs'))

    om = request.args.get('om', '').strip()
    pi = normalizar_pi(request.args.get('pi', '').strip())
    nd = normalizar_nd(request.args.get('nd', '').strip())
    incluir_duplicadas = str(request.args.get('incluir_duplicadas', '0') or '0').strip().lower() in ('1', 'true', 'on', 'sim')

    if not om or not pi or not nd:
        flash('Parâmetros inválidos para detalhar o grupo.', 'error')
        return redirect(url_for('listar_ncs'))

    try:
        solicitacoes = _carregar_solicitacoes_pdr_log_planilha()
        if not incluir_duplicadas:
            solicitacoes, _ = _deduplicar_solicitacoes_para_nc_pdr(solicitacoes)
    except Exception as e:
        flash(f'Erro ao carregar solicitações do PDR Log: {str(e)}', 'error')
        return redirect(url_for('listar_ncs'))

    detalhes = []
    for solicitacao in solicitacoes:
        pi_mapeada = _mapear_pi_pdr_por_assunto(solicitacao.get('nome_assunto', ''))
        nd_mapeada = normalizar_nd(solicitacao.get('cod_nd', ''))
        om_item = str(solicitacao.get('om_solicitante', '') or '').strip()

        if om_item != om or pi_mapeada != pi or nd_mapeada != nd:
            continue

        detalhes.append(solicitacao)

    detalhes.sort(key=lambda x: (x.get('data_solicitacao', ''), x.get('numero', '')), reverse=True)

    return render_template(
        'notas_credito_pdr_grupo.html',
        om=om,
        pi=pi,
        nd=nd,
        solicitacoes=detalhes,
        total_solicitacoes=len(detalhes),
        total_valor=round(sum(float(s.get('valor_total', 0) or 0) for s in detalhes), 2)
    )


@app.route('/auditoria', methods=['GET', 'POST'])
@acesso_requerido('admin', 'usuario')
def auditoria():
    """Auditoria entre o relatorio de NCs do sistema e arquivo do SIAFI."""
    if request.args.get('debug') == '1':
        return make_response(f"DEBUG auditoria OK | {os.path.abspath(__file__)}")
    resultados = []
    nao_encontradas = []
    quase_matches = []
    total_matches = 0
    total_atualizadas = 0
    total_sem_alteracao = 0
    total_sem_nc_siafi = 0
    total_preenchidas_auto = 0
    total_upload_valor = 0.0
    total_sistema_valor = 0.0
    total_diferenca = 0.0
    duplicidades_sistema = []
    total_duplicidades_sistema = 0
    inconsistencias_tres_fontes = []
    inconsistencias_tg_sisnc = []
    resumo_tres_fontes = {
        'total_sistema': 0,
        'total_tg': 0,
        'total_sisnc': 0,
        'sistema_ok_nas_tres': 0,
        'sistema_sem_tg': 0,
        'sistema_sem_sisnc': 0,
        'tg_sem_sistema': 0,
        'sisnc_sem_sistema': 0,
        'divergencia_tg_sisnc': 0,
        'tg_sisnc_chaves_comuns': 0,
        'tg_sisnc_somente_tg': 0,
        'tg_sisnc_somente_sisnc': 0,
        'tg_sisnc_valor_divergente': 0
    }
    diagnostico = {
        'linhas_upload': 0,
        'sem_codom': 0,
        'sem_pi': 0,
        'sem_nd': 0,
        'sem_cod_ug': 0,
        'sem_valor': 0,
        'sem_diex': 0,
        'match_sem_codom': 0,
        'match_sem_valor': 0,
        'match_sem_nd': 0,
        'match_sem_pi': 0,
        'match_sem_cod_ug': 0,
        'ncs_sistema_total': 0,
        'total_sistema_valor': 0.0,
        'total_upload_valor': 0.0,
        'pi_col_idx': None,
        'valor_col_idx': None,
        'keys_sistema': 0,
        'keys_upload': 0,
        'sample_sys_key': None,
        'sample_upload_key': None,
        'sample_sistema_valor_raw': None,
        'sample_sistema_valor_norm': None,
        'sample_upload_valor_raw': None,
        'sample_upload_valor_norm': None,
        'sistema_valores_count': 0,
        'sistema_valores_nonzero': 0,
        'sistema_valores_nulos': 0,
        'sistema_valores_zero': 0,
        'sistema_valor_tipo': None,
        'sistema_valores_min': None,
        'sistema_valores_max': None,
        'sistema_valores_sum_sql': None,
        'upload_valores_nonzero': 0,
        'upload_valores_min': None,
        'upload_valores_max': None,
        'upload_col_total': None,
        'upload_col_nonzero': None,
        'db_path': None,
        'db_path_raw': None,
        'db_path_abs': None,
        'db_file_exists': None,
        'db_url': None,
        'db_uri_config': None,
        'instance_path': None,
        'app_cwd': None,
        'upload_shape': None,
        'upload_cols': None,
        'upload_pi_best_col': None,
        'upload_valor_best_col': None,
        'upload_valor_nonzero_best': None,
        'upload_valor_nonzero_7': None,
        'upload_pi_nonzero_5': None,
        'sistema_shape': None,
        'sistema_col_j_nonzero': None
    }

    if request.method == 'POST':
        arquivo = request.files.get('arquivo')
        total_sistema_relatorio = None
        if arquivo and arquivo.filename:
            try:
                df_upload = pd.read_excel(
                    arquivo,
                    header=None,
                    skiprows=5,
                    converters={
                        0: str,
                        1: str,
                        4: str,
                        5: str,
                        6: str,
                        7: str
                    }
                )
            except Exception as e:
                flash(f'Erro ao ler o Excel: {str(e)}', 'error')
                return redirect(url_for('auditoria'))
        else:
            try:
                # Unifica a lógica: carrega o cache SISNC e TG do extra pdr log usando o mesmo pipeline do pdr log
                df_upload = _carregar_upload_auditoria_sisnc_cache('extra')
                df_tg = _carregar_upload_auditoria_cache('extra')
            except FileNotFoundError:
                flash('Bases de auditoria não encontradas (TG/SISNC). Envie as planilhas no Painel Subsistência.', 'error')
                return redirect(url_for('painel_subsistencia'))
            except Exception as e:
                flash(f'Erro ao carregar bases da auditoria (TG/SISNC): {str(e)}', 'error')
                return redirect(url_for('auditoria'))

        # Normalização e filtragem igual ao pdr log
        df_upload = _normalizar_colunas_numericas(df_upload)
        df_upload = df_upload.dropna(how='all')
        if 0 in df_upload.columns:
            df_upload = df_upload[df_upload[0].notna()]
        diagnostico['upload_shape'] = f"{df_upload.shape[0]}x{df_upload.shape[1]}"
        diagnostico['upload_cols'] = list(df_upload.columns)[:12]

        # Garante que a comparação SISNC x TG do extra pdr log use o mesmo pipeline do pdr log
        # (construir_mapa_reduzido_upload, construir_mapa_tg_sisnc_sem_valor, etc.)
        # O restante do fluxo permanece igual, pois já utiliza as funções compartilhadas

        df_sistema = gerar_relatorio_ncs_dataframe()
        if not df_sistema.empty and 'Valor NC' in df_sistema.columns:
            diagnostico['sistema_shape'] = f"{df_sistema.shape[0]}x{df_sistema.shape[1]}"
            valores_sistema_col = df_sistema['Valor NC'].apply(normalizar_valor)
            diagnostico['sistema_col_j_nonzero'] = int((valores_sistema_col > 0).sum())
            total_sistema_relatorio = round(float(valores_sistema_col.sum()), 2)
            total_sistema_valor = total_sistema_relatorio
            diagnostico['total_sistema_valor'] = total_sistema_valor

        # Detect column indexes for PI and Valor based on content patterns.
        # Padrao conhecido: PI na coluna F (5) e Valor na coluna H (7).
        pi_col_idx = 5
        valor_col_idx = 7

        def contar_pi_col(col_idx):
            try:
                return df_upload[col_idx].astype(str).str.contains(
                    r'\bE\d[A-Z0-9]{8,}\b',
                    case=False,
                    na=False
                ).sum()
            except Exception:
                return 0

        def pontuar_valor_col(col_idx):
            try:
                serie = df_upload[col_idx]
            except Exception:
                return 0

            count_decimal = 0
            count_reasonable = 0
            count_huge = 0
            count_pi_like = 0
            count_codug_like = 0
            count_nd_like = 0
            count_six_digit = 0
            valores = []
            for v in serie:
                texto = str(v or '')
                if re.search(r'\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2}', texto):
                    count_decimal += 1
                if re.search(r'\bE\d[A-Z0-9]{8,}\b', texto.upper()):
                    count_pi_like += 1
                digitos = re.sub(r'\D', '', texto)
                if len(digitos) == 6:
                    count_six_digit += 1
                    if digitos.startswith('16'):
                        count_codug_like += 1
                    if digitos.startswith(('33', '44')):
                        count_nd_like += 1
                valor = normalizar_valor(v)
                if valor > 0:
                    valores.append(valor)
                    if valor < 1_000_000_000:
                        count_reasonable += 1
                    elif valor >= 1_000_000_000_000:
                        count_huge += 1

            # Prefer columns with typical monetary ranges; penalize huge ids.
            total = len(serie) if len(serie) else 1
            score = (count_decimal * 3) + count_reasonable - (count_huge * 2)

            if count_pi_like / total > 0.2:
                score -= 8
            if count_codug_like / total > 0.5:
                score -= 6
            if count_nd_like / total > 0.5:
                score -= 6
            if count_six_digit / total > 0.7:
                score -= 3

            if valores:
                spread = max(valores) - min(valores)
                if spread >= 1_000:
                    score += 2
                if spread >= 100_000:
                    score += 2

            return score

        def contar_valores_nonzero(col_idx):
            try:
                serie = df_upload[col_idx]
            except Exception:
                return 0
            try:
                valores = serie.apply(normalizar_valor)
            except Exception:
                return 0
            return int((valores > 0).sum())

        if len(df_upload.columns) > 0:
            pi_counts = {col: contar_pi_col(col) for col in df_upload.columns}
            pi_best_col = max(pi_counts, key=pi_counts.get)
            if pi_counts.get(pi_best_col, 0) > 0:
                pi_col_idx = pi_best_col

            valor_scores = {col: pontuar_valor_col(col) for col in df_upload.columns}
            valor_best_col = max(valor_scores, key=valor_scores.get)
            if valor_scores.get(valor_best_col, 0) > 0:
                valor_col_idx = valor_best_col

            if 5 in df_upload.columns:
                pi_nonzero_5 = contar_pi_col(5)
                diagnostico['upload_pi_nonzero_5'] = pi_nonzero_5
                if pi_nonzero_5 > 0:
                    pi_col_idx = 5
            if 7 in df_upload.columns:
                nonzero_7 = contar_valores_nonzero(7)
                diagnostico['upload_valor_nonzero_7'] = nonzero_7
                if nonzero_7 > 0:
                    valor_col_idx = 7

            diagnostico['upload_pi_best_col'] = pi_best_col
            diagnostico['upload_valor_best_col'] = valor_best_col
            diagnostico['upload_valor_nonzero_best'] = contar_valores_nonzero(valor_best_col)

        diagnostico['pi_col_idx'] = pi_col_idx
        diagnostico['valor_col_idx'] = valor_col_idx
        if valor_col_idx in df_upload.columns:
            valores_coluna = df_upload[valor_col_idx].apply(normalizar_valor)
            diagnostico['upload_col_total'] = round(float(valores_coluna.sum()), 2)
            diagnostico['upload_col_nonzero'] = int((valores_coluna > 0).sum())

        ncs_sistema = NotaCredito.query.options(
            db.joinedload(NotaCredito.solicitacao_origem)
        ).all()
        diagnostico['ncs_sistema_total'] = len(ncs_sistema)
        engine_url = db.engine.url
        db_path_raw = engine_url.database
        diagnostico['db_path_raw'] = db_path_raw
        if db_path_raw:
            diagnostico['db_path_abs'] = os.path.abspath(db_path_raw)
            diagnostico['db_file_exists'] = os.path.exists(diagnostico['db_path_abs'])
        else:
            diagnostico['db_path_abs'] = None
            diagnostico['db_file_exists'] = None
        diagnostico['db_path'] = db_path_raw or str(engine_url)
        diagnostico['db_url'] = str(engine_url)
        diagnostico['db_uri_config'] = repr(current_app.config.get('SQLALCHEMY_DATABASE_URI'))
        diagnostico['instance_path'] = current_app.instance_path
        diagnostico['app_cwd'] = os.getcwd()
        try:
            from sqlalchemy import func
            soma_sql = db.session.query(func.coalesce(func.sum(NotaCredito.valor), 0)).scalar()
            diagnostico['sistema_valores_sum_sql'] = float(soma_sql or 0)
        except Exception:
            diagnostico['sistema_valores_sum_sql'] = None
        valores_sistema = []
        for nc in ncs_sistema:
            if nc.valor is None or pd.isna(nc.valor):
                diagnostico['sistema_valores_nulos'] += 1
                valores_sistema.append(0.0)
                continue

            if diagnostico['sistema_valor_tipo'] is None:
                diagnostico['sistema_valor_tipo'] = type(nc.valor).__name__

            valor_norm = normalizar_valor(nc.valor)
            valores_sistema.append(valor_norm)
            if valor_norm == 0:
                diagnostico['sistema_valores_zero'] += 1
        if total_sistema_relatorio is None:
            total_sistema_valor = round(sum(valores_sistema), 2)
            diagnostico['total_sistema_valor'] = total_sistema_valor
        diagnostico['sistema_valores_count'] = len(valores_sistema)
        valores_sistema_nonzero = [v for v in valores_sistema if v > 0]
        diagnostico['sistema_valores_nonzero'] = len(valores_sistema_nonzero)
        if valores_sistema_nonzero:
            diagnostico['sistema_valores_min'] = min(valores_sistema_nonzero)
            diagnostico['sistema_valores_max'] = max(valores_sistema_nonzero)
        for nc in ncs_sistema:
            if nc.valor is not None:
                diagnostico['sample_sistema_valor_raw'] = nc.valor
                diagnostico['sample_sistema_valor_norm'] = normalizar_valor(nc.valor)
                break

        nd_possiveis_digitos = {normalizar_nd(nd) for nd in ND_POSSIVEIS}

        indice_ncs = {}
        indice_sem_codom = {}
        indice_sem_valor = {}
        indice_sem_nd = {}
        indice_sem_pi = {}
        indice_sem_cod_ug = {}
        for nc in ncs_sistema:
            solicitacao = nc.solicitacao_origem
            diex_raw_solic = solicitacao.diex if solicitacao else ''
            diex_num = (
                extrair_numero_diex(nc.descricao)
                or extrair_numero_diex(nc.diex_credito)
                or extrair_numero_diex(nc.numero)
                or extrair_numero_diex(diex_raw_solic)
                or extrair_numero_diex(solicitacao.descricao if solicitacao else '')
            )
            codom_norm = extrair_codom_sistema_nc(nc)
            pi_norm = normalizar_pi(nc.pi)
            nd_norm = normalizar_nd(nc.nd)
            cod_ug_norm = normalizar_cod_ug(nc.cod_ug)
            valor_norm = valor_para_chave(nc.valor)
            chave = (codom_norm, cod_ug_norm, pi_norm, nd_norm, valor_norm)
            indice_ncs.setdefault(chave, []).append(nc)
            indice_sem_codom.setdefault((cod_ug_norm, pi_norm, nd_norm, valor_norm), []).append(nc)
            indice_sem_valor.setdefault((codom_norm, cod_ug_norm, pi_norm, nd_norm), []).append(nc)
            indice_sem_nd.setdefault((codom_norm, cod_ug_norm, pi_norm, valor_norm), []).append(nc)
            indice_sem_pi.setdefault((codom_norm, cod_ug_norm, nd_norm, valor_norm), []).append(nc)
            indice_sem_cod_ug.setdefault((codom_norm, pi_norm, nd_norm, valor_norm), []).append(nc)

        # Detecta duplicidades na chave principal; o DIEx sera usado apenas para desempate.
        for chave, ncs in indice_ncs.items():
            if len(ncs) > 1:
                duplicidades_sistema.append({
                    'chave': chave,
                    'ncs': ncs
                })
        total_duplicidades_sistema = sum(len(item['ncs']) for item in duplicidades_sistema)

        diagnostico['keys_sistema'] = len(indice_ncs)
        if indice_ncs:
            diagnostico['sample_sys_key'] = next(iter(indice_ncs.keys()))

        MAX_QUASE_MATCHES = 300
        quase_match_ids = set()

        def extrair_chave_sistema_nc(nc_ref):
            solicitacao_ref = nc_ref.solicitacao_origem
            diex_raw_solic_ref = solicitacao_ref.diex if solicitacao_ref else ''
            diex_num_ref = (
                extrair_numero_diex(nc_ref.descricao)
                or extrair_numero_diex(nc_ref.diex_credito)
                or extrair_numero_diex(nc_ref.numero)
                or extrair_numero_diex(diex_raw_solic_ref)
                or extrair_numero_diex(solicitacao_ref.descricao if solicitacao_ref else '')
            )
            codom_ref = extrair_codom_sistema_nc(nc_ref)
            return {
                'cod_om': codom_ref,
                'cod_ug': normalizar_cod_ug(nc_ref.cod_ug),
                'pi': normalizar_pi(nc_ref.pi),
                'nd': normalizar_nd(nc_ref.nd),
                'valor': valor_para_chave(nc_ref.valor),
                'diex': diex_num_ref,
                'numero_nc': nc_ref.numero
            }

        def filtrar_matches_por_diex(matches_ref, diex_upload_ref):
            if len(matches_ref) <= 1:
                return matches_ref
            if not diex_upload_ref:
                return []

            candidatos = []
            for nc_ref in matches_ref:
                dados_ref = extrair_chave_sistema_nc(nc_ref)
                if dados_ref.get('diex') == diex_upload_ref:
                    candidatos.append(nc_ref)

            # Confirma DIEx somente quando resolve a duplicidade de forma univoca.
            if len(candidatos) == 1:
                return candidatos
            return []

        def extrair_valor_linha(linha, valor_idx):
            valor_raw = linha.get(valor_idx, 0)
            valor_base = normalizar_valor(valor_raw)
            if valor_base > 0:
                return valor_base, valor_raw

            # Fallback: aceita formatos com 1-2 casas decimais (ex.: 2297.0 / 2297,00).
            for valor in linha.values:
                texto = str(valor or '')
                if re.search(r'\d{1,3}(?:\.\d{3})*,\d{1,2}|\d+,\d{1,2}|\d+\.\d{1,2}', texto):
                    candidato = normalizar_valor(texto)
                    if candidato > 0:
                        return candidato, valor
            return 0.0, None

        def extrair_info_base_upload(row, pi_idx=5, valor_idx=7):
            cod_ug_upload_norm = normalizar_cod_ug(row.get(1, ''))
            nd_upload_norm = normalizar_nd(row.get(3, row.get(4, '')))
            pi_upload_norm = normalizar_pi(row.get(pi_idx, ''))
            valor_upload_norm = valor_para_chave(row.get(valor_idx, 0))

            if not nd_upload_norm:
                for valor in row.values:
                    candidato = normalizar_nd(valor)
                    if candidato and candidato in nd_possiveis_digitos:
                        nd_upload_norm = candidato
                        break

            if not cod_ug_upload_norm:
                for valor in row.values:
                    candidato = normalizar_cod_ug(valor)
                    if len(candidato) == 6:
                        cod_ug_upload_norm = candidato
                        break

            if not pi_upload_norm:
                for valor in row.values:
                    texto = str(valor or '')
                    match = re.search(r'\bE\d[A-Z0-9]{8,}\b', texto.upper())
                    if match:
                        pi_upload_norm = match.group(0)
                        break

            if valor_upload_norm == '0.00':
                valor_upload_float, _ = extrair_valor_linha(row, valor_idx)
                valor_upload_norm = valor_para_chave(valor_upload_float)

            chave_reduzida = (cod_ug_upload_norm, pi_upload_norm, nd_upload_norm, valor_upload_norm)
            return {
                'cod_ug': cod_ug_upload_norm,
                'pi': pi_upload_norm,
                'nd': nd_upload_norm,
                'valor': valor_upload_norm,
                'chave_reduzida': chave_reduzida,
                'nc_upload': row.get(0, ''),
                'referencia_sisnc': str(row.get(2, '') or '').strip()
            }

        def construir_mapa_reduzido_upload(df_fonte):
            mapa = {}
            if df_fonte is None or df_fonte.empty:
                return mapa

            for _, row in df_fonte.iterrows():
                info = extrair_info_base_upload(row, pi_idx=5, valor_idx=7)
                chave = info['chave_reduzida']
                if not info['cod_ug'] or not info['pi'] or not info['nd'] or info['valor'] == '0.00':
                    continue
                if chave not in mapa:
                    mapa[chave] = info
            return mapa

        def eh_recolhimento_sisnc_extra(doc_texto):
            """Detecta lançamentos SISNC de recolhimento para abatimento de saldo."""
            doc = _normalizar_texto_planilha(doc_texto)
            if not doc or 'RECOLHIMENTO' not in doc:
                return False
            return (
                'DUPLICIDADE' in doc
                or 'CREDITO' in doc
                or 'RECOLHIMENTO DE CREDITO' in doc
                or 'RECOLHIMENTO POR DUPLICIDADE' in doc
            )

        def extrair_doc_upload_extra(row):
            # Prioriza coluna de documento/referência; depois faz fallback textual.
            candidatos_coluna = [6, 2]
            for col in candidatos_coluna:
                texto = str(row.get(col, '') or '').strip()
                if texto and texto.lower() not in ('nan', 'none', 'null', '--'):
                    return texto

            for valor in row.values:
                texto = str(valor or '').strip()
                if not texto or texto.lower() in ('nan', 'none', 'null', '--'):
                    continue
                if any(token in _normalizar_texto_planilha(texto) for token in ('DIEX', 'RECOLHIMENTO', 'DESCENTRALIZACAO')):
                    return texto
            return ''

        def construir_mapa_tg_sisnc_sem_valor(df_fonte, fonte_nome=''):
            mapa = {}
            if df_fonte is None or df_fonte.empty:
                return mapa

            fonte_norm = _normalizar_texto_planilha(fonte_nome)
            eh_fonte_sisnc = 'SISNC' in fonte_norm

            # Para evitar soma duplicada, inclui referencia_sisnc (ou campo equivalente) na chave
            vistos = set()
            for _, row in df_fonte.iterrows():
                info = extrair_info_base_upload(row, pi_idx=5, valor_idx=7)
                cod_ug = info.get('cod_ug', '')
                pi = info.get('pi', '')
                nd = info.get('nd', '')
                doc = extrair_doc_upload_extra(row)
                codom = extrair_codom_numerico(doc)
                referencia_sisnc = info.get('referencia_sisnc', '')
                try:
                    valor_float = float(info.get('valor', '0.00'))
                except Exception:
                    valor_float = 0.0

                if pd.isna(valor_float) or math.isinf(valor_float):
                    valor_float = 0.0

                if valor_float == 0.0:
                    valor_float, _ = extrair_valor_linha(row, 7)

                if pd.isna(valor_float) or math.isinf(valor_float):
                    valor_float = 0.0

                if not cod_ug or not pi or not nd or valor_float == 0.0:
                    continue

                valor_float = abs(valor_float)
                eh_recolhimento = eh_fonte_sisnc and eh_recolhimento_sisnc_extra(doc)
                if eh_recolhimento:
                    valor_float = -valor_float

                # Chave única inclui referencia_sisnc
                chave_sem_valor = (cod_ug, codom, pi, nd, referencia_sisnc)
                if chave_sem_valor in vistos:
                    continue
                vistos.add(chave_sem_valor)
                registro = mapa.setdefault(chave_sem_valor, {
                    'cod_ug': cod_ug,
                    'codom': codom,
                    'pi': pi,
                    'nd': nd,
                    'valor_total': 0.0,
                    'nc_upload': info.get('nc_upload', ''),
                    'referencia_sisnc': referencia_sisnc,
                    'qtd_recolhimentos': 0
                })
                registro['valor_total'] += valor_float
                if eh_recolhimento:
                    registro['qtd_recolhimentos'] += 1
                if codom and not registro.get('codom'):
                    registro['codom'] = codom
                if not registro.get('referencia_sisnc') and referencia_sisnc:
                    registro['referencia_sisnc'] = referencia_sisnc
            return mapa

        for _, row in df_upload.iterrows():
            diagnostico['linhas_upload'] += 1
            fonte_upload = str(row.get('__fonte_upload', 'Upload') or 'Upload')
            nc_upload = row.get(0, '')
            cod_ug_upload = row.get(1, '')
            nd_upload = row.get(3, row.get(4, ''))
            pi_upload = row.get(pi_col_idx, '')
            doc_info = row.get(6, '')
            valor_upload = row.get(valor_col_idx, 0)
            codom_upload_norm = extrair_codom_numerico(doc_info)

            pi_upload_norm = normalizar_pi(pi_upload)
            cod_ug_upload_norm = normalizar_cod_ug(cod_ug_upload)
            valor_upload_norm = valor_para_chave(valor_upload)
            nd_upload_norm = normalizar_nd(nd_upload)
            if not nd_upload_norm:
                # Fallback: tenta localizar ND em outras colunas da linha.
                for valor in row.values:
                    candidato = normalizar_nd(valor)
                    if candidato and candidato in nd_possiveis_digitos:
                        nd_upload_norm = candidato
                        break

            if not cod_ug_upload_norm:
                # Fallback: tenta localizar COD UG de 6 digitos na linha.
                for valor in row.values:
                    candidato = normalizar_cod_ug(valor)
                    if len(candidato) == 6:
                        cod_ug_upload_norm = candidato
                        break

            if not pi_upload_norm:
                # Fallback: tenta localizar PI padrao na linha (ex: E6SUSOLA5CF).
                for valor in row.values:
                    texto = str(valor or '')
                    match = re.search(r'\bE\d[A-Z0-9]{8,}\b', texto.upper())
                    if match:
                        pi_upload_norm = match.group(0)
                        break

            if not codom_upload_norm:
                for valor in row.values:
                    candidato = extrair_codom_numerico(valor)
                    if candidato:
                        codom_upload_norm = candidato
                        break

            diex_upload_num = extrair_numero_diex(doc_info)
            if not diex_upload_num:
                # Fallback: tenta localizar DIEx em outras colunas da linha.
                for valor in row.values:
                    candidato = extrair_numero_diex(valor)
                    if candidato:
                        diex_upload_num = candidato
                        break

            # Consolida o valor do upload antes dos matches/quase-matches para evitar
            # exibir 0.00 quando o valor real está em outra coluna da mesma linha.
            valor_upload_raw = valor_upload
            valor_upload_norm_float = normalizar_valor(valor_upload_raw)
            if pd.isna(valor_upload_norm_float) or math.isinf(valor_upload_norm_float):
                valor_upload_norm_float = 0.0
            if valor_upload_norm_float == 0.0:
                valor_upload_norm_float, valor_upload_raw = extrair_valor_linha(row, valor_col_idx)
            if pd.isna(valor_upload_norm_float) or math.isinf(valor_upload_norm_float):
                valor_upload_norm_float = 0.0
            if valor_upload_norm_float > 0.0:
                valor_upload_norm = valor_para_chave(valor_upload_norm_float)

            if not codom_upload_norm:
                diagnostico['sem_codom'] += 1

            if not pi_upload_norm:
                diagnostico['sem_pi'] += 1
            if not nd_upload_norm:
                diagnostico['sem_nd'] += 1
            if not cod_ug_upload_norm:
                diagnostico['sem_cod_ug'] += 1
            if valor_upload_norm == '0.00':
                diagnostico['sem_valor'] += 1
            if not diex_upload_num:
                diagnostico['sem_diex'] += 1

            chave_upload = (codom_upload_norm, cod_ug_upload_norm, pi_upload_norm, nd_upload_norm, valor_upload_norm)

            matches = indice_ncs.get(chave_upload, [])
            if len(matches) > 1:
                matches = filtrar_matches_por_diex(matches, diex_upload_num)

            if not matches and codom_upload_norm and pi_upload_norm and nd_upload_norm and cod_ug_upload_norm:
                # Fallback: tenta localizar valor em outras colunas da linha.
                valor_candidatos = set()
                for valor in row.values:
                    candidato = valor_para_chave(valor)
                    if candidato != '0.00':
                        valor_candidatos.add(candidato)
                for candidato in valor_candidatos:
                    candidatos_matches = indice_ncs.get((codom_upload_norm, cod_ug_upload_norm, pi_upload_norm, nd_upload_norm, candidato), [])
                    if len(candidatos_matches) > 1:
                        candidatos_matches = filtrar_matches_por_diex(candidatos_matches, diex_upload_num)
                    if candidatos_matches:
                        valor_upload_norm = candidato
                        chave_upload = (codom_upload_norm, cod_ug_upload_norm, pi_upload_norm, nd_upload_norm, valor_upload_norm)
                        matches = candidatos_matches
                        break

            try:
                valor_upload_norm_float = float(valor_upload_norm)
            except Exception:
                valor_upload_norm_float = 0.0
            if pd.isna(valor_upload_norm_float) or math.isinf(valor_upload_norm_float):
                valor_upload_norm_float = 0.0
            if valor_upload_norm_float == 0.0:
                valor_upload_norm_float, valor_upload_raw = extrair_valor_linha(row, valor_col_idx)
            if pd.isna(valor_upload_norm_float) or math.isinf(valor_upload_norm_float):
                valor_upload_norm_float = 0.0
            total_upload_valor += valor_upload_norm_float
            if diagnostico['sample_upload_valor_norm'] is None and valor_upload_norm_float > 0:
                diagnostico['sample_upload_valor_raw'] = valor_upload_raw if valor_upload_raw is not None else valor_upload
                diagnostico['sample_upload_valor_norm'] = valor_upload_norm_float
            if valor_upload_norm_float > 0:
                diagnostico['upload_valores_nonzero'] += 1
                if diagnostico['upload_valores_min'] is None:
                    diagnostico['upload_valores_min'] = valor_upload_norm_float
                    diagnostico['upload_valores_max'] = valor_upload_norm_float
                else:
                    diagnostico['upload_valores_min'] = min(diagnostico['upload_valores_min'], valor_upload_norm_float)
                    diagnostico['upload_valores_max'] = max(diagnostico['upload_valores_max'], valor_upload_norm_float)
            if matches:
                nc_siafi = re.sub(r'\D', '', str(nc_upload or ''))
                nc_siafi = nc_siafi[-6:] if nc_siafi else ''
                ref_sisnc = str(row.get(2, '') or '').strip()  # coluna C do SISNC

                for nc in matches:
                    total_matches += 1
                    anterior_nc = nc.nc_siafi or ''
                    anterior_ref = nc.ref_sisnc or ''
                    status = 'Sem NC SIAFI/REF SISNC'

                    atualizado = False
                    # Atualiza NC SIAFI se necessário
                    if nc_siafi and (not anterior_nc or anterior_nc != nc_siafi):
                        nc.nc_siafi = nc_siafi
                        atualizado = True
                    # Atualiza REF SISNC se necessário
                    if ref_sisnc and (not anterior_ref or anterior_ref != ref_sisnc):
                        nc.ref_sisnc = ref_sisnc
                        atualizado = True

                    # Recalcula o status com base nos campos atuais (REF SISNC / NC SIAFI)
                    novo_status = _status_automatico_nc(
                        ref_sisnc=nc.ref_sisnc,
                        nc_siafi=nc.nc_siafi,
                        status_atual=nc.status
                    )

                    # Conta preenchimento automático quando houve alteração dos campos
                    if atualizado:
                        total_preenchidas_auto += 1

                    # Aplica novo status se for diferente do atual
                    if novo_status != (nc.status or ''):
                        nc.status = novo_status
                        total_atualizadas += 1
                        status = 'Atualizada'
                    elif atualizado:
                        # Campos (ref_sisnc/nc_siafi) foram alterados, mas o status já estava consistente
                        total_atualizadas += 1
                        status = 'Atualizada'
                    else:
                        total_sem_alteracao += 1
                        status = 'Sem alteracao'

                    resultados.append({
                        'numero_nc': nc.numero,
                        'cod_om': codom_upload_norm,
                        'diex': diex_upload_num,
                        'pi': pi_upload_norm,
                        'nd': nd_upload_norm,
                        'cod_ug': cod_ug_upload_norm,
                        'valor': valor_upload_norm,
                        'nc_siafi': nc.nc_siafi or '',
                        'ref_sisnc': nc.ref_sisnc or '',
                        'status': status
                    })
            else:
                candidatos_sem_codom = indice_sem_codom.get((cod_ug_upload_norm, pi_upload_norm, nd_upload_norm, valor_upload_norm), [])
                if candidatos_sem_codom:
                    diagnostico['match_sem_codom'] += 1
                    for nc_cand in candidatos_sem_codom[:3]:
                        chave_qm = (str(nc_upload), nc_cand.numero, 'COD OM')
                        if chave_qm in quase_match_ids or len(quase_matches) >= MAX_QUASE_MATCHES:
                            continue
                        quase_match_ids.add(chave_qm)
                        dados_sys = extrair_chave_sistema_nc(nc_cand)
                        quase_matches.append({
                            'fonte_upload': fonte_upload,
                            'campo_divergente': 'COD OM',
                            'nc_upload': nc_upload,
                            'numero_nc_sistema': dados_sys['numero_nc'],
                            'cod_om_upload': codom_upload_norm,
                            'cod_om_sistema': dados_sys['cod_om'],
                            'cod_ug_upload': cod_ug_upload_norm,
                            'cod_ug_sistema': dados_sys['cod_ug'],
                            'pi_upload': pi_upload_norm,
                            'pi_sistema': dados_sys['pi'],
                            'nd_upload': nd_upload_norm,
                            'nd_sistema': dados_sys['nd'],
                            'valor_upload': valor_upload_norm,
                            'valor_sistema': dados_sys['valor'],
                            'diex_upload': diex_upload_num,
                            'diex_sistema': dados_sys['diex']
                        })

                candidatos_sem_valor = indice_sem_valor.get((codom_upload_norm, cod_ug_upload_norm, pi_upload_norm, nd_upload_norm), [])
                if candidatos_sem_valor:
                    diagnostico['match_sem_valor'] += 1
                    for nc_cand in candidatos_sem_valor[:3]:
                        chave_qm = (str(nc_upload), nc_cand.numero, 'VALOR')
                        if chave_qm in quase_match_ids or len(quase_matches) >= MAX_QUASE_MATCHES:
                            continue
                        quase_match_ids.add(chave_qm)
                        dados_sys = extrair_chave_sistema_nc(nc_cand)
                        quase_matches.append({
                            'fonte_upload': fonte_upload,
                            'campo_divergente': 'VALOR',
                            'nc_upload': nc_upload,
                            'numero_nc_sistema': dados_sys['numero_nc'],
                            'cod_om_upload': codom_upload_norm,
                            'cod_om_sistema': dados_sys['cod_om'],
                            'cod_ug_upload': cod_ug_upload_norm,
                            'cod_ug_sistema': dados_sys['cod_ug'],
                            'pi_upload': pi_upload_norm,
                            'pi_sistema': dados_sys['pi'],
                            'nd_upload': nd_upload_norm,
                            'nd_sistema': dados_sys['nd'],
                            'valor_upload': valor_upload_norm,
                            'valor_sistema': dados_sys['valor'],
                            'diex_upload': diex_upload_num,
                            'diex_sistema': dados_sys['diex']
                        })

                candidatos_sem_nd = indice_sem_nd.get((codom_upload_norm, cod_ug_upload_norm, pi_upload_norm, valor_upload_norm), [])
                if candidatos_sem_nd:
                    diagnostico['match_sem_nd'] += 1
                    for nc_cand in candidatos_sem_nd[:3]:
                        chave_qm = (str(nc_upload), nc_cand.numero, 'ND')
                        if chave_qm in quase_match_ids or len(quase_matches) >= MAX_QUASE_MATCHES:
                            continue
                        quase_match_ids.add(chave_qm)
                        dados_sys = extrair_chave_sistema_nc(nc_cand)
                        quase_matches.append({
                            'fonte_upload': fonte_upload,
                            'campo_divergente': 'ND',
                            'nc_upload': nc_upload,
                            'numero_nc_sistema': dados_sys['numero_nc'],
                            'cod_om_upload': codom_upload_norm,
                            'cod_om_sistema': dados_sys['cod_om'],
                            'cod_ug_upload': cod_ug_upload_norm,
                            'cod_ug_sistema': dados_sys['cod_ug'],
                            'pi_upload': pi_upload_norm,
                            'pi_sistema': dados_sys['pi'],
                            'nd_upload': nd_upload_norm,
                            'nd_sistema': dados_sys['nd'],
                            'valor_upload': valor_upload_norm,
                            'valor_sistema': dados_sys['valor'],
                            'diex_upload': diex_upload_num,
                            'diex_sistema': dados_sys['diex']
                        })

                candidatos_sem_pi = indice_sem_pi.get((codom_upload_norm, cod_ug_upload_norm, nd_upload_norm, valor_upload_norm), [])
                if candidatos_sem_pi:
                    diagnostico['match_sem_pi'] += 1
                    for nc_cand in candidatos_sem_pi[:3]:
                        chave_qm = (str(nc_upload), nc_cand.numero, 'PI')
                        if chave_qm in quase_match_ids or len(quase_matches) >= MAX_QUASE_MATCHES:
                            continue
                        quase_match_ids.add(chave_qm)
                        dados_sys = extrair_chave_sistema_nc(nc_cand)
                        quase_matches.append({
                            'fonte_upload': fonte_upload,
                            'campo_divergente': 'PI',
                            'nc_upload': nc_upload,
                            'numero_nc_sistema': dados_sys['numero_nc'],
                            'cod_om_upload': codom_upload_norm,
                            'cod_om_sistema': dados_sys['cod_om'],
                            'cod_ug_upload': cod_ug_upload_norm,
                            'cod_ug_sistema': dados_sys['cod_ug'],
                            'pi_upload': pi_upload_norm,
                            'pi_sistema': dados_sys['pi'],
                            'nd_upload': nd_upload_norm,
                            'nd_sistema': dados_sys['nd'],
                            'valor_upload': valor_upload_norm,
                            'valor_sistema': dados_sys['valor'],
                            'diex_upload': diex_upload_num,
                            'diex_sistema': dados_sys['diex']
                        })

                candidatos_sem_cod_ug = indice_sem_cod_ug.get((codom_upload_norm, pi_upload_norm, nd_upload_norm, valor_upload_norm), [])
                if candidatos_sem_cod_ug:
                    diagnostico['match_sem_cod_ug'] += 1
                    for nc_cand in candidatos_sem_cod_ug[:3]:
                        chave_qm = (str(nc_upload), nc_cand.numero, 'COD UG')
                        if chave_qm in quase_match_ids or len(quase_matches) >= MAX_QUASE_MATCHES:
                            continue
                        quase_match_ids.add(chave_qm)
                        dados_sys = extrair_chave_sistema_nc(nc_cand)
                        quase_matches.append({
                            'fonte_upload': fonte_upload,
                            'campo_divergente': 'COD UG',
                            'nc_upload': nc_upload,
                            'numero_nc_sistema': dados_sys['numero_nc'],
                            'cod_om_upload': codom_upload_norm,
                            'cod_om_sistema': dados_sys['cod_om'],
                            'cod_ug_upload': cod_ug_upload_norm,
                            'cod_ug_sistema': dados_sys['cod_ug'],
                            'pi_upload': pi_upload_norm,
                            'pi_sistema': dados_sys['pi'],
                            'nd_upload': nd_upload_norm,
                            'nd_sistema': dados_sys['nd'],
                            'valor_upload': valor_upload_norm,
                            'valor_sistema': dados_sys['valor'],
                            'diex_upload': diex_upload_num,
                            'diex_sistema': dados_sys['diex']
                        })

                nao_encontradas.append({
                    'nc_upload': nc_upload,
                    'cod_om': codom_upload_norm,
                    'diex': diex_upload_num,
                    'pi': pi_upload_norm,
                    'nd': nd_upload_norm,
                    'cod_ug': cod_ug_upload_norm,
                    'valor': valor_upload_norm
                })

            if codom_upload_norm and pi_upload_norm and nd_upload_norm and cod_ug_upload_norm and valor_upload_norm != '0.00':
                diagnostico['keys_upload'] += 1
                if diagnostico['sample_upload_key'] is None:
                    diagnostico['sample_upload_key'] = chave_upload

        mapa_sistema_reduzido = {}
        for nc in ncs_sistema:
            chave_reduzida = (
                normalizar_cod_ug(nc.cod_ug),
                normalizar_pi(nc.pi),
                normalizar_nd(nc.nd),
                valor_para_chave(nc.valor)
            )
            cod_ug_sys, pi_sys, nd_sys, valor_sys = chave_reduzida
            if not cod_ug_sys or not pi_sys or not nd_sys or valor_sys == '0.00':
                continue
            mapa_sistema_reduzido.setdefault(chave_reduzida, []).append(nc)

        try:
            df_tg = _carregar_upload_auditoria_cache('extra')
        except Exception:
            df_tg = pd.DataFrame()

        try:
            df_sisnc = _carregar_upload_auditoria_sisnc_cache('extra')
        except Exception:
            df_sisnc = pd.DataFrame()

        mapa_tg_reduzido = construir_mapa_reduzido_upload(df_tg)
        mapa_sisnc_reduzido = construir_mapa_reduzido_upload(df_sisnc)
        # --- NOVA LÓGICA DE COMPARAÇÃO SISNC x TESOURO GERENCIAL POR NC SIAFI ---
        # Extrair pares NC SIAFI (6 dígitos) -> dados normalizados
        def extrair_ultimos_6_digitos(valor):
            val = str(valor or '').strip()
            digitos = ''.join(filter(str.isdigit, val))
            return digitos[-6:] if len(digitos) >= 6 else digitos

        # SISNC: coluna C (NC SIAFI), TG: coluna A (NC)
        sisnc_ncs = {}
        def _extrair_nc_siafi_de_linha(row, col_idx=2):
            # Primeiro tenta extrair pelos últimos 6 dígitos da coluna esperada
            nc = extrair_ultimos_6_digitos(row.get(col_idx, ''))
            if nc and len(str(nc)) >= 6:
                return nc
            # Se não houver, procura em todas as células por bloco de 6 dígitos
            for v in row.values:
                try:
                    texto = str(v or '')
                except Exception:
                    texto = ''
                m = re.search(r'(?<!\d)(\d{6})(?!\d)', texto)
                if m:
                    return m.group(1)
            # retorna o que encontrou (pode ser vazio ou curto)
            return nc

        for _, row in df_sisnc.iterrows():
            nc_siafi = _extrair_nc_siafi_de_linha(row, col_idx=2)  # coluna C
            if not nc_siafi:
                continue
            cod_ug = normalizar_cod_ug(row.get(1, ''))
            doc = str(row.get(6, ''))
            codom = extrair_codom_numerico(doc)
            pi = normalizar_pi(row.get(5, ''))
            nd = normalizar_nd(row.get(3, row.get(4, '')))
            # identifica valor na coluna esperada ou faz fallback por toda a linha
            valor_f, valor_raw = extrair_valor_linha(row, 7)
            valor = valor_f
            sisnc_ncs[nc_siafi] = {
                'cod_ug': cod_ug,
                'codom': codom,
                'pi': pi,
                'nd': nd,
                'valor': valor,
                'valor_raw': valor_raw,
                'row': row
            }

        tg_ncs = {}
        for _, row in df_tg.iterrows():
            nc_tg = _extrair_nc_siafi_de_linha(row, col_idx=0)  # coluna A
            if not nc_tg:
                continue
            cod_ug = normalizar_cod_ug(row.get(1, ''))
            doc = str(row.get(6, ''))
            codom = extrair_codom_numerico(doc)
            pi = normalizar_pi(row.get(5, ''))
            nd = normalizar_nd(row.get(3, row.get(4, '')))
            valor_f, valor_raw = extrair_valor_linha(row, 7)
            valor = valor_f
            tg_ncs[nc_tg] = {
                'cod_ug': cod_ug,
                'codom': codom,
                'pi': pi,
                'nd': nd,
                'valor': valor,
                'valor_raw': valor_raw,
                'row': row
            }

        inconsistencias_tg_sisnc = []
        def _codug_equivalente(codug_sisnc, codom_sisnc, codug_tg):
            """Retorna True se codug_tg corresponder a codug_sisnc ou ao codom extraído do SISNC."""
            if not codug_tg:
                return False
            if codug_sisnc and codug_sisnc == codug_tg:
                return True
            if codom_sisnc and codom_sisnc == codug_tg:
                return True
            return False

        for nc_siafi, dados_sisnc in sisnc_ncs.items():
            if nc_siafi in tg_ncs:
                dados_tg = tg_ncs[nc_siafi]
                inconsistencias = []
                # comparar CODUG usando fallback: cod_ug do SISNC OU codom extraído
                if not _codug_equivalente(dados_sisnc.get('cod_ug'), dados_sisnc.get('codom'), dados_tg.get('cod_ug')):
                    inconsistencias.append('CODUG')
                if dados_sisnc.get('codom') != dados_tg.get('codom'):
                    inconsistencias.append('CODOM')
                if dados_sisnc.get('pi') != dados_tg.get('pi'):
                    inconsistencias.append('PI')
                if dados_sisnc.get('nd') != dados_tg.get('nd'):
                    inconsistencias.append('ND')
                try:
                    v1 = normalizar_valor(dados_sisnc['valor'])
                    v2 = normalizar_valor(dados_tg['valor'])
                    if abs(v1 - v2) > 0.01:
                        inconsistencias.append('VALOR')
                except Exception:
                    inconsistencias.append('VALOR')
                if inconsistencias:
                    inconsistencias_tg_sisnc.append({
                        'nc_tg': nc_siafi,
                        'nc_sisnc': nc_siafi,
                        'cod_ug': dados_tg.get('cod_ug') or dados_sisnc.get('cod_ug') or '--',
                        'pi': dados_tg.get('pi') or dados_sisnc.get('pi') or '--',
                        'nd': dados_tg.get('nd') or dados_sisnc.get('nd') or '--',
                        'referencia_sisnc': (dados_sisnc.get('row')[2] if hasattr(dados_sisnc.get('row'), '__getitem__') and 2 in dados_sisnc.get('row') else '--') if dados_sisnc.get('row') is not None else '--',
                        'valor_tg': dados_tg.get('valor', 0.0),
                        'valor_sisnc': dados_sisnc.get('valor', 0.0),
                        'detalhe': f'Inconsistência(s) em: {", ".join(inconsistencias)}'
                    })
            else:
                inconsistencias_tg_sisnc.append({
                    'nc_tg': None,
                    'nc_sisnc': nc_siafi,
                    'cod_ug': dados_sisnc.get('cod_ug') or '--',
                    'pi': dados_sisnc.get('pi') or '--',
                    'nd': dados_sisnc.get('nd') or '--',
                    'referencia_sisnc': (dados_sisnc.get('row')[2] if hasattr(dados_sisnc.get('row'), '__getitem__') and 2 in dados_sisnc.get('row') else '--') if dados_sisnc.get('row') is not None else '--',
                    'valor_tg': None,
                    'valor_sisnc': dados_sisnc.get('valor', 0.0),
                    'detalhe': 'NC SIAFI presente no SISNC e ausente no Tesouro Gerencial'
                })
        for nc_tg in tg_ncs:
            if nc_tg not in sisnc_ncs:
                dados_tg = tg_ncs.get(nc_tg, {})
                inconsistencias_tg_sisnc.append({
                    'nc_tg': nc_tg,
                    'nc_sisnc': None,
                    'cod_ug': dados_tg.get('cod_ug') or '--',
                    'pi': dados_tg.get('pi') or '--',
                    'nd': dados_tg.get('nd') or '--',
                    'referencia_sisnc': '--',
                    'valor_tg': dados_tg.get('valor', 0.0),
                    'valor_sisnc': None,
                    'detalhe': 'NC presente no Tesouro Gerencial e ausente no SISNC'
                })
        # --- FIM NOVA LÓGICA ---
        # mapa_tg_sem_valor = construir_mapa_tg_sisnc_sem_valor(df_tg, fonte_nome='Tesouro Gerencial')
        # mapa_sisnc_sem_valor = construir_mapa_tg_sisnc_sem_valor(df_sisnc, fonte_nome='SISNC')

        resumo_tres_fontes['total_sistema'] = len(mapa_sistema_reduzido)
        resumo_tres_fontes['total_tg'] = len(mapa_tg_reduzido)
        resumo_tres_fontes['total_sisnc'] = len(mapa_sisnc_reduzido)

        divergencia_tg_sisnc_keys = set()

        for chave, ncs_chave in mapa_sistema_reduzido.items():
            presente_tg = chave in mapa_tg_reduzido
            presente_sisnc = chave in mapa_sisnc_reduzido

            if presente_tg and presente_sisnc:
                resumo_tres_fontes['sistema_ok_nas_tres'] += len(ncs_chave)
                continue

            if not presente_tg:
                resumo_tres_fontes['sistema_sem_tg'] += len(ncs_chave)
            if not presente_sisnc:
                resumo_tres_fontes['sistema_sem_sisnc'] += len(ncs_chave)

            # Neste relatório específico, exibe apenas inconsistências Sistema x SISNC.
            if not presente_sisnc:
                for nc in ncs_chave:
                    inconsistencias_tres_fontes.append({
                        'origem': 'Sistema',
                        'numero_nc': nc.numero,
                        'cod_ug': chave[0],
                        'pi': chave[1],
                        'nd': chave[2],
                        'valor': chave[3],
                        'presente_sistema': True,
                        'presente_sisnc': False,
                        'detalhe': 'Registro presente no sistema e ausente no SISNC'
                    })

            if presente_tg != presente_sisnc:
                divergencia_tg_sisnc_keys.add(chave)

        for chave, info in mapa_sisnc_reduzido.items():
            if chave in mapa_sistema_reduzido:
                continue
            resumo_tres_fontes['sisnc_sem_sistema'] += 1
            if chave not in mapa_tg_reduzido:
                divergencia_tg_sisnc_keys.add(chave)

            # --- INÍCIO DA AUTOMAÇÃO DE SUGESTÃO DE GERAÇÃO DE NC ---
            # Buscar solicitações compatíveis percorrendo pedidos e itens
            cod_ug, pi, nd, valor = chave
            solicitacoes = SolicitacaoExtraPDRLOG.query.all()
            solicitacoes_compat = []
            for s in solicitacoes:
                for pedido in getattr(s, 'pedidos', []):
                    cod_ug_pedido = normalizar_cod_ug(getattr(pedido, 'codug', ''))
                    for item in getattr(pedido, 'itens', []):
                        pi_item = normalizar_pi(getattr(item, 'pi', ''))
                        nd_item = normalizar_nd(getattr(item, 'nd', ''))
                        valor_item = valor_para_chave(getattr(item, 'valor_aprovado', 0))
                        if (
                            cod_ug_pedido == cod_ug and
                            pi_item == pi and
                            nd_item == nd and
                            abs(float(valor_item.replace(',', '.')) - float(valor.replace(',', '.'))) < 0.01
                        ):
                            solicitacoes_compat.append((s, pedido, item))

            sugestao_automatica = None
            if len(solicitacoes_compat) == 1:
                solicit, pedido, item = solicitacoes_compat[0]
                nc_existente = NotaCredito.query.filter_by(solicitacao_pdrlog_id=solicit.id, pi=pi, nd=nd, valor=float(valor.replace(',', '.'))).first()
                if not nc_existente:
                    ref_sisnc = str(info.get('referencia_sisnc', ''))
                    nc_siafi = str(info.get('nc_upload', ''))
                    status_auto = _status_automatico_nc(ref_sisnc=ref_sisnc, nc_siafi=nc_siafi, status_atual='Pendente')
                    from datetime import datetime
                    base_numero = f"NC-AUTO-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                    numero_nc = base_numero
                    sufixo = 1
                    while NotaCredito.query.filter_by(numero=numero_nc).first():
                        numero_nc = f"{base_numero}-{sufixo}"
                        sufixo += 1
                    nova_nc = NotaCredito(
                        numero=numero_nc,
                        cod_ug=cod_ug,
                        pi=pi,
                        nd=nd,
                        valor=float(valor.replace(',', '.')),
                        ref_sisnc=ref_sisnc,
                        nc_siafi=nc_siafi,
                        status=status_auto,
                        usuario_id=1,  # Ajustar para o usuário correto
                        solicitacao_pdrlog_id=solicit.id,
                        data_criacao=datetime.now(),
                        data_atualizacao=datetime.now()
                    )
                    db.session.add(nova_nc)
                    db.session.commit()
                    sugestao_automatica = f"NC gerada automaticamente para solicitação {solicit.id} (PI: {pi}, ND: {nd}, Valor: {valor}) com ref_sisnc e nc_siafi preenchidos."
                else:
                    sugestao_automatica = f"Solicitação {solicit.id} já possui NC vinculada."
            elif len(solicitacoes_compat) > 1:
                diex_sisnc = str(info.get('referencia_sisnc', ''))
                solicit_diex = [t for t in solicitacoes_compat if getattr(t[0], 'diex', '') == diex_sisnc]
                if len(solicit_diex) == 1:
                    solicit, pedido, item = solicit_diex[0]
                    nc_existente = NotaCredito.query.filter_by(solicitacao_pdrlog_id=solicit.id, pi=pi, nd=nd, valor=float(valor.replace(',', '.'))).first()
                    if not nc_existente:
                        ref_sisnc = str(info.get('referencia_sisnc', ''))
                        nc_siafi = str(info.get('nc_upload', ''))
                        status_auto = _status_automatico_nc(ref_sisnc=ref_sisnc, nc_siafi=nc_siafi, status_atual='Pendente')
                        from datetime import datetime
                        base_numero = f"NC-AUTO-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                        numero_nc = base_numero
                        sufixo = 1
                        while NotaCredito.query.filter_by(numero=numero_nc).first():
                            numero_nc = f"{base_numero}-{sufixo}"
                            sufixo += 1
                        nova_nc = NotaCredito(
                            numero=numero_nc,
                            cod_ug=cod_ug,
                            pi=pi,
                            nd=nd,
                            valor=float(valor.replace(',', '.')),
                            ref_sisnc=ref_sisnc,
                            nc_siafi=nc_siafi,
                            status=status_auto,
                            usuario_id=1,  # Ajustar para o usuário correto
                            solicitacao_pdrlog_id=solicit.id,
                            data_criacao=datetime.now(),
                            data_atualizacao=datetime.now()
                        )
                        db.session.add(nova_nc)
                        db.session.commit()
                        sugestao_automatica = f"NC gerada automaticamente para solicitação {solicit.id} (PI: {pi}, ND: {nd}, Valor: {valor}) com ref_sisnc e nc_siafi preenchidos (match DIEx)."
                    else:
                        sugestao_automatica = f"Solicitação {solicit.id} já possui NC vinculada (match DIEx)."
                else:
                    sugestao_automatica = f"Mais de uma solicitação compatível encontrada para PI: {pi}, ND: {nd}, Valor: {valor}. Não foi possível gerar NC automaticamente."
            else:
                sugestao_automatica = f"Nenhuma solicitação compatível encontrada para PI: {pi}, ND: {nd}, Valor: {valor}."

            inconsistencias_tres_fontes.append({
                'origem': 'SISNC',
                'numero_nc': info.get('nc_upload', ''),
                'cod_ug': chave[0],
                'pi': chave[1],
                'nd': chave[2],
                'valor': chave[3],
                'presente_sistema': False,
                'presente_sisnc': True,
                'detalhe': 'Registro presente no SISNC e ausente no sistema',
                'sugestao_automatica': sugestao_automatica
            })

        resumo_tres_fontes['divergencia_tg_sisnc'] = len(divergencia_tg_sisnc_keys)

        # (Removido: lógica antiga de comparação TG x SISNC por chaves compostas)

        if len(inconsistencias_tres_fontes) > 800:
            inconsistencias_tres_fontes = inconsistencias_tres_fontes[:800]

        if pd.isna(total_upload_valor) or math.isinf(total_upload_valor):
            total_upload_valor = 0.0
        total_upload_valor = round(total_upload_valor, 2)
        upload_col_total = diagnostico.get('upload_col_total') or 0.0
        if pd.isna(upload_col_total) or math.isinf(upload_col_total):
            upload_col_total = 0.0
        if total_upload_valor == 0.0 and upload_col_total > 0:
            total_upload_valor = upload_col_total
        if pd.isna(total_sistema_valor) or math.isinf(total_sistema_valor):
            total_sistema_valor = 0.0
        if total_sistema_relatorio is None and total_sistema_valor == 0.0:
            soma_sql = diagnostico.get('sistema_valores_sum_sql') or 0.0
            if soma_sql > 0:
                total_sistema_valor = round(float(soma_sql), 2)
                diagnostico['total_sistema_valor'] = total_sistema_valor
        total_diferenca = round(total_upload_valor - total_sistema_valor, 2)
        diagnostico['total_upload_valor'] = total_upload_valor
        diagnostico['total_quase_matches'] = len(quase_matches)

        try:
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao salvar NC SIAFI: {str(e)}', 'error')
            return redirect(url_for('auditoria'))

        flash(
            f'Auditoria concluida. Correspondencias: {total_matches}, '
            f'atualizadas: {total_atualizadas}, sem alteracao: {total_sem_alteracao}, '
            f'sem NC SIAFI no upload: {total_sem_nc_siafi}, '
            f'preenchidas automaticamente: {total_preenchidas_auto}.',
            'success'
        )
        flash(
            'Diagnostico: '
            f"NCs sistema={diagnostico['ncs_sistema_total']}, "
            f"PI col={diagnostico['pi_col_idx']}, "
            f"Valor col={diagnostico['valor_col_idx']}, "
            f"Keys sistema={diagnostico['keys_sistema']}, "
            f"Keys upload={diagnostico['keys_upload']}, "
            f"Sem CODOM upload={diagnostico['sem_codom']}, "
            f"Sem DIEx num={diagnostico['sem_diex']}, "
            f"Match sem CODOM={diagnostico['match_sem_codom']}, "
            f"Exemplo sistema={diagnostico['sample_sys_key']}, "
            f"Exemplo upload={diagnostico['sample_upload_key']}",
            'info'
        )

        if request.form.get('exportar_excel') == '1':
            duplicidades_export = []
            for item in duplicidades_sistema:
                chave = item.get('chave', ('', '', '', '', ''))
                for nc in item.get('ncs', []) or []:
                    duplicidades_export.append({
                        'cod_om': chave[0] if len(chave) > 0 else '',
                        'cod_ug': chave[1] if len(chave) > 1 else '',
                        'pi': chave[2] if len(chave) > 2 else '',
                        'nd': chave[3] if len(chave) > 3 else '',
                        'valor_chave': chave[4] if len(chave) > 4 else '',
                        'numero_nc': getattr(nc, 'numero', ''),
                        'valor_nc': getattr(nc, 'valor', 0)
                    })

            tabela_export = str(request.form.get('exportar_tabela', '') or '').strip().lower()
            mapa_tabelas_export = {
                'resumo': (
                    'Resumo',
                    [{
                        'correspondencias': total_matches,
                        'atualizadas': total_atualizadas,
                        'sem_alteracao': total_sem_alteracao,
                        'sem_nc_siafi': total_sem_nc_siafi,
                        'preenchidas_automaticamente': total_preenchidas_auto,
                        'total_upload_valor': total_upload_valor,
                        'total_sistema_valor': total_sistema_valor,
                        'total_diferenca': total_diferenca,
                        'duplicidades_sistema': total_duplicidades_sistema
                    }]
                ),
                'consolidacao_3_fontes': ('Consolidacao_3_Fontes', [resumo_tres_fontes]),
                'resultados': ('Resultados', resultados),
                'nao_encontradas': ('Nao_Encontradas', nao_encontradas),
                'quase_matches': ('Quase_Matches', quase_matches),
                'duplicidades_sistema': ('Duplicidades_Sistema', duplicidades_export),
                'inc_tg_sisnc': ('Inc_TG_SISNC', inconsistencias_tg_sisnc),
                'inc_3_fontes': ('Inc_3_Fontes', inconsistencias_tres_fontes)
            }

            selecionada = mapa_tabelas_export.get(tabela_export)
            if not selecionada:
                flash('Selecione uma tabela válida para exportar.', 'error')
            else:
                nome_planilha, dados_planilha = selecionada
                return _exportar_planilhas_excel('auditoria_extra_pdr_log', [
                    (nome_planilha, dados_planilha)
                ])

    return render_template(
        'auditoria.html',
        resultados=resultados,
        nao_encontradas=nao_encontradas,
        quase_matches=quase_matches,
        total_matches=total_matches,
        total_atualizadas=total_atualizadas,
        total_sem_alteracao=total_sem_alteracao,
        total_sem_nc_siafi=total_sem_nc_siafi,
        total_preenchidas_auto=total_preenchidas_auto,
        total_duplicidades_sistema=total_duplicidades_sistema,
        duplicidades_sistema=duplicidades_sistema,
        inconsistencias_tres_fontes=inconsistencias_tres_fontes,
        inconsistencias_tg_sisnc=inconsistencias_tg_sisnc,
        resumo_tres_fontes=resumo_tres_fontes,
        diagnostico=diagnostico,
        auditoria_executada=(request.method == 'POST'),
        total_upload_valor=total_upload_valor,
        total_sistema_valor=total_sistema_valor,
        total_diferenca=total_diferenca
    )


@app.route('/pdr-log/auditoria', methods=['GET', 'POST'])
@acesso_requerido('admin', 'usuario')
def auditoria_pdr_log():
    """Auditoria do PDR Log por agrupamento OM + ND + PI + valor, com preenchimento de NC SIAFI."""
    PIS_PDR_PERMITIDOS = {'E6SUPLJA5PA', 'E6SUPLJA7PA', 'E6SUPLJA6OP', 'E6SUPLJA8OP'}
    definir_modulo_menu(MENU_MODULO_PDR)

    def texto_sem_acentos(valor):
        if valor is None:
            return ''
        texto = str(valor)
        texto = unicodedata.normalize('NFD', texto)
        texto = ''.join(ch for ch in texto if unicodedata.category(ch) != 'Mn')
        return texto

    def normalizar_texto_chave(valor):
        texto = texto_sem_acentos(valor).upper().strip()
        texto = re.sub(r'\s+', ' ', texto)
        return texto

    def normalizar_om_chave(valor):
        texto = normalizar_texto_chave(valor)
        return re.sub(r'[^A-Z0-9]', '', texto)

    def om_corresponde_doc(om_base_texto, doc_texto):
        om_base = normalizar_om_chave(om_base_texto)
        doc = normalizar_om_chave(doc_texto)
        if not om_base or not doc:
            return False
        if om_base in doc:
            return True

        om_tokens = [tok for tok in re.split(r'\s+', normalizar_texto_chave(om_base_texto)) if len(tok) >= 2]
        if om_tokens:
            tokens_norm = [re.sub(r'[^A-Z0-9]', '', t) for t in om_tokens]
            tokens_norm = [t for t in tokens_norm if t]
            if tokens_norm and all(tok in doc for tok in tokens_norm):
                return True

        return False

    def assunto_para_nd_pi(nome_assunto):
        assunto = normalizar_texto_chave(nome_assunto)
        nd = ''
        pi = ''

        if 'PASA' in assunto:
            pi = 'E6SUPLJA5PA'
        elif ('MNT INSTALACOES ST APROV' in assunto) or (('MNT' in assunto or 'MANUTEN' in assunto) and 'INSTAL' in assunto and 'ST APROV' in assunto):
            pi = 'E6SUPLJA7PA'
        elif ('MNT INSTALACOES OP' in assunto) or (('MNT' in assunto or 'MANUTEN' in assunto) and 'INSTAL' in assunto and 'OP' in assunto):
            pi = 'E6SUPLJA8OP'
        elif 'MNT OP' in assunto or (('MNT' in assunto or 'MANUTEN' in assunto) and 'OP' in assunto):
            pi = 'E6SUPLJA6OP'

        return nd, pi

    def pis_equivalentes(pi_base):
        # No PDR Log a correspondência de PI é estrita, sem variantes do fluxo Extra.
        return {normalizar_pi(pi_base)}

    def encontrar_coluna(df, candidatos):
        normalizados = {}
        for col in df.columns:
            chave = re.sub(r'[^a-z0-9]', '', texto_sem_acentos(col).lower())
            normalizados[chave] = col
        for cand in candidatos:
            chave = re.sub(r'[^a-z0-9]', '', texto_sem_acentos(cand).lower())
            if chave in normalizados:
                return normalizados[chave]
        return None

    def extrair_primeiro_nc_siafi(texto):
        valor = str(texto or '').strip()
        if not valor:
            return ''

        # Regra: usar o último bloco numérico com exatamente 6 dígitos
        blocos = re.findall(r'(?<!\d)\d{6}(?!\d)', valor)
        if not blocos:
            return ''
        return blocos[-1]

    def valor_upload_seguro(valor):
        valor_norm = normalizar_valor(valor)
        try:
            valor_float = float(valor_norm)
        except Exception:
            return 0.0
        if pd.isna(valor_float) or math.isinf(valor_float):
            return 0.0
        return valor_float

    def extrair_valor_upload_linha(linha, col_valor_ref):
        """Obtém valor monetário da linha com fallback para outras colunas."""
        valor_base = valor_upload_seguro(linha.get(col_valor_ref, 0))
        if valor_base > 0:
            return valor_base

        for celula in linha.values:
            texto = str(celula or '').strip()
            if not texto:
                continue

            if re.search(r'\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2}|\d+\.\d{2}', texto):
                candidato = valor_upload_seguro(texto)
                if 0 < candidato <= 1_000_000_000:
                    return candidato

        return 0.0

    def escolher_colunas_por_conteudo(df):
        colunas = list(df.columns)

        def score_nd(col):
            try:
                serie = df[col]
                return int(sum(1 for v in serie if re.fullmatch(r'(33|44)\d{4}', normalizar_nd(v) or '')))
            except Exception:
                return 0

        def score_pi(col):
            try:
                serie = df[col].astype(str)
                return int(serie.str.contains(r'\bE\d[A-Z0-9]{6,}\b', case=False, na=False).sum())
            except Exception:
                return 0

        def score_doc(col):
            try:
                serie = df[col].astype(str)
                score_obs = int(serie.str.contains(r'OBS|DIEX|C SUP|SOL', case=False, na=False).sum())
                score_len = int((serie.str.len() > 25).sum())
                return score_obs + score_len
            except Exception:
                return 0

        def score_valor(col):
            try:
                serie = df[col]
                vals = serie.apply(normalizar_valor)
                positivos = int((vals > 0).sum())

                textos = serie.astype(str)
                com_decimal = int(textos.str.contains(r'\d+[\.,]\d{1,2}$', regex=True, na=False).sum())
                longos_ids = int(textos.str.contains(r'\b\d{12,}\b', regex=True, na=False).sum())

                # Penaliza colunas com cara de identificador (ex.: número da NC muito longo)
                score = (positivos * 2) + (com_decimal * 4) - (longos_ids * 5)

                # Bônus para faixa monetária plausível do domínio (evita capturar IDs gigantes)
                faixa_plausivel = int(((vals >= 0.01) & (vals <= 100000000)).sum())
                score += faixa_plausivel

                return score
            except Exception:
                return 0

        melhor_nd = max(colunas, key=score_nd) if colunas else None
        melhor_pi = max(colunas, key=score_pi) if colunas else None
        melhor_doc = max(colunas, key=score_doc) if colunas else None
        melhor_valor = max(colunas, key=score_valor) if colunas else None

        return melhor_nd, melhor_pi, melhor_doc, melhor_valor

    resumo = {
        'total_base': 0,
        'total_upload': 0,
        'total_correspondentes': 0,
        'total_divergentes': 0,
        'total_somente_upload': 0,
        'total_somente_base': 0,
        'total_nc_siafi_preenchidas': 0,
        'total_quase_matches_tres_fontes': 0,
        'tg_sisnc_somente_tg': 0,
        'tg_sisnc_somente_sisnc': 0,
        'tg_sisnc_valor_divergente': 0,
        'tg_sisnc_valor_divergente_com_recolhimento': 0,
        'tg_sisnc_valor_divergente_operacional': 0,
        'total_inconsistencias_tg_sisnc': 0
    }
    correspondentes = []
    divergentes = []
    somente_upload = []
    somente_base = []
    quase_matches_tres_fontes = []
    inconsistencias_tg_sisnc = []

    try:
        base_lista_bruta = _carregar_solicitacoes_pdr_log_planilha()
        base_lista, resumo_dedup_nc = _deduplicar_solicitacoes_para_nc_pdr(base_lista_bruta)
        base_grupos_nc = _agrupar_notas_credito_pdr(solicitacoes=base_lista, incluir_solicitacoes=True, somente_status_sim=True)
    except Exception as e:
        flash(f'Erro ao carregar planilha base do PDR Log: {e}', 'error')
        return render_template(
            'auditoria_pdr_log.html',
            resumo=resumo,
            correspondentes=correspondentes,
            divergentes=divergentes,
            somente_upload=somente_upload,
            somente_base=somente_base,
            quase_matches_tres_fontes=quase_matches_tres_fontes,
            inconsistencias_tg_sisnc=inconsistencias_tg_sisnc,
            auditoria_executada=False
        )

    if resumo_dedup_nc.get('duplicadas_removidas', 0) > 0:
        flash(
            'Base da auditoria PDR Log consolidada sem duplicidades da lista de NC: '
            f"{resumo_dedup_nc.get('duplicadas_removidas', 0)} solicitação(ões) removida(s) antes do comparativo TG/SISNC.",
            'info'
        )

    base_agrupada = {}
    base_por_numero = {}
    for grupo in base_grupos_nc:
        nd_mapeada = normalizar_nd(grupo.get('nd', ''))
        pi_mapeada = normalizar_pi(grupo.get('pi', ''))
        om_texto = grupo.get('om_solicitante', '')
        om_chave = normalizar_om_chave(om_texto)
        if not om_chave or not nd_mapeada or not pi_mapeada:
            continue

        chave = (om_chave, nd_mapeada, pi_mapeada)
        base_agrupada[chave] = {
            'om': om_texto,
            'nd': nd_mapeada,
            'pi': pi_mapeada,
            'valor_total': valor_upload_seguro(grupo.get('valor_total', 0)),
            'solicitacoes': grupo.get('solicitacoes', []),
            'codug_set': {
                normalizar_cod_ug(s.get('codug', ''))
                for s in grupo.get('solicitacoes', [])
                if normalizar_cod_ug(s.get('codug', ''))
            },
            'codom_set': {
                re.sub(r'\D', '', str(s.get('codom', '') or ''))
                for s in grupo.get('solicitacoes', [])
                if re.sub(r'\D', '', str(s.get('codom', '') or ''))
            }
        }

        for solicitacao in grupo.get('solicitacoes', []):
            numero_norm = _normalizar_numero_solicitacao_pdr(solicitacao.get('numero', ''))
            if not numero_norm:
                continue
            base_por_numero[numero_norm] = {
                'om_solicitante': solicitacao.get('om_solicitante', ''),
                'codug': solicitacao.get('codug', '')
            }

    resumo['total_base'] = len(base_agrupada)

    if request.method == 'POST':
        arquivo = request.files.get('arquivo')
        excel = None
        aba_bd = None
        if arquivo and arquivo.filename:
            try:
                excel = pd.ExcelFile(arquivo)
                aba_bd = next((aba for aba in excel.sheet_names if str(aba).strip().lower() == 'bd'), excel.sheet_names[0])
                df_upload = pd.read_excel(excel, sheet_name=aba_bd)
                df_upload.columns = [str(col).strip() for col in df_upload.columns]
            except Exception as e:
                flash(f'Erro ao ler arquivo de conferência: {e}', 'error')
                return redirect(url_for('auditoria_pdr_log'))
        else:
            try:
                df_upload = _carregar_upload_auditoria_cache_combinado('pdr')
            except FileNotFoundError:
                flash('Bases de auditoria não encontradas (TG/SISNC). Envie as planilhas no Painel Subsistência.', 'error')
                return redirect(url_for('painel_subsistencia'))
            except Exception as e:
                flash(f'Erro ao carregar bases da auditoria (TG/SISNC): {str(e)}', 'error')
                return redirect(url_for('auditoria_pdr_log'))

        df_upload = _normalizar_colunas_numericas(df_upload)

        col_nd = encontrar_coluna(df_upload, ['Natureza Despesa Código', 'Natureza Despesa Codigo', 'ND'])
        col_pi = encontrar_coluna(df_upload, ['PI'])
        col_doc = encontrar_coluna(df_upload, ['Doc - Observação', 'Doc - Observacao', 'Doc Observacao', 'Observacao'])
        col_valor = encontrar_coluna(df_upload, ['Valor', 'Valor Total', 'VALOR_TOTAL'])
        col_codug = encontrar_coluna(df_upload, ['Favorecido Doc. Número', 'Favorecido Doc Numero', 'CODUG', 'COD UG'])
        col_nc_siafi = encontrar_coluna(df_upload, ['NC SIAFI', 'NC_SIAFI', 'NCSIAFI', 'NC Favorecido Doc. Número', 'NC Favorecido Doc Numero'])

        if col_nd is None or col_pi is None or col_doc is None or col_valor is None:
            if excel is not None and aba_bd is not None:
                try:
                    df_upload_alt = pd.read_excel(excel, sheet_name=aba_bd, header=None, skiprows=5)
                    df_upload_alt = _normalizar_colunas_numericas(df_upload_alt)
                    col_nd_auto, col_pi_auto, col_doc_auto, col_valor_auto = escolher_colunas_por_conteudo(df_upload_alt)
                    if col_nd_auto is not None and col_pi_auto is not None and col_doc_auto is not None and col_valor_auto is not None:
                        df_upload = df_upload_alt
                        col_nd = col_nd_auto
                        col_pi = col_pi_auto
                        col_doc = col_doc_auto
                        col_valor = col_valor_auto
                        if 1 in df_upload.columns:
                            col_codug = 1
                        if 0 in df_upload.columns:
                            col_nc_siafi = 0
                except Exception:
                    pass

        if col_nd is None or col_pi is None or col_doc is None or col_valor is None:
            col_nd_auto, col_pi_auto, col_doc_auto, col_valor_auto = escolher_colunas_por_conteudo(df_upload)
            if col_nd_auto is not None and col_pi_auto is not None and col_doc_auto is not None and col_valor_auto is not None:
                col_nd = col_nd_auto
                col_pi = col_pi_auto
                col_doc = col_doc_auto
                col_valor = col_valor_auto
                if 1 in df_upload.columns:
                    col_codug = 1
                if 0 in df_upload.columns:
                    col_nc_siafi = 0

        if col_nd is None or col_pi is None or col_doc is None or col_valor is None:
            flash('Planilha de auditoria sem colunas necessárias (ND, PI, Doc - Observação, Valor). Verifique o layout e tente novamente.', 'error')
            return redirect(url_for('auditoria_pdr_log'))

        upload_agrupada = {}
        upload_agrupada_por_fonte = {}
        for _, linha in df_upload.iterrows():
            nd_chave = normalizar_nd(linha.get(col_nd, ''))
            doc_texto = _normalizar_texto_planilha(linha.get(col_doc, ''))
            pi_chave = normalizar_pi(linha.get(col_pi, ''))
            if pi_chave not in PIS_PDR_PERMITIDOS:
                pi_mapeado_doc = normalizar_pi(_mapear_pi_pdr_por_assunto(doc_texto))
                if pi_mapeado_doc in PIS_PDR_PERMITIDOS:
                    pi_chave = pi_mapeado_doc
            doc_chave = normalizar_om_chave(doc_texto)
            codug_chave = normalizar_cod_ug(linha.get(col_codug, '')) if col_codug is not None else ''
            valor = extrair_valor_upload_linha(linha, col_valor)
            fonte_upload = str(linha.get('__fonte_upload', 'Upload manual') or 'Upload manual').strip()
            referencia_sisnc = str(linha.get(2, '') or '').strip()

            if not nd_chave or not pi_chave or not doc_chave:
                continue
            if pi_chave not in PIS_PDR_PERMITIDOS:
                continue

            chave_upload = (doc_chave, nd_chave, pi_chave, codug_chave)
            registro_upload = upload_agrupada.setdefault(chave_upload, {
                'doc': doc_texto,
                'nd': nd_chave,
                'pi': pi_chave,
                'codug': codug_chave,
                'valor_total': 0.0,
                'nc_siafi': '',
                'fontes': set()
            })
            registro_upload['valor_total'] += valor
            registro_upload['fontes'].add(fonte_upload)

            chave_upload_fonte = (fonte_upload, doc_chave, nd_chave, pi_chave, codug_chave)
            registro_upload_fonte = upload_agrupada_por_fonte.setdefault(chave_upload_fonte, {
                'fonte_upload': fonte_upload,
                'doc': doc_texto,
                'nd': nd_chave,
                'pi': pi_chave,
                'codug': codug_chave,
                'valor_total': 0.0,
                'referencia_sisnc': ''
            })
            registro_upload_fonte['valor_total'] += valor
            if not registro_upload_fonte.get('referencia_sisnc') and referencia_sisnc:
                registro_upload_fonte['referencia_sisnc'] = referencia_sisnc

            if col_nc_siafi is not None:
                nc_candidato = extrair_primeiro_nc_siafi(linha.get(col_nc_siafi, ''))
                if nc_candidato and not registro_upload['nc_siafi']:
                    registro_upload['nc_siafi'] = nc_candidato

        resumo['total_upload'] = len(upload_agrupada)

        mapa_nc_siafi = _carregar_cache_nc_pdr()

        base_keys_matched = set()
        upload_total_por_base = {}
        nc_siafi_por_base = {}
        diagnostico_match = {
            'upload_grupos': len(upload_agrupada),
            'base_grupos': len(base_agrupada),
            'match_nd_pi': 0,
            'match_om': 0,
            'match_valor': 0,
            'nc_conflitos_limpados': 0
        }

        for upload_key, upload_item in upload_agrupada.items():
            _, nd_upload, pi_upload, codug_upload = upload_key

            candidatos_nd_pi = []
            for base_key, base_item in base_agrupada.items():
                _, nd_base, pi_base = base_key
                if nd_base != nd_upload:
                    continue
                if pi_upload not in pis_equivalentes(pi_base):
                    continue
                candidatos_nd_pi.append((base_key, base_item))

            if candidatos_nd_pi:
                diagnostico_match['match_nd_pi'] += 1

            candidatos = []
            for base_key, base_item in candidatos_nd_pi:
                if codug_upload:
                    if codug_upload in base_item.get('codug_set', set()):
                        candidatos.append((base_key, base_item))
                else:
                    if om_corresponde_doc(base_item.get('om', ''), upload_item.get('doc', '')):
                        candidatos.append((base_key, base_item))

            if candidatos:
                diagnostico_match['match_om'] += 1

            if not candidatos:
                somente_upload.append({
                    'numero': '--',
                    'nc_siafi': upload_item.get('nc_siafi', ''),
                    'valor_total': round(valor_upload_seguro(upload_item.get('valor_total', 0)), 2),
                    'status': f"ND {nd_upload} / PI {pi_upload} / OM não identificada em Doc"
                })
                continue

            if len(candidatos) > 1:
                candidatos.sort(key=lambda c: abs(
                    round(valor_upload_seguro(c[1].get('valor_total', 0)), 2)
                    - round(valor_upload_seguro(upload_item.get('valor_total', 0)), 2)
                ))

            base_key, base_item = candidatos[0]
            base_keys_matched.add(base_key)

            upload_valor = round(valor_upload_seguro(upload_item.get('valor_total', 0)), 2)
            upload_total_por_base[base_key] = round(upload_total_por_base.get(base_key, 0.0) + upload_valor, 2)

            nc_siafi_upload = upload_item.get('nc_siafi', '')
            if nc_siafi_upload and base_key not in nc_siafi_por_base:
                nc_siafi_por_base[base_key] = nc_siafi_upload
            if nc_siafi_upload:
                numeros_alvo = {
                    _normalizar_numero_solicitacao_pdr(s.get('numero', ''))
                    for s in base_item.get('solicitacoes', [])
                    if _normalizar_numero_solicitacao_pdr(s.get('numero', ''))
                }
                diagnostico_match['nc_conflitos_limpados'] += _limpar_conflitos_nc_siafi_pdr(
                    mapa_nc=mapa_nc_siafi,
                    nc_siafi_alvo=nc_siafi_upload,
                    numeros_alvo=numeros_alvo,
                    base_por_numero=base_por_numero,
                    codug_set_alvo=base_item.get('codug_set', set()),
                    om_chave_alvo=normalizar_om_chave(base_item.get('om', ''))
                )

                for solicitacao in base_item.get('solicitacoes', []):
                    numero_solic = _normalizar_numero_solicitacao_pdr(solicitacao.get('numero', ''))
                    if not numero_solic:
                        continue
                    chave_cache = _chave_cache_nc_pdr(
                        codom=solicitacao.get('codom', ''),
                        codug=solicitacao.get('codug', ''),
                        pi=normalizar_pi(_mapear_pi_pdr_por_assunto(solicitacao.get('nome_assunto', ''))),
                        nd=normalizar_nd(solicitacao.get('cod_nd', '')),
                        valor=solicitacao.get('valor_total', 0)
                    )
                    nc_atual = re.sub(r'\D', '', str(mapa_nc_siafi.get(chave_cache, '') or ''))
                    if nc_atual != nc_siafi_upload:
                        mapa_nc_siafi[chave_cache] = nc_siafi_upload
                        if numero_solic in mapa_nc_siafi:
                            del mapa_nc_siafi[numero_solic]
                        resumo['total_nc_siafi_preenchidas'] += 1

        for base_key, item in base_agrupada.items():
            numeros_solicitacao = [s.get('numero', '') for s in item.get('solicitacoes', []) if s.get('numero')]
            codom_set = sorted({str(s.get('codom', '') or '').strip() for s in item.get('solicitacoes', []) if str(s.get('codom', '') or '').strip()})
            codug_set = sorted({str(s.get('codug', '') or '').strip() for s in item.get('solicitacoes', []) if str(s.get('codug', '') or '').strip()})
            numeros_texto = ', '.join(numeros_solicitacao) if numeros_solicitacao else '--'
            base_valor = round(valor_upload_seguro(item.get('valor_total', 0)), 2)

            if base_key in base_keys_matched:
                upload_valor_consolidado = round(upload_total_por_base.get(base_key, 0.0), 2)
                if base_valor != upload_valor_consolidado:
                    divergentes.append({
                        'om_solicitante': item.get('om', '--'),
                        'codom': ', '.join(codom_set) if codom_set else '--',
                        'codug': ', '.join(codug_set) if codug_set else '--',
                        'numero': numeros_texto,
                        'campos': 'Valor Total consolidado (base x upload)',
                        'valor_base': base_valor,
                        'valor_upload': upload_valor_consolidado,
                        'status_base': 'gerar NC',
                        'status_upload': '--'
                    })
                else:
                    diagnostico_match['match_valor'] += 1
                    tem_nc = bool(nc_siafi_por_base.get(base_key, ''))
                    correspondentes.append({
                        'numero': numeros_texto,
                        'valor': base_valor,
                        'status': 'NC SIAFI gerada' if tem_nc else 'gerar NC'
                    })
                continue

            somente_base.append({
                'om_solicitante': item.get('om', '--'),
                'codom': ', '.join(codom_set) if codom_set else '--',
                'codug': ', '.join(codug_set) if codug_set else '--',
                'numero': numeros_texto,
                'valor_total': base_valor,
                'status': f"ND {item.get('nd', '--')} / PI {item.get('pi', '--')}"
            })

        # Quase match entre 3 bases no PDR Log (Sistema x TG x SISNC).
        # Critério: 2 de 3 campos em comum (ND, PI e OM/CODUG), com 1 campo divergente.
        MAX_QUASE_MATCHES = 300
        candidatos_quase = {}

        for base_key, base_item in base_agrupada.items():
            om_base = base_item.get('om', '')
            nd_base = base_item.get('nd', '')
            pi_base = base_item.get('pi', '')
            valor_base = round(valor_upload_seguro(base_item.get('valor_total', 0)), 2)

            numeros_solicitacao = [s.get('numero', '') for s in base_item.get('solicitacoes', []) if s.get('numero')]
            numeros_texto = ', '.join(numeros_solicitacao) if numeros_solicitacao else '--'

            for upload_key_fonte, upload_item_fonte in upload_agrupada_por_fonte.items():
                fonte_upload, _, nd_upload, pi_upload, codug_upload = upload_key_fonte

                match_nd = (nd_base == nd_upload)
                match_pi = (pi_upload in pis_equivalentes(pi_base))
                match_om = False
                if codug_upload:
                    match_om = codug_upload in base_item.get('codug_set', set())
                if not match_om:
                    match_om = om_corresponde_doc(om_base, upload_item_fonte.get('doc', ''))

                score = int(match_nd) + int(match_pi) + int(match_om)
                valor_upload = round(valor_upload_seguro(upload_item_fonte.get('valor_total', 0)), 2)
                delta_valor = round(abs(valor_upload - valor_base), 2)

                # Quase match por chave: 2 de 3 campos batendo.
                # Quase match por valor: 3 de 3 campos batendo e apenas valor divergente.
                if score == 3:
                    if delta_valor == 0:
                        continue
                    campo_divergente = 'VALOR'
                elif score == 2:
                    if not match_nd:
                        campo_divergente = 'ND'
                    elif not match_pi:
                        campo_divergente = 'PI'
                    else:
                        campo_divergente = 'OM/CODUG'
                else:
                    continue

                chave_candidato = (base_key, fonte_upload, campo_divergente)
                candidato_atual = candidatos_quase.get(chave_candidato)

                registro_candidato = {
                    'fonte_upload': fonte_upload,
                    'numero_base': numeros_texto,
                    'om_base': om_base or '--',
                    'nd_base': nd_base or '--',
                    'pi_base': pi_base or '--',
                    'valor_base': valor_base,
                    'doc_upload': upload_item_fonte.get('doc', '--') or '--',
                    'codug_upload': codug_upload or '--',
                    'nd_upload': nd_upload or '--',
                    'pi_upload': pi_upload or '--',
                    'valor_upload': valor_upload,
                    'campo_divergente': campo_divergente,
                    'diferenca_valor': round(valor_upload - valor_base, 2),
                    '_delta_abs': delta_valor
                }

                if candidato_atual is None or delta_valor < candidato_atual.get('_delta_abs', float('inf')):
                    candidatos_quase[chave_candidato] = registro_candidato

        quase_matches_tres_fontes = sorted(
            candidatos_quase.values(),
            key=lambda item: item.get('_delta_abs', 0)
        )[:MAX_QUASE_MATCHES]

        for item in quase_matches_tres_fontes:
            if '_delta_abs' in item:
                del item['_delta_abs']

        resumo['total_quase_matches_tres_fontes'] = len(quase_matches_tres_fontes)

        def classificar_fonte_tg_sisnc(nome_fonte):
            nome = normalizar_texto_chave(nome_fonte)
            if 'SISNC' in nome:
                return 'SISNC'
            if 'TESOURO' in nome or nome == 'TG':
                return 'TG'
            return ''

        def limpar_referencia_sisnc(valor):
            texto = str(valor or '').strip()
            if texto.lower() in ('nan', 'none', 'null', '--'):
                return ''
            return texto

        def eh_recolhimento_sisnc(doc_texto):
            """Identifica lançamentos SISNC que devem abater o saldo descentralizado."""
            doc = normalizar_texto_chave(doc_texto)
            if not doc:
                return False

            if 'RECOLHIMENTO' not in doc:
                return False

            return (
                'DUPLICIDADE' in doc
                or 'CREDITO' in doc
                or 'CREDITO ' in doc
                or 'RECOLHIMENTO DE CREDITO' in doc
                or 'RECOLHIMENTO POR DUPLICIDADE' in doc
            )

        def _extrair_valor_tg_sisnc_por_fonte(linha, fonte_upload):
            # TG e SISNC usam layouts de colunas diferentes nos caches CSV.
            # Primeiro tenta colunas esperadas por fonte; depois usa fallback genérico.
            candidatos = []
            if fonte_upload == 'Tesouro Gerencial':
                candidatos = [8, 7, col_valor]
            elif fonte_upload == 'SISNC':
                candidatos = [7, 8, col_valor]
            else:
                candidatos = [col_valor, 7, 8]

            vistos = set()
            for col in candidatos:
                if col is None or col in vistos:
                    continue
                vistos.add(col)
                valor = valor_upload_seguro(linha.get(col, 0))
                if valor > 0:
                    return round(valor, 2)

            return round(valor_upload_seguro(extrair_valor_upload_linha(linha, col_valor)), 2)

        def extrair_registro_tg_sisnc_linha(linha, fonte_upload):
            nd = normalizar_nd(linha.get(col_nd, ''))
            pi = normalizar_pi(linha.get(col_pi, ''))
            codug = normalizar_cod_ug(linha.get(col_codug, '')) if col_codug is not None else ''
            doc = _normalizar_texto_planilha(linha.get(col_doc, ''))
            codom = extrair_codom_numerico(doc)
            valor_float = _extrair_valor_tg_sisnc_por_fonte(linha, fonte_upload)
            diex = extrair_numero_diex(doc)
            referencia = limpar_referencia_sisnc(linha.get(2, ''))
            eh_recolhimento = (fonte_upload == 'SISNC') and eh_recolhimento_sisnc(doc)
            sinal = -1.0 if eh_recolhimento else 1.0
            valor_assinado = round(valor_float * sinal, 2)
            return {
                'doc': doc,
                'nd': nd,
                'pi': pi,
                'codug': codug,
                'codom': codom,
                'diex': diex,
                'valor': valor_assinado,
                'valor_original': valor_float,
                'eh_recolhimento': eh_recolhimento,
                'referencia_sisnc': referencia
            }

        def agregar_item_tg_sisnc(mapa_destino, chave_operacional, item):
            registro = mapa_destino.setdefault(chave_operacional, {
                'nd': item.get('nd', ''),
                'pi': item.get('pi', ''),
                'codug': item.get('codug', ''),
                'codom': item.get('codom', ''),
                'valor_total': 0.0,
                'docs': [],
                'diex_set': set(),
                'referencia_sisnc': '',
                'qtd_recolhimentos': 0,
                'qtd_lancamentos': 0
            })

            registro['valor_total'] += float(item.get('valor', 0.0) or 0.0)
            registro['qtd_lancamentos'] += 1

            doc = item.get('doc', '')
            if doc and len(registro['docs']) < 5:
                registro['docs'].append(doc)

            diex = item.get('diex', '')
            if diex:
                registro['diex_set'].add(diex)

            if item.get('eh_recolhimento'):
                registro['qtd_recolhimentos'] += 1

            ref = item.get('referencia_sisnc', '')
            if ref and not registro.get('referencia_sisnc'):
                registro['referencia_sisnc'] = ref

        mapa_tg_upload = {}
        mapa_sisnc_upload = {}

        for _, linha in df_upload.iterrows():
            fonte_upload = str(linha.get('__fonte_upload', 'Upload manual') or 'Upload manual').strip()
            classe_fonte = classificar_fonte_tg_sisnc(fonte_upload)
            if classe_fonte not in ('TG', 'SISNC'):
                continue

            item = extrair_registro_tg_sisnc_linha(linha, fonte_upload)
            if not item['nd'] or not item['pi'] or not item['codug'] or not item['codom']:
                continue
            if item['valor'] <= 0:
                continue

            chave_base = (item['nd'], item['pi'], item['codug'], item['codom'])
            alvo = mapa_tg_upload if classe_fonte == 'TG' else mapa_sisnc_upload
            agregar_item_tg_sisnc(alvo, chave_base, item)

        chaves_tg_sisnc = set(mapa_tg_upload.keys()) | set(mapa_sisnc_upload.keys())
        for chave in chaves_tg_sisnc:
            item_tg = mapa_tg_upload.get(chave)
            item_sisnc = mapa_sisnc_upload.get(chave)

            valor_tg = round(float(item_tg.get('valor_total', 0.0)) if item_tg else 0.0, 2)
            valor_sisnc = round(float(item_sisnc.get('valor_total', 0.0)) if item_sisnc else 0.0, 2)

            if item_tg and item_sisnc and valor_tg == valor_sisnc:
                continue

            if item_tg and not item_sisnc:
                resumo['tg_sisnc_somente_tg'] += 1
                detalhe = 'Registro presente apenas no Tesouro Gerencial'
                classificacao_divergencia = 'Sem contraparte SISNC'
            elif item_sisnc and not item_tg:
                resumo['tg_sisnc_somente_sisnc'] += 1
                detalhe = 'Registro presente apenas no SISNC'
                classificacao_divergencia = 'Sem contraparte TG'
            else:
                resumo['tg_sisnc_valor_divergente'] += 1
                if int(item_sisnc.get('qtd_recolhimentos', 0) or 0) > 0:
                    resumo['tg_sisnc_valor_divergente_com_recolhimento'] += 1
                    classificacao_divergencia = 'Divergencia com recolhimento'
                else:
                    resumo['tg_sisnc_valor_divergente_operacional'] += 1
                    classificacao_divergencia = 'Divergencia operacional'
                detalhe = 'Valor divergente entre Tesouro Gerencial e SISNC'

            base_info = item_tg or item_sisnc or {}
            docs = base_info.get('docs', [])
            diex_set = sorted(list(base_info.get('diex_set', set())))

            inconsistencias_tg_sisnc.append({
                'doc_upload': (docs[0] if docs else '--') or '--',
                'codug_upload': base_info.get('codug', '') or '--',
                'codom_upload': base_info.get('codom', '') or '--',
                'nd': base_info.get('nd', '') or '--',
                'pi': base_info.get('pi', '') or '--',
                'diex': ', '.join(diex_set) if diex_set else '--',
                'referencia_sisnc': item_sisnc.get('referencia_sisnc', '') if item_sisnc else '',
                'valor_tg': valor_tg,
                'valor_sisnc': valor_sisnc,
                'classificacao_divergencia': classificacao_divergencia,
                'detalhe': detalhe
            })

        total_inconsistencias_tg_sisnc = len(inconsistencias_tg_sisnc)
        if len(inconsistencias_tg_sisnc) > 800:
            inconsistencias_tg_sisnc = inconsistencias_tg_sisnc[:800]
        resumo['total_inconsistencias_tg_sisnc'] = total_inconsistencias_tg_sisnc

        resumo['total_correspondentes'] = len(correspondentes)
        resumo['total_divergentes'] = len(divergentes)
        resumo['total_somente_upload'] = len(somente_upload)
        resumo['total_somente_base'] = len(somente_base)

        _salvar_cache_nc_pdr(mapa_nc_siafi)
        PDR_LOG_CACHE['mtime'] = None
        PDR_LOG_CACHE['data'] = []

        flash(
            f"Auditoria do PDR Log concluída com sucesso. NC SIAFI preenchidas/atualizadas: {resumo['total_nc_siafi_preenchidas']}.",
            'success'
        )
        flash(
            f"Diagnóstico de correspondência: base={diagnostico_match['base_grupos']}, upload={diagnostico_match['upload_grupos']}, ND/PI={diagnostico_match['match_nd_pi']}, OM={diagnostico_match['match_om']}, valor={diagnostico_match['match_valor']}, conflitos NC limpos={diagnostico_match['nc_conflitos_limpados']}.",
            'info'
        )

        if request.form.get('exportar_excel') == '1':
            tabela_export = str(request.form.get('exportar_tabela', '') or '').strip().lower()
            mapa_tabelas_export = {
                'resumo': ('Resumo', [resumo]),
                'correspondentes': ('Correspondentes', correspondentes),
                'divergentes': ('Divergentes', divergentes),
                'somente_upload': ('Somente_Upload', somente_upload),
                'somente_base': ('Somente_Base', somente_base),
                'quase_match_3_fontes': ('Quase_Match_3_Fontes', quase_matches_tres_fontes),
                'inc_tg_sisnc': ('Inc_TG_SISNC', inconsistencias_tg_sisnc)
            }

            selecionada = mapa_tabelas_export.get(tabela_export)
            if not selecionada:
                flash('Selecione uma tabela válida para exportar.', 'error')
            else:
                nome_planilha, dados_planilha = selecionada
                return _exportar_planilhas_excel('auditoria_pdr_log', [
                    (nome_planilha, dados_planilha)
                ])

    return render_template(
        'auditoria_pdr_log.html',
        resumo=resumo,
        correspondentes=correspondentes,
        divergentes=divergentes,
        somente_upload=somente_upload,
        somente_base=somente_base,
        quase_matches_tres_fontes=quase_matches_tres_fontes,
        inconsistencias_tg_sisnc=inconsistencias_tg_sisnc,
        auditoria_executada=(request.method == 'POST')
    )


@app.route('/__auditoria_debug')
def auditoria_debug():
    """Endpoint simples para confirmar a instancia que esta servindo o app."""
    return make_response(
        f"DEBUG auditoria OK | {os.path.abspath(__file__)} | cwd={os.getcwd()}",
        200
    )


@app.route('/__app_debug')
def app_debug():
    """Mostra o caminho do app e se moeda_para_float existe."""
    has_moeda = 'moeda_para_float' in globals()
    return make_response(
        f"DEBUG app OK | {os.path.abspath(__file__)} | cwd={os.getcwd()} | moeda_para_float={has_moeda}",
        200
    )

@app.route('/nc/<int:id>')
@login_required
def detalhes_nc(id):
    """Detalhes de uma Nota de Crédito específica - todos os usuários"""
    try:
        # CORREÇÃO: Carregar os relacionamentos com os nomes corretos
        nc = NotaCredito.query\
    .options(
        db.joinedload(NotaCredito.usuario),
        db.joinedload(NotaCredito.solicitacao_origem),  # CORRIGIDO
        db.joinedload(NotaCredito.pro_origem)  # CORRIGIDO
    )\
    .filter(NotaCredito.id == id)\
    .first()
        
        if not nc:
            flash('Nota de Crédito não encontrada!', 'error')
            return redirect(url_for('listar_ncs'))
            
        return render_template('detalhes_nc.html',
                            nc=nc,
                            status_list=STATUS_NC)
    except Exception as e:
        flash(f'Erro ao carregar Nota de Crédito: {str(e)}', 'error')
        return redirect(url_for('listar_ncs'))

@app.route('/nc/<int:id>/editar', methods=['POST'])
@login_required
def editar_nc(id):
    """Editar uma Nota de Crédito com controle automático de status - todos os usuários"""
    nc = db.session.get(NotaCredito, id)
    if not nc:
        flash('Nota de Crédito não encontrada!', 'error')
        return redirect(url_for('listar_ncs'))

    try:
        # Salvar status anterior para verificar mudanças
        status_anterior = nc.status
        
        # Obter valores dos campos

        ref_sisnc = request.form.get('ref_sisnc', '').strip()
        nc_siafi = request.form.get('nc_siafi', '').strip()
        status_selecionado = request.form.get('status')
        descricao = request.form.get('descricao', '').strip()
        conferida = request.form.get('conferida') == 'true'

        # Permitir edição de todos os campos enquanto status for 'Não Conferida'
        if nc.status == 'Não Conferida':
            nc.numero = request.form.get('numero', nc.numero)
            nc.cod_ug = request.form.get('cod_ug', nc.cod_ug)
            nc.sigla_ug = request.form.get('sigla_ug', nc.sigla_ug)
            nc.pi = request.form.get('pi', nc.pi)
            nc.nd = request.form.get('nd', nc.nd)
            valor_str = request.form.get('valor', '').replace('.', '').replace(',', '.')
            try:
                nc.valor = float(valor_str)
            except Exception:
                pass

        # SEMPRE permitir edição da descrição
        nc.descricao = descricao
        nc.ref_sisnc = ref_sisnc
        nc.nc_siafi = nc_siafi
        
        print(f"📝 Descrição recebida: {descricao[:100]}...")
        
        # LÓGICA AUTOMÁTICA DE STATUS - FLUXO CORRIGIDO
        if status_selecionado == 'Cancelada':
            # Usuario escolheu cancelar manualmente
            nc.status = 'Cancelada'
            print(f"🔄 Status definido como 'Cancelada' por seleção do usuário")
        
        elif conferida:
            # Se marcou como conferida
            if nc.status == 'Não Conferida':
                nc.status = 'Pendente'  # Não Conferida → Pendente
                print(f"🔄 Status definido como 'Pendente' (após conferência)")
            else:
                flash('Apenas NCs com status "Não Conferida" podem ser conferidas.', 'warning')
        
        elif nc_siafi and nc_siafi.strip():
            # Se tem NC SIAFI preenchido -> Processada SIAFI
            nc.status = 'Processada SIAFI'
            print(f"🔄 Status automático: NC SIAFI preenchido -> Processada SIAFI")
        
        elif ref_sisnc and ref_sisnc.strip():
            # Se tem REF SISNC preenchido -> Cadastrada SISNC
            nc.status = 'Cadastrada SISNC'
            print(f"🔄 Status automático: REF SISNC preenchido -> Cadastrada SISNC")
        
        # NOTA: Não voltamos para "Pendente" automaticamente, apenas através da conferência
        
        # Permitir NC SIAFI mesmo sem REF SISNC
        
        # Validar: não pode preencher REF SISNC ou NC SIAFI se ainda não foi conferida
        if (ref_sisnc or nc_siafi) and nc.status == 'Não Conferida':
            flash('Atenção: Para preencher REF SISNC ou NC SIAFI, a NC deve primeiro ser conferida (status "Pendente").', 'warning')
            nc.status = 'Pendente'  # Força para Pendente se tentar preencher referências

        if status_anterior != 'Cancelada' and nc.status == 'Cancelada' and nc.item_id:
            item = db.session.get(ItemPedido, nc.item_id)
            if item:
                saldo_item = item.valor_restante if item.valor_restante is not None else (item.valor_aprovado or 0)
                item.valor_restante = saldo_item + (nc.valor or 0)
        elif status_anterior == 'Cancelada' and nc.status != 'Cancelada' and nc.item_id:
            item = db.session.get(ItemPedido, nc.item_id)
            if item:
                saldo_item = item.valor_restante if item.valor_restante is not None else (item.valor_aprovado or 0)
                item.valor_restante = max(saldo_item - (nc.valor or 0), 0)
        
        db.session.commit()
        
        # Log de mudança de status
        if status_anterior != nc.status:
            print(f"✅ Status atualizado: {status_anterior} -> {nc.status}")
            flash(f'Nota de Crédito atualizada! Status alterado para: {nc.status}', 'success')
        else:
            flash('Nota de Crédito atualizada com sucesso!', 'success')
    
    except Exception as e:
        db.session.rollback()
        print(f"❌ Erro ao atualizar Nota de Crédito: {str(e)}")
        import traceback
        traceback.print_exc()
        flash(f'Erro ao atualizar Nota de Crédito: {str(e)}', 'error')

    return redirect(url_for('detalhes_nc', id=id))


@app.route('/nc/<int:id>/excluir', methods=['POST'], endpoint='excluir_nc')
@acesso_requerido('admin')
def excluir_nc(id):
    """Exclui definitivamente uma Nota de Crédito"""
    nc = db.session.get(NotaCredito, id)
    if not nc:
        flash('Nota de Crédito não encontrada!', 'error')
        return redirect(url_for('listar_ncs'))
    try:
        pro = None
        # Se a NC está vinculada a uma PRO, atualizar saldo e status da PRO
        if nc.pro_id:
            pro = db.session.get(Pro, nc.pro_id)
            if pro:
                if nc.item_id:
                    item = db.session.get(ItemPedido, nc.item_id)
                    if item:
                        saldo_item = item.valor_restante if item.valor_restante is not None else (item.valor_aprovado or 0)
                        item.valor_restante = saldo_item + (nc.valor or 0)
                # Restaurar saldo da PRO
                if pro.valor_restante is None:
                    pro.valor_restante = pro.valor_total
                pro.valor_restante += nc.valor or 0
                # Atualizar status da PRO
                ncs_ativas = NotaCredito.query.filter_by(pro_id=pro.id).filter(NotaCredito.id != nc.id).filter(NotaCredito.status != 'Cancelada').all()
                if not ncs_ativas:
                    pro.status = 'Aguardando término do processo licitatório'
                else:
                    pro.status = 'Parcialmente Convertida' if pro.valor_restante > 0 else 'Convertida em NC'
                pro.data_atualizacao = datetime.now()
        db.session.delete(nc)
        db.session.commit()
        flash('Nota de Crédito excluída com sucesso.', 'success')
        # Se era vinculada a PRO, redirecionar para detalhes da PRO
        if pro:
            return redirect(url_for('detalhes_pro', id=pro.id))
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao excluir Nota de Crédito: {str(e)}', 'error')
    return redirect(url_for('listar_ncs'))
 

@app.route('/nc/<int:id>/reativar', methods=['POST'])
@login_required
def reativar_nc(id):
    """Reativa uma Nota de Crédito cancelada - todos os usuários"""
    nc = db.session.get(NotaCredito, id)
    if not nc:
        flash('Nota de Crédito não encontrada!', 'error')
        return redirect(url_for('listar_ncs'))
    
    try:
        # Verificar se a NC está cancelada
        if nc.status != 'Cancelada':
            flash('Apenas Notas de Crédito canceladas podem ser reativadas!', 'error')
            return redirect(url_for('detalhes_nc', id=id))
        
        # Obter dados do formulário
        novo_status = request.form.get('status', '').strip()
        motivo_reativacao = request.form.get('motivo_reativacao', '').strip()
        
        if not novo_status:
            flash('Erro: É necessário selecionar o novo status para a reativação!', 'error')
            return redirect(url_for('detalhes_nc', id=id))
        
        if not motivo_reativacao:
            flash('Erro: É necessário informar o motivo da reativação!', 'error')
            return redirect(url_for('detalhes_nc', id=id))
        
        # Validar status selecionado
        status_validos = ['Pendente', 'Cadastrada SISNC', 'Processada SIAFI']
        if novo_status not in status_validos:
            flash(f'Erro: Status inválido para reativação. Use: {", ".join(status_validos)}', 'error')
            return redirect(url_for('detalhes_nc', id=id))
        
        # Validar se pode usar status Cadastrada SISNC ou Processada SIAFI
        if novo_status == 'Cadastrada SISNC' and not nc.ref_sisnc:
            flash('Erro: Não é possível reativar como "Cadastrada SISNC" sem ter REF SISNC anterior!', 'error')
            return redirect(url_for('detalhes_nc', id=id))
        
        if novo_status == 'Processada SIAFI' and not nc.nc_siafi:
            flash('Erro: Não é possível reativar como "Processada SIAFI" sem ter NC SIAFI anterior!', 'error')
            return redirect(url_for('detalhes_nc', id=id))
        
        # Salvar histórico do cancelamento
        motivo_cancelamento_anterior = nc.motivo_cancelamento or ''
        
        # Reativar a NC
        nc.status = novo_status
        nc.motivo_cancelamento = f"{motivo_cancelamento_anterior}\n\n--- REATIVADA ---\nData: {datetime.now().strftime('%d/%m/%Y %H:%M')}\nPor: {current_user.nome}\nMotivo da reativação: {motivo_reativacao}\nNovo status: {novo_status}"
        nc.data_atualizacao = datetime.now()

        if nc.item_id:
            item = db.session.get(ItemPedido, nc.item_id)
            if item:
                saldo_item = item.valor_restante if item.valor_restante is not None else (item.valor_aprovado or 0)
                item.valor_restante = max(saldo_item - (nc.valor or 0), 0)
        
        db.session.commit()
        
        print(f"✅ Nota de Crédito {nc.numero} reativada de 'Cancelada' para '{novo_status}'")
        flash(f'Nota de Crédito {nc.numero} reativada com sucesso! Status alterado para: {novo_status}', 'success')
        
    except Exception as e:
        db.session.rollback()
        print(f"❌ Erro ao reativar Nota de Crédito: {str(e)}")
        import traceback
        traceback.print_exc()
        flash(f'Erro ao reativar Nota de Crédito: {str(e)}', 'error')
    
    return redirect(url_for('detalhes_nc', id=id))



# ===== ROTAS PARA PRO (APENAS ADMIN E USUARIO) =====

@app.route('/pros')
@acesso_requerido('admin', 'usuario')
def listar_pros():
    """Lista PROs com filtros básicos."""
    try:
        status_filter = request.args.get('status', '')
        numero_filter = request.args.get('numero', '')
        data_inicio_filter = request.args.get('data_inicio', '')
        data_fim_filter = request.args.get('data_fim', '')
        cod_ug_filter = request.args.get('cod_ug', '')
        sigla_ug_filter = request.args.get('sigla_ug', '')
        import unicodedata

        def normalizar(txt: str) -> str:
            if not txt:
                return ''
            # Substitui caracteres de substituição por 'e' e remove acentos
            txt = str(txt).replace('\ufffd', 'e')
            txt = unicodedata.normalize('NFD', txt)
            txt = ''.join(ch for ch in txt if unicodedata.category(ch) != 'Mn')
            return txt.lower().strip()

        pros_all = Pro.query.order_by(Pro.data_criacao.desc()).all()
        pros = []

        # Pré-processar datas
        data_inicio = None
        data_fim = None
        if data_inicio_filter:
            try:
                data_inicio = datetime.strptime(data_inicio_filter, '%Y-%m-%d')
            except ValueError:
                data_inicio = None
        if data_fim_filter:
            try:
                data_fim = datetime.strptime(data_fim_filter, '%Y-%m-%d')
            except ValueError:
                data_fim = None

        for pro in pros_all:
            # Marcar como vencida se aplicável
            if pro.vencida and pro.status != 'Vencida (mais de 120 dias)':
                pro.marcar_vencida()
                db.session.commit()

            if status_filter:
                alvo = normalizar(status_filter)
                fonte = normalizar(pro.status)
                if fonte != alvo:
                    continue
            if numero_filter and numero_filter.lower() not in (pro.numero or '').lower():
                continue
            if cod_ug_filter and cod_ug_filter not in str(pro.cod_ug or ''):
                continue
            if sigla_ug_filter and sigla_ug_filter.lower() not in (pro.sigla_ug or '').lower():
                continue
            if data_inicio and pro.data_criacao and pro.data_criacao < data_inicio:
                continue
            if data_fim and pro.data_criacao and pro.data_criacao > data_fim:
                continue

            # Normalizar status para exibição
            if pro.status in STATUS_PRO_VARIANTES:
                pro.status = STATUS_PRO_VARIANTES[pro.status]
            pros.append(pro)

        return render_template('pros.html',
                               pros=pros,
                               status_filter=status_filter,
                               numero_filter=numero_filter,
                               data_inicio_filter=data_inicio_filter,
                               data_fim_filter=data_fim_filter,
                               cod_ug_filter=cod_ug_filter,
                               sigla_ug_filter=sigla_ug_filter,
                               status_list=STATUS_PRO)
    except Exception as e:
        print(f"❌ Erro ao carregar PROs: {str(e)}")
        import traceback
        traceback.print_exc()
        flash(f'Erro ao carregar PROs: {str(e)}', 'error')
        return render_template('pros.html', pros=[], status_filter='', numero_filter='', data_inicio_filter='', data_fim_filter='', cod_ug_filter='', sigla_ug_filter='', status_list=STATUS_PRO)


@app.route('/pro/<int:id>')
@acesso_requerido('admin', 'usuario')
def detalhes_pro(id):
    """Detalhes de uma PRO específica - apenas admin e usuario"""
    try:
        pro = Pro.query\
            .options(
                db.joinedload(Pro.usuario),
                db.joinedload(Pro.solicitacao_origem)
                    .joinedload(SolicitacaoExtraPDRLOG.pedidos)
                    .joinedload(PedidoSolicitacao.itens)
            )\
            .filter(Pro.id == id)\
            .first()

        if not pro:
            flash('PRO não encontrada!', 'error')
            return redirect(url_for('listar_pros'))

        if pro.status in STATUS_PRO_VARIANTES:
            pro.status = STATUS_PRO_VARIANTES[pro.status]

        # Buscar todas as Notas de Crédito geradas a partir desta PRO
        ncs_geradas = NotaCredito.query.filter_by(pro_id=id).all()
        nc_relacionada = ncs_geradas[0] if ncs_geradas else None
        # Buscar a solicitação relacionada para o modal
        solicitacao = None
        diex_solicitacao_pro = ''
        diex_credito_list = []
        itens_om_nd = []
        if pro.solicitacao_pdrlog_id:
            solicitacao = db.session.get(SolicitacaoExtraPDRLOG, pro.solicitacao_pdrlog_id)
            diex_solicitacao_pro = solicitacao.diex if solicitacao and solicitacao.diex else ''

        if ncs_geradas:
            diex_vistos = set()
            for nc in ncs_geradas:
                diex_credito = (getattr(nc, 'diex_credito', '') or '').strip()
                if diex_credito and diex_credito not in diex_vistos:
                    diex_vistos.add(diex_credito)
                    diex_credito_list.append(diex_credito)

        if pro.solicitacao_pdrlog_id:
            itens_solicitacao = db.session.query(
                PedidoSolicitacao.om,
                PedidoSolicitacao.codom,
                ItemPedido.nd,
                ItemPedido.valor_aprovado,
                ItemPedido.valor_solicitado
            ).join(
                ItemPedido, ItemPedido.pedido_id == PedidoSolicitacao.id
            ).filter(
                PedidoSolicitacao.solicitacao_id == pro.solicitacao_pdrlog_id
            ).all()

            for item in itens_solicitacao:
                valor_aprovado = item.valor_aprovado if item.valor_aprovado is not None else 0
                valor_solicitado = item.valor_solicitado if item.valor_solicitado is not None else 0
                valor_om = valor_aprovado if valor_aprovado > 0 else valor_solicitado
                itens_om_nd.append({
                    'om': item.om or '',
                    'cod_om': item.codom or '',
                    'nd': item.nd or '',
                    'valor_om': valor_om
                })

        if not itens_om_nd and not pro.solicitacao_pdrlog_id and ncs_geradas:
            for nc in ncs_geradas:
                itens_om_nd.append({
                    'om': '',
                    'cod_om': '',
                    'nd': nc.nd or '',
                    'valor_om': nc.valor or 0
                })

        return render_template(
            'detalhes_pro.html',
            pro=pro,
            nc_relacionada=nc_relacionada,
            ncs_geradas=ncs_geradas,
            status_list=STATUS_PRO,
            solicitacao=solicitacao,
            itens_om_nd=itens_om_nd,
            diex_solicitacao_pro=diex_solicitacao_pro,
            diex_credito_list=diex_credito_list
        )
    except Exception as e:
        print(f"❌ Erro detalhado ao carregar PRO: {str(e)}")
        import traceback
        traceback.print_exc()
        flash(f'Erro ao carregar PRO: {str(e)}', 'error')
        return redirect(url_for('listar_pros'))


@app.route('/pro/<int:id>/converter_nc', methods=['GET', 'POST'])
@acesso_requerido('admin', 'usuario')
def converter_pro_em_nc(id):
    """Converte uma PRO em Nota de Crédito - com opção de editar valores - apenas admin e usuario"""
    # Primeiro carregar a PRO sem os relacionamentos complexos
    pro = Pro.query.filter(Pro.id == id).first()
    
    if not pro:
        flash('PRO não encontrada!', 'error')
        return redirect(url_for('listar_pros'))
    
    # Se for POST, processar a conversão ou avançar de etapa
    if request.method == 'POST':
        try:
            total_nc = 0
            if pro.status == 'Cancelada':
                flash('Não é possível converter uma PRO cancelada em Nota de Crédito!', 'error')
                return redirect(url_for('detalhes_pro', id=id))
            # Permitir múltiplas conversões parciais enquanto houver saldo
            if pro.valor_restante is not None and pro.valor_restante <= 0:
                flash('O saldo da PRO já foi totalmente convertido em Notas de Crédito!', 'warning')
                return redirect(url_for('detalhes_pro', id=id))

            solicitacao = SolicitacaoExtraPDRLOG.query\
                .options(db.joinedload(SolicitacaoExtraPDRLOG.pedidos))\
                .filter(SolicitacaoExtraPDRLOG.id == pro.solicitacao_pdrlog_id)\
                .first()
            if not solicitacao:
                flash('Não foi possível encontrar a solicitação relacionada!', 'error')
                return redirect(url_for('detalhes_pro', id=id))

            diex_nc = request.form.get('diex_nc', '').strip()
            data_diex_nc_str = request.form.get('data_diex_nc', '')
            om_diex_nc = request.form.get('om_diex_nc', '').strip()
            descricao_usuario_nc = request.form.get('descricao_nc_usuario', '').strip()
            acao = request.form.get('acao', '')
            session_key = f"pro_nc_diex_{id}"

            if acao == 'avancar':
                if not diex_nc or not data_diex_nc_str or not om_diex_nc:
                    flash('Preencha todos os campos obrigatórios!', 'error')
                    return redirect(url_for('converter_pro_em_nc', id=id))
                try:
                    datetime.strptime(data_diex_nc_str, '%Y-%m-%d')
                except ValueError:
                    flash('Data do DIEx inválida!', 'error')
                    return redirect(url_for('converter_pro_em_nc', id=id))

                session[session_key] = {
                    'diex_nc': diex_nc,
                    'data_diex_nc': data_diex_nc_str,
                    'om_diex_nc': om_diex_nc,
                    'descricao_nc_usuario': descricao_usuario_nc,
                }
                return redirect(url_for('converter_pro_em_nc', id=id, step='valores'))

            if (not diex_nc or not data_diex_nc_str or not om_diex_nc) and session_key in session:
                cached = session.get(session_key) or {}
                diex_nc = diex_nc or cached.get('diex_nc', '')
                data_diex_nc_str = data_diex_nc_str or cached.get('data_diex_nc', '')
                om_diex_nc = om_diex_nc or cached.get('om_diex_nc', '')
                if not descricao_usuario_nc:
                    descricao_usuario_nc = cached.get('descricao_nc_usuario', '')
            converter_tipo = request.form.get('converter_tipo')
            forcar_total = converter_tipo == 'total'

            if not diex_nc or not data_diex_nc_str or not om_diex_nc:
                flash('Preencha todos os campos obrigatórios!', 'error')
                return redirect(url_for('converter_pro_em_nc', id=id))
            try:
                data_diex_nc = datetime.strptime(data_diex_nc_str, '%Y-%m-%d')
            except ValueError:
                flash('Data do DIEx inválida!', 'error')
                return redirect(url_for('converter_pro_em_nc', id=id))

            # Extrair itens para conversao
            if forcar_total:
                itens_selecionados = [
                    item.id
                    for pedido in solicitacao.pedidos
                    for item in pedido.itens
                    if item.valor_aprovado and item.valor_aprovado > 0
                ]
            else:
                itens_selecionados = []
                for key in request.form:
                    if key.startswith('item_selecionado_') and request.form.get(key) == '1':
                        try:
                            item_id = int(key.replace('item_selecionado_', ''))
                            itens_selecionados.append(item_id)
                        except Exception:
                            continue

            if not itens_selecionados:
                flash('Selecione ao menos um item para converter em NC!', 'error')
                return redirect(url_for('converter_pro_em_nc', id=id))

            nc_count = 0
            nc_numero_base = f"NC-{pro.numero}-"
            ultimo_numero = (
                db.session.query(db.func.max(NotaCredito.numero))
                .filter(NotaCredito.numero.like(f"{nc_numero_base}%"))
                .scalar()
            )
            if ultimo_numero and ultimo_numero.startswith(nc_numero_base):
                try:
                    nc_count = int(ultimo_numero.replace(nc_numero_base, ""))
                except ValueError:
                    nc_count = 0
            for item_id in itens_selecionados:
                item = ItemPedido.query.get(item_id)
                if not item:
                    continue
                if item.valor_restante is None:
                    item.valor_restante = float(item.valor_aprovado or 0)
                saldo_item = item.valor_restante if item.valor_restante is not None else (item.valor_aprovado or 0)
                valor_str_raw = request.form.get(f'valor_nc_{item_id}', '')
                print(f"[DEBUG] Valor bruto recebido do formulário para item {item_id}: {valor_str_raw} (type: {type(valor_str_raw)})")
                valor_str = str(valor_str_raw).strip()
                # Se vier como número, não precisa converter
                if isinstance(valor_str_raw, float) or isinstance(valor_str_raw, int):
                    valor_aprovado = float(valor_str_raw)
                else:
                    valor_str = valor_str.replace('R$', '').replace(' ', '')
                    if ',' in valor_str and '.' in valor_str:
                        # Formato pt-BR com milhar e decimal: 180.320,99
                        valor_str = valor_str.replace('.', '').replace(',', '.')
                    elif ',' in valor_str:
                        # Decimal com virgula: 180320,99
                        valor_str = valor_str.replace(',', '.')
                    # Se vier com decimal ponto (ex: 180320.99), mantem como esta
                    try:
                        valor_aprovado = float(valor_str) if valor_str else 0.0
                    except (ValueError, TypeError):
                        valor_aprovado = 0.0

                # No modo "total", se campo vier vazio, usa o saldo do item como fallback.
                if forcar_total and valor_aprovado <= 0:
                    valor_aprovado = float(saldo_item or 0)

                print(f"[DEBUG] Valor final convertido para NC item {item_id}: {valor_aprovado}")
                # Só cria NC se valor_aprovado for positivo e diferente de None
                if item and valor_aprovado is not None and valor_aprovado > 0:
                    item.valor_restante = max(float(saldo_item or 0) - float(valor_aprovado or 0), 0)
                    pedido = getattr(item, 'pedido', None)
                    if not pedido and hasattr(item, 'pedido_id'):
                        pedido = PedidoSolicitacao.query.get(item.pedido_id)
                    pi = request.form.get(f'pi_nc_{item_id}', item.pi)
                    nd = request.form.get(f'nd_nc_{item_id}', item.nd)
                    om = request.form.get(f'om_nc_{item_id}', pedido.om if pedido else '')
                    finalidade = request.form.get(f'finalidade_nc_{item_id}', item.finalidade)
                    codom = pedido.codom if pedido and hasattr(pedido, 'codom') else ''
                    descricao_om = getattr(pedido, 'descricao_om', None) if pedido else ''
                    descricao_adicional = (descricao_om or '').strip()
                    nr_opus = (solicitacao.nr_opus or '').strip() if solicitacao else ''
                    # Sempre priorizar OPUS se existir
                    if nr_opus:
                        opus_texto = f"Nº OPUS: {nr_opus}"
                        if descricao_adicional:
                            # Se já contém OPUS, não duplica
                            if 'opus' not in descricao_adicional.lower():
                                descricao_adicional = f"{descricao_adicional} – {opus_texto}"
                        else:
                            descricao_adicional = opus_texto
                    data_diex_formatada = data_diex_nc.strftime('%d/%m/%Y') if data_diex_nc_str else ''
                    descricao_nc = construir_descricao_nc(
                        pi,
                        nd,
                        diex_nc,
                        data_diex_formatada,
                        om,
                        codom,
                        descricao_usuario_nc,
                        pro.numero,
                        descricao_om=descricao_adicional
                    )
                    nc_count += 1
                    numero_nc = f"{nc_numero_base}{nc_count}"
                    while NotaCredito.query.filter_by(numero=numero_nc).first() is not None:
                        nc_count += 1
                        numero_nc = f"{nc_numero_base}{nc_count}"
                    nc = NotaCredito(
                        numero=numero_nc,
                        cod_ug=pro.cod_ug,
                        sigla_ug=pro.sigla_ug,
                        pi=pi,
                        nd=nd,
                        valor=valor_aprovado,
                        diex_credito=diex_nc,
                        descricao=descricao_nc,
                        status='Não Conferida',
                        usuario_id=pro.usuario_id,
                        data_criacao=datetime.now(),
                        data_atualizacao=datetime.now(),
                        solicitacao_pdrlog_id=solicitacao.id,
                        pro_id=pro.id,
                        item_id=item.id
                    )
                    db.session.add(nc)
                    total_nc += valor_aprovado
                else:
                    print(f"[DEBUG] NC NÃO CRIADA para item {item_id} pois valor_aprovado={valor_aprovado}")
            db.session.commit()
            # Atualizar saldo da PRO com base no saldo restante por item
            saldo_itens = sum(
                (item.valor_restante if item.valor_restante is not None else (item.valor_aprovado or 0))
                for pedido in solicitacao.pedidos
                for item in pedido.itens
            )
            if forcar_total:
                # "Converter Totalmente" encerra a PRO para novas NCs,
                # independentemente do valor digitado para esta conversão.
                pro.valor_restante = 0
                pro.status = 'Convertida em NC'
            else:
                pro.valor_restante = float(saldo_itens)
                # Corrigir possíveis diferenças de ponto flutuante
                if abs(pro.valor_restante) < 0.01 or pro.valor_restante < 0:
                    pro.valor_restante = 0
                    pro.status = 'Convertida em NC'
                else:
                    pro.status = 'Parcialmente Convertida'
            pro.data_atualizacao = datetime.now()
            # Preservar DIEx original da solicitação da PRO.
            # O DIEx de solicitação de crédito é registrado em NotaCredito.diex_credito.
            if descricao_usuario_nc:
                if solicitacao.descricao:
                    solicitacao.descricao += f"\n\nDESCRIÇÃO PARA NC: {descricao_usuario_nc}"
                else:
                    solicitacao.descricao = f"DESCRIÇÃO PARA NC: {descricao_usuario_nc}"
            db.session.commit()
            session.pop(session_key, None)
            msg_total_nc = format_currency(total_nc)
            msg_saldo_pro = format_currency(pro.valor_restante)
            flash(f'{nc_count} Notas de Crédito geradas. Valor total: {msg_total_nc}. Saldo restante da PRO: {msg_saldo_pro}.', 'success')
        except Exception as e:
            db.session.rollback()
            print(f"❌ Erro ao converter PRO: {str(e)}")
            import traceback
            traceback.print_exc()
            flash(f'Erro ao converter PRO: {str(e)}', 'error')
        return redirect(url_for('detalhes_pro', id=id))
    
    # Se for GET, mostrar formulário de conversão
    try:
        # Verificar se a PRO está cancelada
        if pro.status == 'Cancelada':
            flash('Não é possível converter uma PRO cancelada em Nota de Crédito!', 'error')
            return redirect(url_for('detalhes_pro', id=id))
        
        if pro.status == 'Convertida em NC':
            flash('Esta PRO já foi convertida em Nota de Crédito!', 'warning')
            return redirect(url_for('detalhes_pro', id=id))
        
        # Buscar a solicitação relacionada com todos os relacionamentos
        solicitacao = SolicitacaoExtraPDRLOG.query\
            .options(
                db.joinedload(SolicitacaoExtraPDRLOG.pedidos).joinedload(PedidoSolicitacao.itens)
            )\
            .filter(SolicitacaoExtraPDRLOG.id == pro.solicitacao_pdrlog_id)\
            .first()
        
        if not solicitacao:
            flash('Não foi possível encontrar a solicitação relacionada!', 'error')
            return redirect(url_for('detalhes_pro', id=id))
        
        # Coletar todos os itens cadastrados na solicitação vinculada para exibir no formulário
        itens_aprovados = []
        for pedido in solicitacao.pedidos:
            for item in pedido.itens:
                valor_disponivel = item.valor_restante if item.valor_restante is not None else item.valor_aprovado
                valor_disponivel = float(valor_disponivel or 0)
                valor_disponivel_formatado = f"{valor_disponivel:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                itens_aprovados.append({
                    'pedido': pedido,
                    'item': item,
                    'om': pedido.om,
                    'nd': item.nd,
                    'finalidade': item.finalidade,
                    'pi': item.pi,
                    'valor_aprovado': item.valor_aprovado,
                    'valor_disponivel': valor_disponivel,
                    'valor_disponivel_formatado': valor_disponivel_formatado
                })
        
        step = request.args.get('step', 'dados')
        session_key = f"pro_nc_diex_{id}"
        diex_cache = session.get(session_key, {})

        if step == 'valores' and not diex_cache:
            flash('Informe os dados do DIEx antes de definir os valores.', 'warning')
            return redirect(url_for('converter_pro_em_nc', id=id))

        return render_template('converter_pro_nc.html',
                            pro=pro,
                            solicitacao=solicitacao,
                            itens_aprovados=itens_aprovados,
                            PI_POR_FINALIDADE=PI_POR_FINALIDADE,
                    oms_data=_obter_oms_data(),
                    step=step,
                    diex_cache=diex_cache)
        
    except Exception as e:
        print(f"❌ Erro ao carregar formulário de conversão: {str(e)}")
        import traceback
        traceback.print_exc()
        flash(f'Erro ao carregar formulário de conversão: {str(e)}', 'error')
        return redirect(url_for('detalhes_pro', id=id))

@app.route('/pro/<int:id>/cancelar', methods=['POST'])
@acesso_requerido('admin', 'usuario')
def cancelar_pro(id):
    """Cancela uma PRO - apenas admin e usuario"""
    pro = db.session.get(Pro, id)
    if not pro:
        flash('PRO não encontrada!', 'error')
        return redirect(url_for('listar_pros'))
    
    try:
        # Verificar se a PRO já está cancelada
        if pro.status == 'Cancelada':
            flash('Esta PRO já está cancelada!', 'warning')
            return redirect(url_for('detalhes_pro', id=id))
        
        # Permitir cancelar PRO convertida em NC se não houver NC ativa
        # (mantém bloqueio apenas se houver NC ativa)
        
        # Verificar se existe NC relacionada ATIVA (não cancelada)
        nc_ativas = NotaCredito.query.filter_by(pro_id=id).filter(NotaCredito.status != 'Cancelada').all()
        if nc_ativas:
            flash('Não é possível cancelar uma PRO que possui Nota de Crédito ativa relacionada!', 'error')
            return redirect(url_for('detalhes_pro', id=id))
        
        # Cancelar a PRO
        pro.status = 'Cancelada'
        pro.data_atualizacao = datetime.now()
        
        db.session.commit()
        
        print(f"✅ PRO {pro.numero} cancelada com sucesso")
        flash(f'PRO {pro.numero} cancelada com sucesso!', 'success')
        
    except Exception as e:
        db.session.rollback()
        print(f"❌ Erro ao cancelar PRO: {str(e)}")
        import traceback
        traceback.print_exc()
        flash(f'Erro ao cancelar PRO: {str(e)}', 'error')
    
    return redirect(url_for('detalhes_pro', id=id))

@app.route('/pro/<int:id>/excluir', methods=['POST'])
@acesso_requerido('admin')
def excluir_pro(id):
    """Exclui uma PRO do sistema (apenas admin)"""
    pro = db.session.get(Pro, id)
    if not pro:
        flash('PRO não encontrada!', 'error')
        return redirect(url_for('listar_pros'))
    try:
        db.session.delete(pro)
        db.session.commit()
        flash('PRO excluída com sucesso!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao excluir PRO: {str(e)}', 'error')
    return redirect(url_for('listar_pros'))

# ===== ROTAS DE RELATÓRIOS (APENAS ADMIN E USUARIO) =====

@app.route('/relatorios')
@acesso_requerido('admin', 'usuario')
def relatorios():
    """Página de relatórios - apenas admin e usuario"""
    try:
        # Estatísticas para a página de relatórios
        total_solicitacoes = SolicitacaoExtraPDRLOG.query.count()
        solicitacoes_ultimo_mes = SolicitacaoExtraPDRLOG.query.filter(
            SolicitacaoExtraPDRLOG.data_solicitacao >= datetime.now() - timedelta(days=30)
        ).count()
        
        return render_template('relatorios.html',
                            total_solicitacoes=total_solicitacoes,
                            solicitacoes_ultimo_mes=solicitacoes_ultimo_mes,
                            finalidades=FINALIDADES,
                            status_list=STATUS_SOLICITACAO)
    except Exception as e:
        flash(f'Erro ao carregar relatórios: {str(e)}', 'error')
        return render_template('relatorios.html',
                            total_solicitacoes=0,
                            solicitacoes_ultimo_mes=0,
                            finalidades=FINALIDADES,
                            status_list=STATUS_SOLICITACAO)

@app.route('/relatorios/despacho_cmt')
@acesso_requerido('admin', 'usuario')
def relatorio_despacho_cmt():
    """Relatório de solicitações aguardando despacho do Cmt - apenas admin e usuario"""
    try:
        # Buscar solicitações com status "Aguardando despacho"
        solicitacoes = SolicitacaoExtraPDRLOG.query.filter_by(
            status='Aguardando despacho'
        ).order_by(SolicitacaoExtraPDRLOG.data_criacao.desc()).all()
        
        return render_template('relatorio_despacho_cmt.html',
                            solicitacoes=solicitacoes)
    except Exception as e:
        flash(f'Erro ao gerar relatório: {str(e)}', 'error')
        return render_template('relatorio_despacho_cmt.html',
                            solicitacoes=[])


@app.route('/relatorios/despacho_cmt/pdf')
@acesso_requerido('admin', 'usuario')
def relatorio_despacho_cmt_pdf():
    """Gera PDF do Relatório 'Despacho com o Cmt' contendo DIEx, OM, Descrição da OM,
    ND, Finalidade, Valor solicitado, Parecer da Análise e campo em branco para Despacho Ch Sup."""
    try:
        solicitacoes = SolicitacaoExtraPDRLOG.query.filter_by(
            status='Aguardando despacho'
        ).order_by(SolicitacaoExtraPDRLOG.data_criacao.desc()).all()

        buffer = BytesIO()
        # Usar orientação paisagem para acomodar todas as colunas
        doc = SimpleDocTemplate(buffer, pagesize=(A4[1], A4[0]),
                    rightMargin=1.5*cm, leftMargin=1.5*cm,
                    topMargin=1.5*cm, bottomMargin=2*cm)

        styles = getSampleStyleSheet()
        header_style = ParagraphStyle('Header', parent=styles['Normal'], fontSize=12, alignment=1, fontName='Helvetica-Bold')
        title_style = ParagraphStyle('Title', parent=styles['Normal'], fontSize=14, alignment=1, fontName='Helvetica-Bold')
        normal_style = ParagraphStyle('Normal', parent=styles['Normal'], fontSize=10, fontName='Helvetica')
        small_style = ParagraphStyle('Small', parent=styles['Normal'], fontSize=9, fontName='Helvetica')
        valor_style = ParagraphStyle('Valor', parent=styles['Normal'], fontSize=9, fontName='Helvetica', alignment=2)

        story = []
        story.append(Paragraph('Relatório - Despacho com o Comandante', title_style))
        story.append(Spacer(1, 0.3*cm))

        # Tabela principal por itens
        table_header = [
            Paragraph('DIEx', small_style),
            Paragraph('OM', small_style),
            Paragraph('Descrição:', small_style),
            Paragraph('ND', small_style),
            Paragraph('Finalidade', small_style),
            Paragraph('Valor solicitado (R$)', small_style),
            Paragraph('Div Subs', small_style),
            Paragraph('Despacho Ch Sup', small_style)
        ]

        # Construir linhas
        table_data = [table_header]

        for solicitacao in solicitacoes:
            diex = solicitacao.diex or ''
            for pedido in solicitacao.pedidos:
                om = pedido.om or ''
                desc_om = (pedido.descricao_om or '') if hasattr(pedido, 'descricao_om') else ''
                for item in pedido.itens:
                    nd = item.nd or ''
                    finalidade = item.finalidade or solicitacao.finalidade or ''
                    valor = item.valor_solicitado if hasattr(item, 'valor_solicitado') else (getattr(item, 'valor', 0) or 0)
                    valor_txt = f'R$ {float(valor):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
                    # Parecer da análise (por solicitação) será repetido em cada linha
                    parecer = solicitacao.parecer_analise or ''
                    row = [
                        Paragraph(diex, small_style),
                        Paragraph(om, small_style),
                        Paragraph(desc_om, small_style),
                        Paragraph(nd, small_style),
                        Paragraph(finalidade, small_style),
                        Paragraph(valor_txt, valor_style),
                        Paragraph(parecer, small_style),
                        Paragraph('', small_style)
                    ]
                    table_data.append(row)
            # Não adicionar bloco separado de parecer por solicitação (parecer já é coluna 'Div Subs')
            pass

        # Se não houver linhas além do cabeçalho, mostrar aviso
        if len(table_data) == 1:
            story.append(Paragraph('Nenhuma solicitação aguardando despacho.', normal_style))
        else:
            # Criar tabela em orientação paisagem com larguras ajustadas para caber na página
            col_widths = [2.8*cm, 3.5*cm, 6.0*cm, 2.2*cm, 4.0*cm, 3.0*cm, 3.5*cm, 1.7*cm]
            table = Table(table_data, colWidths=col_widths, repeatRows=1)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#87CEEB')),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,0), 9),
                ('ALIGN', (0,0), (4,-1), 'CENTER'),
                ('ALIGN', (5,1), (5,-1), 'RIGHT'),
                ('ALIGN', (6,1), (6,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor('#dee2e6')),
                ('LEFTPADDING', (0,0), (-1,-1), 4),
                ('RIGHTPADDING', (0,0), (-1,-1), 4),
            ]))
            story.append(table)

        # Linha de despacho com a data atual
        hoje = datetime.now().strftime('%d/%m/%Y')
        story.append(Spacer(1, 0.4*cm))
        story.append(Paragraph(f'Despacho em {hoje}.', normal_style))

        doc.build(story)
        buffer.seek(0)

        nome_arquivo = f"relatorio_despacho_cmt_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        response = make_response(buffer.read())
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'attachment; filename="{nome_arquivo}"'
        return response

    except Exception as e:
        print(f"Erro ao gerar PDF despacho Cmt: {e}")
        import traceback
        traceback.print_exc()
        flash(f'Erro ao gerar PDF: {str(e)}', 'error')
        return redirect(url_for('relatorio_despacho_cmt'))

@app.route('/exportar/excel')
@acesso_requerido('admin', 'usuario')
def exportar_excel():
    """Exporta solicitações para Excel - apenas admin e usuario"""
    try:
        # Aplicar filtros
        status_filter = request.args.get('status', '')
        finalidade_filter = request.args.get('finalidade', '')
        data_inicio = request.args.get('data_inicio', '')
        data_fim = request.args.get('data_fim', '')
        
        query = SolicitacaoExtraPDRLOG.query
        
        if status_filter:
            query = query.filter_by(status=status_filter)
        if finalidade_filter:
            query = query.filter_by(finalidade=finalidade_filter)
        if data_inicio:
            data_inicio = datetime.strptime(data_inicio, '%Y-%m-%d')
            query = query.filter(SolicitacaoExtraPDRLOG.data_solicitacao >= data_inicio)
        if data_fim:
            data_fim = datetime.strptime(data_fim, '%Y-%m-%d')
            query = query.filter(SolicitacaoExtraPDRLOG.data_solicitacao <= data_fim)
        
        solicitacoes = query.order_by(SolicitacaoExtraPDRLOG.data_solicitacao.desc()).all()
        
        # Exportar cada NC como uma linha
        data = []
        # Relatório: todas as NCs do sistema
        ncs = NotaCredito.query.order_by(NotaCredito.data_criacao.desc()).all()
        data = []
        for nc in ncs:
            solicitacao = nc.solicitacao_origem if hasattr(nc, 'solicitacao_origem') else None
            cod_om = ''
            cod_ug = nc.cod_ug or ''
            sigla_ug = nc.sigla_ug or ''
            if nc.pro_id and nc.pro_origem:
                cod_om = nc.pro_origem.cod_ug or ''
            elif nc.solicitacao_pdrlog_id and solicitacao:
                for pedido in solicitacao.pedidos:
                    cod_om = pedido.codom or ''
                    if cod_om:
                        break
            data.append({
                'Número Solicitação': solicitacao.numero if solicitacao else '',
                'DIEx': solicitacao.diex if solicitacao else '',
                'Órgão Demandante': solicitacao.orgao_demandante if solicitacao else '',
                'Data Solicitação': solicitacao.data_solicitacao.strftime('%d/%m/%Y') if solicitacao and solicitacao.data_solicitacao else '',
                'Finalidade Solicitação': solicitacao.finalidade if solicitacao else '',
                'Status Solicitação': solicitacao.status if solicitacao else '',
                'Modalidade': solicitacao.modalidade if solicitacao else '',
                'Descrição Solicitação': solicitacao.descricao if solicitacao else '',
                'Número NC': nc.numero,
                'Valor NC': nc.valor,
                'PI': nc.pi or '',
                'COD OM': cod_om,
                'COD UG': cod_ug,
                'Sigla UG': sigla_ug,
                'ND': nc.nd or '',
                'Status NC': nc.status or '',
                'Data NC': nc.data_criacao.strftime('%d/%m/%Y') if nc.data_criacao else '',
                'Descrição NC': nc.descricao or '',
                'Ref SISNC': nc.ref_sisnc or '',
                'NC SIAFI': nc.nc_siafi or '',
                'Motivo Cancelamento': getattr(nc, 'motivo_cancelamento', '') or '',
                'Data Atualização NC': nc.data_atualizacao.strftime('%d/%m/%Y') if hasattr(nc, 'data_atualizacao') and nc.data_atualizacao else '',
                'Usuário NC': nc.usuario.nome if hasattr(nc, 'usuario') and nc.usuario else ''
            })
        df = pd.DataFrame(data)
        
        # Criar arquivo Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Solicitações', index=False)
            
            # Ajustar largura das colunas
            worksheet = writer.sheets['Solicitações']
            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).str.len().max(), len(col)) + 2
                worksheet.column_dimensions[chr(65 + idx)].width = max_len
        
        output.seek(0)
        
        return send_file(
            output,
            download_name=f'relatorio_pdrlog_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        flash(f'Erro ao exportar Excel: {str(e)}', 'error')
        return redirect(url_for('relatorios'))

@app.route('/exportar/csv')
@acesso_requerido('admin', 'usuario')
def exportar_csv():
    """Exporta solicitações para CSV - apenas admin e usuario"""
    try:
        # Aplicar filtros se existirem
        status_filter = request.args.get('status', '')
        finalidade_filter = request.args.get('finalidade', '')
        data_inicio = request.args.get('data_inicio', '')
        data_fim = request.args.get('data_fim', '')
        
        query = SolicitacaoExtraPDRLOG.query
        
        if status_filter:
            query = query.filter_by(status=status_filter)
        if finalidade_filter:
            query = query.filter_by(finalidade=finalidade_filter)
        if data_inicio:
            data_inicio = datetime.strptime(data_inicio, '%Y-%m-%d')
            query = query.filter(SolicitacaoExtraPDRLOG.data_solicitacao >= data_inicio)
        if data_fim:
            data_fim = datetime.strptime(data_fim, '%Y-%m-%d')
            query = query.filter(SolicitacaoExtraPDRLOG.data_solicitacao <= data_fim)
        
        solicitacoes = query.order_by(SolicitacaoExtraPDRLOG.data_solicitacao.desc()).all()
        
        output = StringIO()
        writer = csv.writer(output, delimiter=';')
        
        # Cabeçalho atualizado
        writer.writerow([
            'Número', 'DIEx', 'Órgão Demandante', 'Data', 'Finalidade', 'Status', 'Descrição'
        ])
        
        # Dados
        for solicitacao in solicitacoes:
            writer.writerow([
                solicitacao.numero or '',
                solicitacao.diex or '',
                solicitacao.orgao_demandante or '',
                solicitacao.data_solicitacao.strftime('%d/%m/%Y') if solicitacao.data_solicitacao else '',
                solicitacao.finalidade or '',
                solicitacao.status or '',
                solicitacao.descricao or ''
            ])
        
        output.seek(0)
        
        bio = BytesIO()
        bio.write(output.getvalue().encode('utf-8'))
        bio.seek(0)
        
        return send_file(
            bio,
            download_name=f'relatorio_pdrlog_{datetime.now().strftime("%Y%m%d")}.csv',
            as_attachment=True,
            mimetype='text/csv'
        )
    except Exception as e:
        flash(f'Erro ao exportar CSV: {str(e)}', 'error')
        return redirect(url_for('relatorios'))

# ===== APIs =====

@app.route('/api/oms')
@login_required
def api_oms():
    """API para busca de OMs"""
    termo = request.args.get('q', '').lower()
    base_oms = _carregar_tabela_oms()
    if termo:
        resultados = [
            om for om in base_oms
            if termo in (om.get('OM', '').lower())
            or termo in (str(om.get('CODOM', '')).lower())
            or termo in (str(om.get('CODUG', '')).lower())
            or termo in (om.get('SIGLA_UG', '').lower())
        ]
    else:
        resultados = base_oms  # retorna todas se não houver filtro
    
    resp = jsonify(resultados)
    resp.headers['X-OMS-Source'] = _caminho_tabela_oms()
    resp.headers['X-OMS-Total'] = str(len(base_oms))
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp

@app.route('/api/om/<string:nome_om>')
@login_required
def api_om_detalhes(nome_om):
    """API para obter detalhes de uma OM específica"""
    nome_ref = (nome_om or '').strip().upper()
    base_oms = _carregar_tabela_oms()
    for om in base_oms:
        if (om.get('OM', '') or '').strip().upper() == nome_ref:
            resp = jsonify(om)
            resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
            resp.headers['Pragma'] = 'no-cache'
            resp.headers['Expires'] = '0'
            return resp
    resp = jsonify({'error': 'OM não encontrada'})
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp, 404


@app.route('/api/v2/oms')
@login_required
def api_oms_v2():
    termo = (request.args.get('q', '') or '').strip().lower()
    base_oms = _carregar_tabela_oms()
    if termo:
        itens = [
            om for om in base_oms
            if termo in (om.get('OM', '').lower())
            or termo in (str(om.get('CODOM', '')).lower())
            or termo in (str(om.get('CODUG', '')).lower())
            or termo in (om.get('SIGLA_UG', '').lower())
        ]
    else:
        itens = base_oms

    resp = jsonify({
        'items': itens,
        'total': len(base_oms),
        'source': _caminho_tabela_oms()
    })
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp


@app.route('/api/v2/om/<string:nome_om>')
@login_required
def api_om_detalhes_v2(nome_om):
    nome_ref = (nome_om or '').strip().upper()
    base_oms = _carregar_tabela_oms()
    for om in base_oms:
        if (om.get('OM', '') or '').strip().upper() == nome_ref:
            resp = jsonify({'item': om, 'source': _caminho_tabela_oms()})
            resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
            resp.headers['Pragma'] = 'no-cache'
            resp.headers['Expires'] = '0'
            return resp

    resp = jsonify({'item': None, 'source': _caminho_tabela_oms(), 'error': 'OM não encontrada'})
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp, 404

@app.route('/api/dashboard/stats')
@login_required
def api_dashboard_stats():
    """API para dados do dashboard"""
    try:
        total_solicitacoes = SolicitacaoExtraPDRLOG.query.count()
        solicitacoes_aguardando = SolicitacaoExtraPDRLOG.query.filter_by(status='Aguardando Análise').count()
        solicitacoes_aprovadas = SolicitacaoExtraPDRLOG.query.filter_by(status='Aprovado Ch Sup').count()
        
        solicitacoes_por_finalidade = {}
        for finalidade in FINALIDADES:
            count = SolicitacaoExtraPDRLOG.query.filter_by(finalidade=finalidade).count()
            solicitacoes_por_finalidade[finalidade] = count
        
        solicitacoes_por_status = {}
        for status in STATUS_SOLICITACAO:
            count = SolicitacaoExtraPDRLOG.query.filter_by(status=status).count()
            solicitacoes_por_status[status] = count

        return jsonify({
            'total_solicitacoes': total_solicitacoes,
            'solicitacoes_aguardando': solicitacoes_aguardando,
            'solicitacoes_aprovadas': solicitacoes_aprovadas,
            'solicitacoes_por_finalidade': solicitacoes_por_finalidade,
            'solicitacoes_por_status': solicitacoes_por_status
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ===== ROTA PARA TRANSFERÊNCIAS INTERNAS COLOG =====
@app.route('/transferencias')
@acesso_requerido('admin', 'usuario')
def listar_transferencias():
    """Lista de Transferências Interna COLOG"""
    transferencias = SolicitacaoExtraPDRLOG.query.filter_by(modalidade='Transferência Interna COLOG').order_by(SolicitacaoExtraPDRLOG.data_criacao.desc()).all()
    return render_template('transferencia_interna_colog.html', transferencias=transferencias)


# ===== AUDITORIA NE – CONFORMIDADE NOTA DE EMPENHO × NORMAS =====

def _format_nd_code(val):
    """Converte código ND numérico (ex: 339030.0) para formato XX.XX.XX."""
    try:
        s = str(int(float(val))).zfill(6)
        return f'{s[0:2]}.{s[2:4]}.{s[4:6]}'
    except Exception:
        return str(val)


def _format_subitem_ne(val):
    """Formata subitem numérico (ex: 17.0) para string zero-padded de 2 dígitos (ex: '17')."""
    try:
        return str(int(float(val))).zfill(2)
    except Exception:
        return str(val)


def _format_valor_br(val):
    """Formata valor numérico para padrão brasileiro com milhar e 2 casas decimais."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return '0,00'
    try:
        if isinstance(val, str):
            txt = val.strip()
            if not txt:
                return '0,00'
            txt = txt.replace('R$', '').replace(' ', '')
            if ',' in txt and '.' in txt:
                txt = txt.replace('.', '').replace(',', '.')
            elif ',' in txt:
                txt = txt.replace(',', '.')
            numero = float(txt)
        else:
            numero = float(val)
    except Exception:
        return str(val)
    return f'{numero:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def _build_enquadramento_lookup():
    """
    Carrega Enquadramento Despesas.xlsx e constrói dicionário de conformidade:
      (PI_UPPER, 'XX.XX.XX') -> set(['01', '02', ...])
    Prefere versão em instance/; caso ausente, usa a da raiz do projeto.
    """
    instance_enq = os.path.join(app.instance_path, 'enquadramento_despesas.xlsx')
    root_enq = os.path.join(os.path.dirname(__file__), 'Enquadramento Despesas.xlsx')
    enq_path = instance_enq if os.path.exists(instance_enq) else root_enq
    if not os.path.exists(enq_path):
        raise FileNotFoundError(
            'Planilha Enquadramento Despesas.xlsx não encontrada. '
            'Coloque-a na raiz do projeto ou faça upload via admin.'
        )
    df_e = pd.read_excel(enq_path, header=0)
    lookup = {}
    for _, row in df_e.iterrows():
        pi_raw = row.get('Plano Interno', '')
        sub_raw = row.get('Subelemento (ND)', '')
        if pd.isna(pi_raw) or pd.isna(sub_raw):
            continue
        pi = str(pi_raw).strip().upper()
        sub_str = str(sub_raw).strip()
        parts = sub_str.split('.')
        if len(parts) < 4:
            continue
        nd = '.'.join(parts[:3])
        # last segment may be 'xx,yy,zz'
        subitems = [s.strip().zfill(2) for s in parts[3].split(',') if s.strip()]
        key = (pi, nd)
        if key not in lookup:
            lookup[key] = set()
        lookup[key].update(subitems)
    return lookup


def _resolver_rm_sigla_por_ug(ug_cod='', ug_nome='', base_oms=None):
    if base_oms is None:
        base_oms = _obter_oms_data()
    codug_ref = normalizar_cod_ug(ug_cod)
    om_ref = normalizar_om_solicitante_chave(ug_nome)

    def chave_om(item):
        return normalizar_om_solicitante_chave(item.get('OM', ''))

    def chave_sigla(item):
        return normalizar_om_solicitante_chave(item.get('SIGLA_UG', ''))

    candidatos_codug = []
    if codug_ref:
        for om in base_oms:
            if normalizar_cod_ug(om.get('CODUG', '')) == codug_ref:
                candidatos_codug.append(om)

    escolhido = None

    # Regra principal: UG do item analisado corresponde ao COD UG da tabela.
    # Se houver mais de um candidato, prioriza vínculo por OM/Sigla UG.
    if candidatos_codug:
        if om_ref:
            for om in candidatos_codug:
                if om_ref == chave_om(om) or om_ref == chave_sigla(om):
                    escolhido = om
                    break
        if not escolhido:
            escolhido = candidatos_codug[0]

    # Fallback: quando OM = Sigla UG, encontrar RM mesmo sem COD UG válido.
    if not escolhido and om_ref:
        for om in base_oms:
            if om_ref == chave_om(om) or om_ref == chave_sigla(om):
                escolhido = om
                break

    if escolhido:
        rm = str(escolhido.get('RM', '') or '').strip()
        sigla_ug = str(escolhido.get('SIGLA_UG', '') or '').strip()
        if not rm and codug_ref:
            rm = _obter_rm_por_codug(codug_ref)
        return rm, sigla_ug

    return (_obter_rm_por_codug(codug_ref) if codug_ref else ''), ''


def _enriquecer_itens_auditoria_ne(itens):
    base_oms = _obter_oms_data()
    atualizados = False
    for item in (itens or []):
        rm_atual = str(item.get('rm', '') or '').strip()
        sigla_atual = str(item.get('sigla_ug', '') or '').strip()
        if rm_atual and sigla_atual:
            continue

        rm_novo, sigla_nova = _resolver_rm_sigla_por_ug(
            item.get('ug_cod', ''),
            item.get('ug_nome', ''),
            base_oms=base_oms,
        )
        if not rm_atual and rm_novo:
            item['rm'] = rm_novo
            atualizados = True
        if not sigla_atual and sigla_nova:
            item['sigla_ug'] = sigla_nova
            atualizados = True

    return itens, atualizados


def _filtrar_itens_auditoria_ne(itens, filtro_texto='', filtro_conformidade='', filtro_pi=''):
    texto_ref = str(filtro_texto or '').strip().lower()
    conformidade_ref = str(filtro_conformidade or '').strip()
    pi_ref = str(filtro_pi or '').strip()

    filtrados = []
    for item in (itens or []):
        if conformidade_ref and str(item.get('conformidade', '') or '') != conformidade_ref:
            continue
        if pi_ref and str(item.get('pi', '') or '') != pi_ref:
            continue
        if texto_ref:
            valores = [
                item.get('rm', ''),
                item.get('ug_cod', ''),
                item.get('ug_nome', ''),
                item.get('sigla_ug', ''),
                item.get('ne', ''),
                item.get('pi', ''),
                item.get('nd_fmt', ''),
                item.get('subitem_fmt', ''),
                item.get('subitem_nome', ''),
                item.get('item_desc', ''),
                item.get('item_num', ''),
                item.get('valor', ''),
                item.get('conformidade', ''),
            ]
            texto_item = ' '.join(str(v or '') for v in valores).lower()
            if texto_ref not in texto_item:
                continue
        filtrados.append(item)

    return filtrados


def _dados_exportacao_auditoria_ne(itens):
    linhas = []
    for item in (itens or []):
        linhas.append({
            'RM': item.get('rm', ''),
            'UG': item.get('ug_cod', ''),
            'Sigla UG': item.get('sigla_ug', ''),
            'NE': item.get('ne', ''),
            'PI': item.get('pi', ''),
            'Natureza Despesa': item.get('nd_fmt', ''),
            'Subitem': item.get('subitem_fmt', ''),
            'Nome Subitem': item.get('subitem_nome', ''),
            'Descrição do Item': item.get('item_desc', ''),
            'Qtde Item': item.get('qtde_item', ''),
            'Valor': item.get('valor', ''),
            'Conformidade': item.get('conformidade', ''),
        })
    return linhas


def _carregar_cache_auditoria_ne():
    ne_cache_path = os.path.join(app.instance_path, 'ne_auditoria_cache.json')
    if not os.path.exists(ne_cache_path):
        return [], None, ne_cache_path

    with open(ne_cache_path, encoding='utf-8') as f:
        cache = json.load(f)

    itens = cache.get('itens', [])
    timestamp = cache.get('timestamp')
    itens, mudou = _enriquecer_itens_auditoria_ne(itens)
    if mudou:
        with open(ne_cache_path, 'w', encoding='utf-8') as f:
            json.dump({'itens': itens, 'timestamp': timestamp}, f, ensure_ascii=False)
    return itens, timestamp, ne_cache_path


def _exportar_auditoria_ne_excel(itens):
    linhas = _dados_exportacao_auditoria_ne(itens)
    df = pd.DataFrame(linhas)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Itens Analisados', index=False)
        worksheet = writer.sheets['Itens Analisados']
        larguras = {
            'A': 12,
            'B': 12,
            'C': 18,
            'D': 22,
            'E': 16,
            'F': 18,
            'G': 10,
            'H': 28,
            'I': 80,
            'J': 8,
            'K': 16,
            'L': 18,
        }
        for coluna, largura in larguras.items():
            worksheet.column_dimensions[coluna].width = largura
    output.seek(0)
    return send_file(
        output,
        download_name=f'auditoria_ne_itens_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


def _exportar_auditoria_ne_pdf(itens, timestamp=None):
    styles = getSampleStyleSheet()
    titulo_style = ParagraphStyle('TituloAuditoriaNE', parent=styles['Normal'], fontSize=12, alignment=1, fontName='Helvetica-Bold')
    texto_style = ParagraphStyle('TextoAuditoriaNE', parent=styles['Normal'], fontSize=6.3, leading=7.2, fontName='Helvetica')
    cabecalho_style = ParagraphStyle('CabecalhoAuditoriaNE', parent=styles['Normal'], fontSize=6.5, leading=7, alignment=1, fontName='Helvetica-Bold')
    valor_style = ParagraphStyle('ValorAuditoriaNE', parent=texto_style, alignment=2)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=(A4[1], A4[0]),
        rightMargin=0.8 * cm,
        leftMargin=0.8 * cm,
        topMargin=0.9 * cm,
        bottomMargin=1.0 * cm,
    )

    story = [Paragraph('Auditoria de Conformidade - Notas de Empenho - Itens Analisados', titulo_style)]
    if timestamp:
        story.append(Spacer(1, 0.15 * cm))
        story.append(Paragraph(f'Gerado em {datetime.now().strftime("%d/%m/%Y %H:%M")} | Última análise: {timestamp}', texto_style))
    story.append(Spacer(1, 0.25 * cm))

    dados = [[
        Paragraph('RM', cabecalho_style),
        Paragraph('UG', cabecalho_style),
        Paragraph('Sigla UG', cabecalho_style),
        Paragraph('NE', cabecalho_style),
        Paragraph('PI', cabecalho_style),
        Paragraph('ND', cabecalho_style),
        Paragraph('Subitem', cabecalho_style),
        Paragraph('Nome Subitem', cabecalho_style),
        Paragraph('Descrição do Item', cabecalho_style),
        Paragraph('Qtde Item', cabecalho_style),
        Paragraph('Valor', cabecalho_style),
        Paragraph('Conformidade', cabecalho_style),
    ]]

    for item in (itens or []):
        dados.append([
            Paragraph(str(item.get('rm', '') or '-'), texto_style),
            Paragraph(str(item.get('ug_cod', '') or ''), texto_style),
            Paragraph(str(item.get('sigla_ug', '') or '-'), texto_style),
            Paragraph(str(item.get('ne', '') or ''), texto_style),
            Paragraph(str(item.get('pi', '') or ''), texto_style),
            Paragraph(str(item.get('nd_fmt', '') or ''), texto_style),
            Paragraph(str(item.get('subitem_fmt', '') or ''), texto_style),
            Paragraph(str(item.get('subitem_nome', '') or ''), texto_style),
            Paragraph(str(item.get('item_desc', '') or ''), texto_style),
            Paragraph(str(item.get('qtde_item', '') or ''), texto_style),
            Paragraph(str(item.get('valor', '') or ''), valor_style),
            Paragraph(str(item.get('conformidade', '') or ''), texto_style),
        ])

    tabela = Table(
        dados,
        repeatRows=1,
        colWidths=[1.2 * cm, 1.7 * cm, 2.3 * cm, 2.8 * cm, 2.1 * cm, 1.8 * cm, 1.3 * cm, 3.0 * cm, 7.9 * cm, 0.9 * cm, 2.0 * cm, 2.3 * cm],
    )
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#d9b100')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('ALIGN', (9, 1), (9, -1), 'CENTER'),
        ('ALIGN', (10, 1), (10, -1), 'RIGHT'),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#c7ced6')),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(tabela)
    doc.build(story)
    buffer.seek(0)

    response = make_response(buffer.read())
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = f'attachment; filename="auditoria_ne_itens_{datetime.now().strftime("%Y%m%d_%H%M")}.pdf"'
    return response


def _processar_auditoria_ne(file_obj):
    """
    Lê planilha NE (cabeçalho na 3ª linha) e retorna lista de dicts por item.
    Colunas usadas:
      D – Natureza Despesa   E – PI   I – NE Item - Subelemento Código Subitem
    Demais colunas são capturadas para exibição.
    """
    df = pd.read_excel(file_obj, header=2)
    # Forward-fill colunas de cabeçalho da NE (aparecem apenas na 1ª linha de cada NE)
    header_cols = [
        'NE', 'Emitente - UG Código', 'Emitente - UG Nome',
        'Natureza Despesa', 'PI', 'NE - Descrição',
    ]
    for col in header_cols:
        if col in df.columns:
            df[col] = df[col].ffill()

    subitem_col = 'NE Item - Subelemento Código Subitem'
    if subitem_col not in df.columns:
        raise ValueError(
            f'Coluna "{subitem_col}" não encontrada. Verifique se o cabeçalho '
            'está na 3ª linha da planilha NE.'
        )

    # Descrição do item: prefere coluna nomeada e usa fallback pela posição L (índice 11)
    item_desc_col = 'NE Item Descrição'
    if item_desc_col not in df.columns and len(df.columns) > 11:
        item_desc_col = df.columns[11]

    df = df[df[subitem_col].notna()].copy()

    lookup = _build_enquadramento_lookup()
    base_oms = _obter_oms_data()

    itens = []
    for _, row in df.iterrows():
        pi = str(row.get('PI', '') or '').strip().upper()
        nd_fmt = _format_nd_code(row.get('Natureza Despesa', ''))
        subitem_fmt = _format_subitem_ne(row[subitem_col])
        key = (pi, nd_fmt)
        if key in lookup:
            conf = 'Conforme' if subitem_fmt in lookup[key] else 'Não Conforme'
        else:
            conf = 'PI Não Mapeado'

        # Normalize UG code (remove trailing '.0' from float)
        ug_cod_raw = str(row.get('Emitente - UG Código', '') or '').strip()
        if ug_cod_raw.endswith('.0'):
            ug_cod_raw = ug_cod_raw[:-2]
        ug_nome = str(row.get('Emitente - UG Nome', '') or '').strip()
        rm, sigla_ug = _resolver_rm_sigla_por_ug(ug_cod_raw, ug_nome, base_oms=base_oms)

        itens.append({
            'ug_cod': ug_cod_raw,
            'ug_nome': ug_nome,
            'rm': rm,
            'sigla_ug': sigla_ug,
            'ne': str(row.get('NE', '') or '').strip(),
            'descricao': str(row.get('NE - Descrição', '') or '').strip()[:200],
            'pi': pi,
            'nd_fmt': nd_fmt,
            'subitem_fmt': subitem_fmt,
            'subitem_nome': str(row.get('NE Item - Subelemento Nome', '') or '').strip(),
            'item_num': str(row.get('NE Item Item', '') or '').strip().rstrip('.0'),
            'qtde_item': str(row.get('NE Item - Qtde Operação', '') or '').strip().rstrip('.0'),
            'item_desc': str(row.get(item_desc_col, '') or '').strip()[:200],
            'valor': _format_valor_br(row.get('NE Item - Valor', '')),
            'conformidade': conf,
        })
    return itens


@app.route('/auditoria-ne', methods=['GET', 'POST'])
@login_required
def auditoria_ne():
    """Auditoria de conformidade das Notas de Empenho com as normas de enquadramento."""
    ne_cache_path = os.path.join(app.instance_path, 'ne_auditoria_cache.json')

    if request.method == 'POST':
        if current_user.nivel_acesso != 'admin':
            flash('Apenas administradores podem realizar o upload de NE.', 'error')
            return redirect(url_for('auditoria_ne'))
        ne_file = request.files.get('ne_file')
        if not ne_file or not ne_file.filename:
            flash('Nenhum arquivo selecionado.', 'warning')
            return redirect(url_for('auditoria_ne'))
        try:
            itens = _processar_auditoria_ne(ne_file)
            timestamp = datetime.now().strftime('%d/%m/%Y %H:%M')
            with open(ne_cache_path, 'w', encoding='utf-8') as f:
                json.dump({'itens': itens, 'timestamp': timestamp}, f, ensure_ascii=False)
            flash(
                f'Planilha NE processada com sucesso. '
                f'{len(itens)} itens analisados.',
                'success'
            )
        except Exception as e:
            flash(f'Erro ao processar planilha NE: {str(e)}', 'error')
        return redirect(url_for('auditoria_ne'))

    # GET – carrega cache se existir
    resultado = None
    resumo = None
    erro = None
    timestamp = None
    if os.path.exists(ne_cache_path):
        try:
            with open(ne_cache_path, encoding='utf-8') as f:
                cache = json.load(f)
            resultado = cache.get('itens', [])
            timestamp = cache.get('timestamp')
            resultado, mudou = _enriquecer_itens_auditoria_ne(resultado)
            if mudou:
                with open(ne_cache_path, 'w', encoding='utf-8') as f:
                    json.dump({'itens': resultado, 'timestamp': timestamp}, f, ensure_ascii=False)
        except Exception as e:
            erro = f'Erro ao carregar cache da auditoria NE: {str(e)}'

    if resultado:
        total = len(resultado)
        conforme = sum(1 for r in resultado if r['conformidade'] == 'Conforme')
        nao_conforme = sum(1 for r in resultado if r['conformidade'] == 'Não Conforme')
        pi_nao_mapeado = sum(1 for r in resultado if r['conformidade'] == 'PI Não Mapeado')
        resumo = {
            'total': total,
            'conforme': conforme,
            'nao_conforme': nao_conforme,
            'pi_nao_mapeado': pi_nao_mapeado,
            'pct_conforme': round(conforme / total * 100, 1) if total else 0,
        }

    return render_template(
        'auditoria_ne.html',
        resultado=resultado,
        resumo=resumo,
        timestamp=timestamp,
        erro=erro,
    )


@app.route('/auditoria-ne/exportar-excel')
@login_required
def exportar_auditoria_ne_excel():
    try:
        itens, timestamp, _ = _carregar_cache_auditoria_ne()
        if not itens:
            flash('Nenhum item analisado disponível para exportação.', 'warning')
            return redirect(url_for('auditoria_ne'))

        itens_filtrados = _filtrar_itens_auditoria_ne(
            itens,
            filtro_texto=request.args.get('filter_text', ''),
            filtro_conformidade=request.args.get('filter_conformidade', ''),
            filtro_pi=request.args.get('filter_pi', ''),
        )
        if not itens_filtrados:
            flash('Nenhum item corresponde aos filtros informados para exportação.', 'warning')
            return redirect(url_for('auditoria_ne'))

        return _exportar_auditoria_ne_excel(itens_filtrados)
    except Exception as e:
        flash(f'Erro ao exportar Excel da auditoria NE: {str(e)}', 'error')
        return redirect(url_for('auditoria_ne'))


@app.route('/auditoria-ne/exportar-pdf')
@login_required
def exportar_auditoria_ne_pdf():
    try:
        itens, timestamp, _ = _carregar_cache_auditoria_ne()
        if not itens:
            flash('Nenhum item analisado disponível para exportação.', 'warning')
            return redirect(url_for('auditoria_ne'))

        itens_filtrados = _filtrar_itens_auditoria_ne(
            itens,
            filtro_texto=request.args.get('filter_text', ''),
            filtro_conformidade=request.args.get('filter_conformidade', ''),
            filtro_pi=request.args.get('filter_pi', ''),
        )
        if not itens_filtrados:
            flash('Nenhum item corresponde aos filtros informados para exportação.', 'warning')
            return redirect(url_for('auditoria_ne'))

        return _exportar_auditoria_ne_pdf(itens_filtrados, timestamp=timestamp)
    except Exception as e:
        flash(f'Erro ao exportar PDF da auditoria NE: {str(e)}', 'error')
        return redirect(url_for('auditoria_ne'))


@app.route('/auditoria-ne/upload-enquadramento', methods=['POST'])
@login_required
def upload_enquadramento_ne():
    """Permite ao admin atualizar a tabela de Enquadramento Despesas."""
    if current_user.nivel_acesso != 'admin':
        flash('Apenas administradores podem atualizar o enquadramento.', 'error')
        return redirect(url_for('auditoria_ne'))
    enq_file = request.files.get('enquadramento_file')
    if not enq_file or not enq_file.filename:
        flash('Nenhum arquivo selecionado.', 'warning')
        return redirect(url_for('auditoria_ne'))
    dest = os.path.join(app.instance_path, 'enquadramento_despesas.xlsx')
    enq_file.save(dest)
    flash('Tabela de Enquadramento Despesas atualizada com sucesso.', 'success')
    return redirect(url_for('auditoria_ne'))


if __name__ == '__main__':
    # Inicializa o banco de dados
    init_database()
    # Executa o app permitindo acesso externo (sem reloader para evitar dupla importação)
    app.run(debug=app.config.get('DEBUG', False), host='0.0.0.0', port=5000, use_reloader=False)
