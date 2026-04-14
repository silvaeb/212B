from models import Pro, SolicitacaoExtraPDRLOG, PedidoSolicitacao, ItemPedido
from database import db
from app import app

with app.app_context():
    pros = Pro.query.all()
    for pro in pros:
        solicitacao = SolicitacaoExtraPDRLOG.query.filter_by(id=pro.solicitacao_pdrlog_id).first()
        if not solicitacao:
            continue
        valor_total = sum(
            item.valor_aprovado for pedido in solicitacao.pedidos for item in pedido.itens
        )
        pro.valor_total = valor_total
        # Se já houver NCs, desconta o valor já convertido
        valor_convertido = sum(nc.valor for nc in pro.notas_credito if nc.status != 'Cancelada')
        pro.valor_restante = max(valor_total - valor_convertido, 0)
    db.session.commit()
print('Valores de PROs atualizados com sucesso!')
