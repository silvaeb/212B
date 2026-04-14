


import pandas as pd
from collections import defaultdict
import re

ARQUIVO = 'instance/auditoria_upload/auditoria_pdr_log_cache.csv'
PI_ALVO = 'E6SUPLJA5PA'
ARQUIVO_SAIDA = 'relatorio_om_nc_pi_nd.xlsx'



# Índices das colunas conforme análise do arquivo
IDX_NC = 0
IDX_CODUG = 1
IDX_OM = 2
IDX_ND = 3
IDX_PI = 5
IDX_DOC = 6  # Coluna G (Doc - Observação)

def main():
    try:
        df = pd.read_csv(ARQUIVO, dtype=str)
    except Exception as e:
        print(f'Erro ao ler o arquivo: {e}')
        return

    # Filtra PI
    df_pi = df[df[str(IDX_PI)].astype(str).str.upper().str.strip() == PI_ALVO]
    if df_pi.empty:
        print(f'Nenhum registro encontrado para PI {PI_ALVO}')
        return

    # Agrupa por OM e ND, contando NCs distintos

    agrupamento = defaultdict(lambda: {'ncs': set(), 'codug': None, 'codom': None, 'pi': None})
    for _, row in df_pi.iterrows():
        om = str(row[str(IDX_OM)]).strip()
        nd = str(row[str(IDX_ND)]).strip()
        nc = str(row[str(IDX_NC)]).strip()
        codug = str(row[str(IDX_CODUG)]).strip()
        doc_obs = str(row[str(IDX_DOC)]).strip()
        pi = str(row[str(IDX_PI)]).strip()
        # Extrai CODOM do Doc - Observação: padrão (NUMERO-HIFEN)
        codom = None
        match = re.search(r'\((\d+)-', doc_obs)
        if match:
            codom = match.group(1).zfill(5)
        if om and nd and nc and codug and codom and pi:
            agrupamento[(om, nd, codug, codom, pi)]['ncs'].add(nc)
            if agrupamento[(om, nd, codug, codom, pi)]['codug'] is None:
                agrupamento[(om, nd, codug, codom, pi)]['codug'] = codug
            if agrupamento[(om, nd, codug, codom, pi)]['codom'] is None:
                agrupamento[(om, nd, codug, codom, pi)]['codom'] = codom
            if agrupamento[(om, nd, codug, codom, pi)]['pi'] is None:
                agrupamento[(om, nd, codug, codom, pi)]['pi'] = pi

    # Monta DataFrame para exportação
    dados_export = []
    for (om, nd, codug, codom, pi), info in agrupamento.items():
        ncs = info['ncs']
        if len(ncs) > 1:
            dados_export.append({
                'OM': om,
                'ND': nd,
                'CODUG': codug,
                'CODOM': codom,
                'PI': pi,
                'Qtd_NCs': len(ncs),
                'NCs': ', '.join(sorted(ncs))
            })

    if dados_export:
        df_export = pd.DataFrame(dados_export)
        df_export.to_excel(ARQUIVO_SAIDA, index=False)
        print(f'Relatório exportado para {ARQUIVO_SAIDA}')
    else:
        print('Nenhuma OM encontrada com mais de uma NC para o PI e ND especificados.')

if __name__ == '__main__':
    main()
