# -*- coding: utf-8 -*-
# gerar_json.py -- Agroquima Painel Orcamentario
# Execute: python gerar_json.py  (ou de 2 cliques neste arquivo)
import pandas as pd, json, os, glob

DESPESA_FILE   = 'Mapa_Despesa.xlsx'
REGIONAIS_FILE = 'Base_regionais.xlsx'
OUT            = 'aqm_data.json'

# ANO_REFERENCIA:
#   None  -> exporta TODOS os anos encontrados na planilha (ex.: 2025 e 2026)
#   2026  -> exporta SOMENTE o ano informado
ANO_REFERENCIA = None

# Filiais e regionais excluídas de TODOS os cálculos (filiais sem operação,
# sem mapeamento ou pertencentes a estruturas não comparáveis).
SIGLAS_EXCLUIR   = {'F2', 'FAP', 'F18', 'F33', 'GRA'}
REGIONAIS_EXCLUIR = {'ANNAPAULLA GARCIA', 'EVANDRO MACEDO', 'HENISLEY SABINO', 'N/D'}


def encontrar_arquivo(nome_esperado, pistas):
    if os.path.exists(nome_esperado):
        return nome_esperado
    candidatos = [f for f in glob.glob('*.xlsx') if not os.path.basename(f).startswith('~$')]
    for f in candidatos:
        chave = f.lower().replace(' ', '').replace('_', '')
        if all(p in chave for p in pistas):
            print(f"ℹ️  '{nome_esperado}' não encontrado — usando '{f}' no lugar (nome parecido).")
            return f

    print(f"""
❌ Não encontrei nenhum arquivo parecido com '{nome_esperado}' nesta pasta.""")
    if candidatos:
        print("   Arquivos .xlsx encontrados aqui:")
        for f in candidatos:
            print(f"     - {f}")
    else:
        print("   Nenhum arquivo .xlsx encontrado nesta pasta.")
    print(f"   Coloque o arquivo certo aqui, ou renomeie para '{nome_esperado}'.")
    raise FileNotFoundError(nome_esperado)


def processar_ano(d_ano, r, ano_alvo):
    """Recebe as linhas já filtradas para 'ano_alvo' e retorna o payload desse ano."""

    mes_nomes = {1:'Jan',2:'Fev',3:'Mar',4:'Abr',5:'Mai',6:'Jun',
                 7:'Jul',8:'Ago',9:'Set',10:'Out',11:'Nov',12:'Dez'}

    GRUPOS_RECEITA = ['01-VENDAS DE MERCADORIAS']

    def _norm(s):
        return str(s).strip().upper()

    grupos_receita_norm = {_norm(g) for g in GRUPOS_RECEITA}

    # Exclui filiais sem operação/mapeamento antes de qualquer cálculo
    sig_excl = {s.upper() for s in SIGLAS_EXCLUIR}
    reg_excl = {x.upper() for x in REGIONAIS_EXCLUIR}
    d_ano = d_ano.copy()
    d_ano = d_ano[
        ~(d_ano['SIGLA'].apply(lambda x: str(x).strip().upper()).isin(sig_excl) |
          d_ano['REGIONAL'].apply(lambda x: str(x).strip().upper()).isin(reg_excl))
    ]

    d_ano['_GC_NORM'] = d_ano['GRUPO_CONTA'].apply(_norm)

    fat  = d_ano[ d_ano['_GC_NORM'].isin(grupos_receita_norm)].copy()
    desp = d_ano[~d_ano['_GC_NORM'].isin(grupos_receita_norm)].copy()

    if fat.empty:
        raise RuntimeError(
            f"Ano {ano_alvo}: nenhum lançamento de '01-VENDAS DE MERCADORIAS' encontrado.")

    # Meses com faturamento realizado > 0 (janela de comparação)
    meses_com_realizado = sorted(fat[fat['VLR_REALIZADO'] > 0]['MES'].unique())
    if meses_com_realizado:
        periodo_str = f"{mes_nomes[min(meses_com_realizado)]}-{mes_nomes[max(meses_com_realizado)]} {ano_alvo}"
    else:
        periodo_str = str(ano_alvo)

    def pct(numer, denom):
        return (numer / denom * 100).replace([float('inf'), float('-inf')], 0).fillna(0).round(2)

    def pct_s(n, d_val):
        return round(n / d_val * 100, 2) if d_val else 0.0

    fat_comp  = fat[ fat['MES'].isin(meses_com_realizado)].copy()
    desp_comp = desp[desp['MES'].isin(meses_com_realizado)].copy()

    # Filiais (período comparável)
    fat_fil = fat_comp.groupby(
        ['ID_FILIAL','SIGLA','LOJA','REGIONAL','AUDITOR','GERENTE','SUPERVISOR','ESTADO']
    ).agg(FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_fil = desp_comp.groupby('ID_FILIAL').agg(
        DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    filiais = fat_fil.merge(desp_fil, on='ID_FILIAL', how='left').fillna(0)
    filiais['PERC_DESP']   = pct(filiais['DESP_REAL'], filiais['FAT_REAL'])
    filiais['DESVIO_FAT']  = pct(filiais['FAT_REAL']  - filiais['FAT_PREV'],  filiais['FAT_PREV'])
    filiais['DESVIO_DESP'] = pct(filiais['DESP_REAL'] - filiais['DESP_PREV'], filiais['DESP_PREV'])
    filiais = filiais.sort_values('FAT_REAL', ascending=False)

    # Meses (ano inteiro — para gráfico de sazonalidade)
    fat_mes  = fat.groupby('MES').agg(
        FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_mes = desp.groupby('MES').agg(
        DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    meses = fat_mes.merge(desp_mes, on='MES').fillna(0)
    meses['NOME'] = meses['MES'].map(mes_nomes)
    meses['PERC_DESP'] = pct(meses['DESP_REAL'], meses['FAT_REAL'])

    # Filiais × Meses (ano inteiro — para filtro de período na Visão Geral)
    fat_fil_mes  = fat.groupby(['ID_FILIAL','MES']).agg(
        FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_fil_mes = desp.groupby(['ID_FILIAL','MES']).agg(
        DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    filiais_mes = fat_fil_mes.merge(desp_fil_mes, on=['ID_FILIAL','MES'], how='outer').fillna(0)
    filiais_mes['ID_FILIAL'] = filiais_mes['ID_FILIAL'].astype(int)
    filiais_mes['MES']       = filiais_mes['MES'].astype(int)

    # Regionais
    fat_reg  = fat_comp.groupby('REGIONAL').agg(
        FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_reg = desp_comp.groupby('REGIONAL').agg(
        DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    regionais_out = fat_reg.merge(desp_reg, on='REGIONAL').fillna(0)
    regionais_out['PERC_DESP'] = pct(regionais_out['DESP_REAL'], regionais_out['FAT_REAL'])

    # Auditores
    fat_aud  = fat_comp.groupby('AUDITOR').agg(
        FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_aud = desp_comp.groupby('AUDITOR').agg(
        DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    auditores = fat_aud.merge(desp_aud, on='AUDITOR').fillna(0)
    auditores['PERC_DESP'] = pct(auditores['DESP_REAL'], auditores['FAT_REAL'])

    # Grupos de conta
    grupos = (desp_comp.groupby('GRUPO_CONTA')
              .agg(DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum'))
              .reset_index().sort_values('DESP_REAL', ascending=False))
    grupos['DESVIO'] = pct(grupos['DESP_REAL'] - grupos['DESP_PREV'], grupos['DESP_PREV'])

    # Grupos × Filial
    grupos_filial = (desp_comp.groupby(['ID_FILIAL','GRUPO_CONTA'])
                    .agg(DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum'))
                    .reset_index())
    grupos_filial['ID_FILIAL'] = grupos_filial['ID_FILIAL'].astype(int)

    # Totais
    tot_fat_real      = fat_comp['VLR_REALIZADO'].sum()
    tot_fat_prev      = fat_comp['VLR_PREVISAO'].sum()
    tot_desp_real     = desp_comp['VLR_REALIZADO'].sum()
    tot_desp_prev     = desp_comp['VLR_PREVISAO'].sum()
    tot_fat_prev_ano  = fat['VLR_PREVISAO'].sum()
    tot_desp_prev_ano = desp['VLR_PREVISAO'].sum()

    return {
        'meta': {
            'periodo':          periodo_str,
            'ano':              int(ano_alvo),
            'filiais_count':    int(len(filiais)),
            'tot_fat_real':     round(tot_fat_real,  2),
            'tot_fat_prev':     round(tot_fat_prev,  2),
            'tot_desp_real':    round(tot_desp_real, 2),
            'tot_desp_prev':    round(tot_desp_prev, 2),
            'tot_fat_prev_ano': round(tot_fat_prev_ano,  2),
            'tot_desp_prev_ano':round(tot_desp_prev_ano, 2),
            'perc_desp_real':   pct_s(tot_desp_real, tot_fat_real),
            'perc_desp_plan':   pct_s(tot_desp_prev, tot_fat_prev),
            'resultado':        round(tot_fat_real - tot_desp_real, 2),
            'dev_fat':          pct_s(tot_fat_real  - tot_fat_prev,  tot_fat_prev),
            'dev_desp':         pct_s(tot_desp_real - tot_desp_prev, tot_desp_prev),
        },
        'filiais':       filiais.to_dict('records'),
        'meses':         meses.to_dict('records'),
        'filiais_mes':   filiais_mes.to_dict('records'),
        'regionais':     regionais_out.to_dict('records'),
        'auditores':     auditores.to_dict('records'),
        'grupos':        grupos.to_dict('records'),
        'grupos_filial': grupos_filial.to_dict('records'),
    }


def main():
    despesa_path   = encontrar_arquivo(DESPESA_FILE,   ['despesa'])
    regionais_path = encontrar_arquivo(REGIONAIS_FILE, ['regio'])

    d = pd.read_excel(despesa_path)
    r = pd.read_excel(regionais_path)

    # Mapeia dados de regionais para cada filial
    fl  = dict(zip(r['FILIAL'], r['LOJA']))
    fr  = dict(zip(r['FILIAL'], r['REGIONAL']))
    fa  = dict(zip(r['FILIAL'], r['AUDITOR']))
    fs  = dict(zip(r['FILIAL'], r['SIGLA']))
    fg  = dict(zip(r['FILIAL'], r['GERENTE']))
    fsu = dict(zip(r['FILIAL'], r['SUPERVISOR']))
    fe  = dict(zip(r['FILIAL'], r['ESTADO']))

    d['LOJA']       = d['ID_FILIAL'].map(fl)
    d['REGIONAL']   = d['ID_FILIAL'].map(fr)
    d['AUDITOR']    = d['ID_FILIAL'].map(fa)
    d['SIGLA']      = d['ID_FILIAL'].map(fs).fillna(d['ID_FILIAL'].apply(lambda x: f'F{x}'))
    d['GERENTE']    = d['ID_FILIAL'].map(fg)
    d['SUPERVISOR'] = d['ID_FILIAL'].map(fsu)
    d['ESTADO']     = d['ID_FILIAL'].map(fe)
    for col in ['LOJA','REGIONAL','AUDITOR','GERENTE','SUPERVISOR','ESTADO']:
        d[col] = d[col].fillna('N/D')

    # Detecta todos os anos presentes na planilha
    todos_anos = sorted(d['ANO'].dropna().unique().astype(int))
    if not todos_anos:
        raise RuntimeError("Coluna ANO não encontrada ou sem valores válidos.")

    if ANO_REFERENCIA:
        anos_processar = [int(ANO_REFERENCIA)]
    else:
        anos_processar = todos_anos

    print(f"\n📅 Anos encontrados na planilha: {todos_anos}")
    print(f"📅 Anos que serão exportados:    {anos_processar}")
    if len(todos_anos) > 1 and ANO_REFERENCIA is None:
        print("   (Para exportar apenas um ano, defina ANO_REFERENCIA no topo do script)")

    # Diagnóstico de classificação — usa todos os dados (não filtra por ano)
    def _norm(s):
        return str(s).strip().upper()
    GRUPOS_RECEITA = ['01-VENDAS DE MERCADORIAS']
    grupos_receita_norm = {_norm(g) for g in GRUPOS_RECEITA}
    d['_GC_NORM'] = d['GRUPO_CONTA'].apply(_norm)
    _resumo = d.groupby('GRUPO_CONTA')['VLR_REALIZADO'].agg(['count','sum']).reset_index()
    _resumo['CLASSIFICACAO'] = _resumo['GRUPO_CONTA'].apply(
        lambda g: 'RECEITA' if _norm(g) in grupos_receita_norm else 'DESPESA')
    print("\n📋 Classificação por GRUPO_CONTA (todos os anos, confira se está correto):")
    for _, row in _resumo.sort_values('sum', ascending=False).iterrows():
        print(f"  [{row['CLASSIFICACAO']:8s}] {row['GRUPO_CONTA']:45s} | {int(row['count']):6d} lanç. | R$ {row['sum']:,.2f}")
    print()

    # Processa cada ano separadamente
    anos_data = {}
    for ano in anos_processar:
        print(f"⏳ Processando {ano}...")
        d_ano = d[d['ANO'] == ano].copy()
        dados = processar_ano(d_ano, r, ano)
        anos_data[str(ano)] = dados
        m = dados['meta']
        print(f"   ✅ {m['filiais_count']} filiais | {m['periodo']}")
        print(f"   💰 Fat: R$ {m['tot_fat_real']:,.0f} | Desp: R$ {m['tot_desp_real']:,.0f} | Resultado: R$ {m['resultado']:,.0f}")

    # Ano ativo = mais recente entre os exportados
    ano_ativo = max(int(a) for a in anos_data.keys())

    # Estrutura do JSON de saída
    # O painel detecta '_multi_ano': true e carrega todos os anos em _dataByYear
    payload = {
        '_multi_ano': True,
        'ano_ativo':  ano_ativo,
        'anos':       anos_data,
        # Campos do ano ativo também na raiz para compatibilidade com versões antigas
        **anos_data[str(ano_ativo)],
    }

    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, default=str)

    print(f"\n✅ {OUT} gerado com sucesso!")
    print(f"✅ Anos incluídos: {', '.join(str(a) for a in sorted(int(k) for k in anos_data.keys()))}")
    print(f"✅ Ano ativo no painel: {ano_ativo}")


if __name__ == '__main__':
    try:
        main()
        print("\n🎉 Concluído! Arraste o aqm_data.json no painel (aba Atualizar Base).")
    except Exception as e:
        print("\n" + "="*60)
        print("❌ ERRO ao gerar o aqm_data.json:")
        print(f"   {type(e).__name__}: {e}")
        print("="*60)
    input("\nPressione ENTER para fechar esta janela...")
