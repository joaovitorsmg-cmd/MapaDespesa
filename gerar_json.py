# -*- coding: utf-8 -*-
# gerar_json.py — Agroquima Painel Orcamentario
# Execute: python gerar_json.py  (ou de 2 cliques neste arquivo no Windows)
#
# Coloque na mesma pasta:
#   - Mapa_Despesa.xlsx   (exportado do sistema)
#   - Base_regionais.xlsx (mapeamento filial → LOJA/REGIONAL/AUDITOR/SIGLA)
#   - gerar_json.py       (este arquivo)
#
# O script gera aqm_data.json — arraste esse arquivo no painel (aba Atualizar Base).
import pandas as pd, json, os, glob

DESPESA_FILE   = 'Mapa_Despesa.xlsx'
REGIONAIS_FILE = 'Base_regionais.xlsx'
OUT            = 'aqm_data.json'


def encontrar_arquivo(nome_esperado, pistas):
    if os.path.exists(nome_esperado):
        return nome_esperado
    candidatos = [f for f in glob.glob('*.xlsx') if not os.path.basename(f).startswith('~$')]
    for f in candidatos:
        chave = f.lower().replace(' ', '').replace('_', '')
        if all(p in chave for p in pistas):
            print(f"ℹ️  '{nome_esperado}' não encontrado — usando '{f}' no lugar.")
            return f
    print(f"\n❌ Não encontrei nenhum arquivo parecido com '{nome_esperado}' nesta pasta.")
    if candidatos:
        print("   Arquivos .xlsx encontrados aqui:")
        for f in candidatos:
            print(f"     - {f}")
    else:
        print("   Nenhum arquivo .xlsx encontrado nesta pasta.")
    print(f"   Coloque o arquivo certo aqui, ou renomeie para '{nome_esperado}'.")
    raise FileNotFoundError(nome_esperado)


def main():
    despesa_path   = encontrar_arquivo(DESPESA_FILE,   ['despesa'])
    regionais_path = encontrar_arquivo(REGIONAIS_FILE, ['regio'])

    d = pd.read_excel(despesa_path)
    r = pd.read_excel(regionais_path)

    # ==== PERÍODO: usar SEMPRE ANO/MES/DIA, NUNCA DATA_LANCAMENTO ====
    # DATA_LANCAMENTO/DATA_APROVACAO são a data em que a contabilidade fechou o
    # lote — não representam a competência. A competência real é ANO e MES.
    # ANO_REFERENCIA: None = usa automaticamente o ano mais recente na base.
    ANO_REFERENCIA = None
    ano_alvo = ANO_REFERENCIA if ANO_REFERENCIA else int(d['ANO'].max())
    print(f"📅 Ano de referência: {ano_alvo}")
    d = d[d['ANO'] == ano_alvo].copy()

    # ==== CRUZAMENTO COM BASE REGIONAIS ====
    fl  = dict(zip(r['FILIAL'], r['LOJA']))
    fr  = dict(zip(r['FILIAL'], r['REGIONAL']))
    fa  = dict(zip(r['FILIAL'], r['AUDITOR']))
    fs  = dict(zip(r['FILIAL'], r['SIGLA']))
    fg  = dict(zip(r['FILIAL'], r['GERENTE']))
    fsu = dict(zip(r['FILIAL'], r['SUPERVISOR']))
    fe  = dict(zip(r['FILIAL'], r['ESTADO']))

    d['LOJA']      = d['ID_FILIAL'].map(fl)
    d['REGIONAL']  = d['ID_FILIAL'].map(fr)
    d['AUDITOR']   = d['ID_FILIAL'].map(fa)
    d['SIGLA']     = d['ID_FILIAL'].map(fs).fillna(d['ID_FILIAL'].apply(lambda x: f'F{x}'))
    d['GERENTE']   = d['ID_FILIAL'].map(fg)
    d['SUPERVISOR']= d['ID_FILIAL'].map(fsu)
    d['ESTADO']    = d['ID_FILIAL'].map(fe)
    for col in ['LOJA', 'REGIONAL', 'AUDITOR', 'GERENTE', 'SUPERVISOR', 'ESTADO']:
        d[col] = d[col].fillna('N/D')

    # ==== CLASSIFICAÇÃO: FATURAMENTO x DESPESA ====
    # Faturamento = SOMA de VLR_REALIZADO dos GRUPOS_RECEITA abaixo.
    # Tudo fora dessa lista entra como despesa.
    GRUPOS_RECEITA = [
        '01-VENDAS DE MERCADORIAS',
    ]

    def _norm(s):
        return str(s).strip().upper()

    grupos_receita_norm = {_norm(g) for g in GRUPOS_RECEITA}
    d['_GC_NORM'] = d['GRUPO_CONTA'].apply(_norm)

    fat  = d[ d['_GC_NORM'].isin(grupos_receita_norm)].copy()
    desp = d[~d['_GC_NORM'].isin(grupos_receita_norm)].copy()

    _resumo = d.groupby('GRUPO_CONTA')['VLR_REALIZADO'].agg(['count', 'sum']).reset_index()
    _resumo['CLASSIFICACAO'] = _resumo['GRUPO_CONTA'].apply(
        lambda g: 'RECEITA' if _norm(g) in grupos_receita_norm else 'DESPESA')
    print("\n📋 Classificação por GRUPO_CONTA:")
    for _, row in _resumo.sort_values('sum', ascending=False).iterrows():
        print(f"  [{row['CLASSIFICACAO']:8s}] {row['GRUPO_CONTA']:45s} | {int(row['count']):6d} lanç. | R$ {row['sum']:,.2f}")
    print()

    if fat.empty:
        raise RuntimeError(
            "Nenhum lançamento casou com GRUPOS_RECEITA — faturamento ficaria zerado. "
            "Confira os nomes exatos em GRUPO_CONTA na planilha.")

    mes_nomes = {1:'Jan',2:'Fev',3:'Mar',4:'Abr',5:'Mai',6:'Jun',
                 7:'Jul',8:'Ago',9:'Set',10:'Out',11:'Nov',12:'Dez'}

    meses_com_realizado = sorted(fat[fat['VLR_REALIZADO'] > 0]['MES'].unique())
    if meses_com_realizado:
        periodo_str = f"{mes_nomes[min(meses_com_realizado)]}-{mes_nomes[max(meses_com_realizado)]} {ano_alvo}"
    else:
        periodo_str = str(ano_alvo)

    def pct(numer, denom):
        return (numer / denom * 100).replace([float('inf'), float('-inf')], 0).fillna(0).round(2)

    def pct_s(n, d):
        return round(n / d * 100, 2) if d else 0.0

    # Janela comparável: apenas meses com realizado (evita desvio inflado vs orçamento anual)
    fat_comp  = fat[ fat['MES'].isin(meses_com_realizado)].copy()
    desp_comp = desp[desp['MES'].isin(meses_com_realizado)].copy()

    fat_fil  = fat_comp.groupby(
        ['ID_FILIAL','SIGLA','LOJA','REGIONAL','AUDITOR','GERENTE','SUPERVISOR','ESTADO']
    ).agg(FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_fil = desp_comp.groupby('ID_FILIAL').agg(
        DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    filiais = fat_fil.merge(desp_fil, on='ID_FILIAL', how='left').fillna(0)
    filiais['PERC_DESP']   = pct(filiais['DESP_REAL'], filiais['FAT_REAL'])
    filiais['DESVIO_FAT']  = pct(filiais['FAT_REAL'] - filiais['FAT_PREV'],  filiais['FAT_PREV'])
    filiais['DESVIO_DESP'] = pct(filiais['DESP_REAL'] - filiais['DESP_PREV'], filiais['DESP_PREV'])
    filiais = filiais.sort_values('FAT_REAL', ascending=False)

    fat_mes  = fat.groupby('MES').agg(FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_mes = desp.groupby('MES').agg(DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    meses = fat_mes.merge(desp_mes, on='MES').fillna(0)
    meses['NOME']     = meses['MES'].map(mes_nomes)
    meses['PERC_DESP']= pct(meses['DESP_REAL'], meses['FAT_REAL'])

    fat_fil_mes  = fat.groupby(['ID_FILIAL','MES']).agg(
        FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_fil_mes = desp.groupby(['ID_FILIAL','MES']).agg(
        DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    filiais_mes = fat_fil_mes.merge(desp_fil_mes, on=['ID_FILIAL','MES'], how='outer').fillna(0)
    filiais_mes['ID_FILIAL'] = filiais_mes['ID_FILIAL'].astype(int)
    filiais_mes['MES']       = filiais_mes['MES'].astype(int)

    fat_reg  = fat_comp.groupby('REGIONAL').agg(FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_reg = desp_comp.groupby('REGIONAL').agg(DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    regionais_out = fat_reg.merge(desp_reg, on='REGIONAL').fillna(0)
    regionais_out['PERC_DESP'] = pct(regionais_out['DESP_REAL'], regionais_out['FAT_REAL'])

    fat_aud  = fat_comp.groupby('AUDITOR').agg(FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_aud = desp_comp.groupby('AUDITOR').agg(DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    auditores = fat_aud.merge(desp_aud, on='AUDITOR').fillna(0)
    auditores['PERC_DESP'] = pct(auditores['DESP_REAL'], auditores['FAT_REAL'])

    grupos = desp_comp.groupby('GRUPO_CONTA').agg(
        DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')
    ).reset_index().sort_values('DESP_REAL', ascending=False)
    grupos['DESVIO'] = pct(grupos['DESP_REAL'] - grupos['DESP_PREV'], grupos['DESP_PREV'])

    grupos_filial = desp_comp.groupby(['ID_FILIAL', 'GRUPO_CONTA']).agg(
        DESP_REAL=('VLR_REALIZADO', 'sum'),
        DESP_PREV=('VLR_PREVISAO', 'sum')
    ).reset_index()
    grupos_filial['ID_FILIAL'] = grupos_filial['ID_FILIAL'].astype(int)

    tot_fat_real  = fat_comp['VLR_REALIZADO'].sum()
    tot_fat_prev  = fat_comp['VLR_PREVISAO'].sum()
    tot_desp_real = desp_comp['VLR_REALIZADO'].sum()
    tot_desp_prev = desp_comp['VLR_PREVISAO'].sum()
    tot_fat_prev_ano  = fat['VLR_PREVISAO'].sum()
    tot_desp_prev_ano = desp['VLR_PREVISAO'].sum()

    payload = {
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
        'filiais':     filiais.to_dict('records'),
        'meses':       meses.to_dict('records'),
        'filiais_mes': filiais_mes.to_dict('records'),
        'regionais':   regionais_out.to_dict('records'),
        'auditores':   auditores.to_dict('records'),
        'grupos':        grupos.to_dict('records'),
        'grupos_filial': grupos_filial.to_dict('records'),
    }

    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, default=str)

    print(f"✅ {OUT} gerado com sucesso!")
    print(f"✅ {len(filiais)} filiais | Período: {periodo_str}")
    print(f"💰 Fat: R$ {tot_fat_real:,.0f} | Desp: R$ {tot_desp_real:,.0f} | Resultado: R$ {tot_fat_real - tot_desp_real:,.0f}")
    print(f"\n🎉 Pronto! Arraste o aqm_data.json no painel (aba Atualizar Base).")


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print("\n" + "=" * 60)
        print("❌ ERRO ao gerar o aqm_data.json:")
        print(f"   {type(e).__name__}: {e}")
        print("=" * 60)
    input("\nPressione ENTER para fechar esta janela...")
