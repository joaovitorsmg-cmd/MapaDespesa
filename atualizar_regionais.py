# -*- coding: utf-8 -*-
# atualizar_regionais.py
# Aplica alterações pontuais no Base_regionais.xlsx sem precisar abrir o Excel.
# Execute: python atualizar_regionais.py

import pandas as pd, os, shutil
from datetime import datetime

ARQUIVO = 'Base_regionais.xlsx'

# ── Alterações a aplicar ─────────────────────────────────────────────────────
# Cada item: { 'FILIAL': <id>, coluna: novo_valor, ... }
ALTERACOES = [
    # Filial 37 – CANARANA/MT: Sergio Adriano assumiu como Regional
    {
        'FILIAL': 37,
        'SIGLA':       'CAN',
        'LOJA':        'CANARANA',
        'ESTADO':      'MT',
        'GERENTE':     'LEONARDO FREITAS',
        'SUPERVISOR':  'INGRID KAROLAYNE',
        'REGIONAL':    'SERGIO ADRIANO',
        'AUDITOR':     'JOÃO VITOR MENDES',
    },
    # Filial 38 – GURUPI/TO: Aldai Lemos assumiu como Gerente
    {
        'FILIAL': 38,
        'SIGLA':       'GUR',
        'LOJA':        'GURUPI',
        'ESTADO':      'TO',
        'GERENTE':     'ALDAI LEMOS',
        'SUPERVISOR':  'RENATA CERQUEIRA',
        'REGIONAL':    'SERGIO ADRIANO',
        'AUDITOR':     'FÁBIO RIBEIRO',
    },
]
# ─────────────────────────────────────────────────────────────────────────────


def main():
    if not os.path.exists(ARQUIVO):
        print(f"❌  '{ARQUIVO}' não encontrado nesta pasta.")
        return

    # Backup antes de alterar
    bk = f"{ARQUIVO}.bak_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    shutil.copy2(ARQUIVO, bk)
    print(f"📦  Backup criado: {bk}")

    df = pd.read_excel(ARQUIVO)

    colunas_arquivo = set(df.columns.str.strip())
    for alt in ALTERACOES:
        filial = alt['FILIAL']
        mask = df['FILIAL'] == filial
        if not mask.any():
            print(f"⚠️   Filial {filial} não encontrada — linha ignorada.")
            continue

        idx = df.index[mask][0]
        mudancas = []
        for col, novo in alt.items():
            if col == 'FILIAL':
                continue
            if col not in colunas_arquivo:
                print(f"   ℹ️  Coluna '{col}' não existe no arquivo — ignorado.")
                continue
            antigo = df.at[idx, col]
            if str(antigo).strip() != str(novo).strip():
                df.at[idx, col] = novo
                mudancas.append(f"    {col}: '{antigo}' → '{novo}'")

        if mudancas:
            print(f"✅  Filial {filial} ({alt.get('LOJA','')}) atualizada:")
            print('\n'.join(mudancas))
        else:
            print(f"ℹ️   Filial {filial}: sem diferenças detectadas (já estava atualizada).")

    df.to_excel(ARQUIVO, index=False)
    print(f"\n💾  '{ARQUIVO}' salvo com as alterações.")
    print("▶️   Agora execute: python gerar_json.py")


if __name__ == '__main__':
    main()
    input("\nPressione ENTER para fechar...")
