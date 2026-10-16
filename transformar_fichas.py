"""
Ferramenta de Transformação de Fichas de Clientes (CLI)
======================================================
Utiliza o motor inteligente do Transformador de Fichas para ler
qualquer formato de Excel (.xlsx, .xls, .csv), detetar automaticamente
colunas (CRM MEO delimitado por ';', listas com colunas de Morada/CP/Telefone, etc.),
extrair todos os dados fidedignos e exportar para Excel estruturado e CSV CallHub.

Utilização:
    python transformar_fichas.py <ficheiro_entrada.xlsx> [ficheiro_saida.xlsx]
"""

import sys
import os

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

import pandas as pd
import openpyxl

# Importar o motor completo e robusto do app.py
from app import (
    processar_folha,
    construir_tabela,
    remover_duplicados,
    gravar_excel,
    gravar_csv,
    ler_ficheiro,
    _limpar_colunas_vazias
)


def transformar_ficheiro(ficheiro_entrada, ficheiro_saida=None, prefixo351=False):
    """
    Lê o ficheiro de origem e gera o ficheiro Excel estruturado e CSV CallHub.
    """
    if not os.path.exists(ficheiro_entrada):
        raise FileNotFoundError(f"Ficheiro de entrada não encontrado: {ficheiro_entrada}")

    pasta = os.path.dirname(os.path.abspath(ficheiro_entrada))
    nome_base = os.path.splitext(os.path.basename(ficheiro_entrada))[0]

    if ficheiro_saida is None:
        ficheiro_saida = os.path.join(pasta, f"{nome_base}_organizado.xlsx")
    
    ficheiro_csv = os.path.splitext(ficheiro_saida)[0] + "_CallHub.csv"

    print(f"\n[+] A processar ficheiro: {ficheiro_entrada}")
    
    # Suporta ficheiros com múltiplas folhas
    motor = 'openpyxl' if ficheiro_entrada.lower().endswith('.xlsx') else None
    dict_folhas = pd.read_excel(ficheiro_entrada, sheet_name=None, header=None, engine=motor)
    
    todos_registos = []
    total_linhas_lidas = 0

    for nome_folha, df in dict_folhas.items():
        if df.empty:
            continue
        print(f"  Folha '{nome_folha}': {len(df)} linhas")
        regs, n_lidas = processar_folha(df, f"{os.path.basename(ficheiro_entrada)}/{nome_folha}")
        todos_registos.extend(regs)
        total_linhas_lidas += n_lidas

    if not todos_registos:
        print("  ⚠️ Nenhum dado de cliente encontrado!")
        return None

    # Construir tabela e deduplicar de forma 100% segura
    df_tabela = construir_tabela(todos_registos, prefixo351)
    df_final, dups = remover_duplicados(df_tabela)

    tem_tel = df_final['_n_tel'] > 0
    multi_tel = int((df_final['_n_tel'] > 1).sum())
    df_sem_ntel = df_final.drop(columns=['_n_tel'])
    df_com = _limpar_colunas_vazias(df_sem_ntel[tem_tel.values].copy(), obrigatorias=('Nome Completo', 'Contacto 1'))
    df_sem = _limpar_colunas_vazias(df_sem_ntel[~tem_tel.values].copy(), obrigatorias=('Nome Completo', 'Contacto 1'))

    # Gravar Excel com folhas 'Clientes' e 'Sem Contacto'
    gravar_excel(ficheiro_saida, df_com, df_sem)
    # Gravar CSV limpo para CallHub
    gravar_csv(ficheiro_csv, df_com)

    print(f"\n  ✅ Concluído:")
    print(f"    Linhas lidas: {total_linhas_lidas}")
    print(f"    Clientes com contacto: {len(df_com)} ({multi_tel} com 2+ contactos separados por colunas)")
    print(f"    Clientes sem contacto: {len(df_sem)}")
    print(f"    Duplicados consolidados: {dups}")
    print(f"    Ficheiro Excel: {ficheiro_saida}")
    print(f"    Ficheiro CSV CallHub: {ficheiro_csv}")

    return ficheiro_saida


if __name__ == '__main__':
    entrada = sys.argv[1] if len(sys.argv) > 1 else 'alcino.xlsx'
    saida = sys.argv[2] if len(sys.argv) > 2 else None
    prefixo = '--prefixo' in sys.argv or '-p' in sys.argv
    transformar_ficheiro(entrada, saida, prefixo351=prefixo)
