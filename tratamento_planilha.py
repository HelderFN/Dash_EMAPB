import os
import pandas as pd
import matplotlib.pyplot as plt

import warnings

# Suprimir avisos específicos do openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def limpar_df(path):
    """
    Limpa o DataFrame extraindo as informações relevantes e ajustando o formato.
    """
    df = pd.read_excel(path, skiprows=12, engine="openpyxl", index_col=0)

    # Remover últimas duas linhas
    df = df.drop(df.index[-2:])

    # Manter apenas as primeiras 12 colunas
    df.drop(columns=df.columns[12:], inplace=True)

    # Renomear colunas
    colunas_rename = ["nome", "Notas_1bim", "Faltas_1bim", "Notas_2bim", "Faltas_2bim",
                      "Notas_3bim", "Faltas_3bim", "Notas_4bim", "Faltas_4bim", "Rec", "Media", "Faltas"]
    df.columns = colunas_rename

    # Remover colunas desnecessárias
    df.drop(columns=["Rec", "Media", "Faltas"], inplace=True)

    # Transformar em formato longo
    df_long = df.melt(
        id_vars=["nome"],
        var_name="tipo",
        value_name="valor"
    )

    # Separar informações de 'tipo'
    df_long[['variavel', 'bimestre']
            ] = df_long['tipo'].str.split("_", expand=True)

    # Converter bimestre para número inteiro
    df_long['bimestre'] = df_long['bimestre'].str.extract(r'(\d+)').astype(int)

    # Pivotar de volta ao formato desejado
    df_final = df_long.pivot_table(
        index=["nome", "bimestre"],
        columns="variavel",
        values="valor",
        aggfunc="first").reset_index()

    # Limpar nomes de colunas e reordenar
    df_final.columns.name = None
    df_final = df_final.rename(columns={'nota': 'notas', 'faltas': 'faltas'})
    df_final = df_final.sort_values(
        by=['bimestre', 'nome']).reset_index(drop=True)

    return df_final


def processar_planilhas(caminhos):
    """
    Processa uma lista de caminhos de planilhas, adicionando a coluna 'Turma'.
    Retorna uma lista de DataFrames limpos.
    """
    lista_dfs = []

    for caminho in caminhos:
        if not os.path.exists(caminho):
            print(f"Arquivo não encontrado: {caminho}")
            continue

        # Extrai número da turma a partir do nome do arquivo
        turma = int(caminho.split("_")[-1].split(".")[0])
        df = limpar_df(caminho)
        df['turma'] = turma  # Adiciona a coluna Turma
        lista_dfs.append(df)

    return lista_dfs


def mesclar_dfs(lista_dfs):
    """
    Mescla uma lista de DataFrames em um único DataFrame.
    """
    return pd.concat(lista_dfs, ignore_index=True)


if __name__ == "__main__":
    # Caminhos para as planilhas
    caminho_planilhas = ["dados_turmas/Relatório de Rendimento Escolar_900.xlsx",
                         "dados_turmas/Relatório de Rendimento Escolar_901.xlsx",
                         "dados_turmas/Relatório de Rendimento Escolar_902.xlsx"]

    # Processar as planilhas e obter DataFrames limpos
    lista_dfs = processar_planilhas(caminho_planilhas)

    # Mesclar os DataFrames
    df_total = mesclar_dfs(lista_dfs)

    # Exportar para Excel
    planilha_final = "data_frame_total.xlsx"
    df_total.to_excel(planilha_final, index=False)

    print(f"\nDataFrame mesclado salvo com sucesso em '{planilha_final}'!\n")
