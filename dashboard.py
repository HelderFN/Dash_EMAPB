import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from math import pi

# Carregar os dados
file_path = 'Notas_9ano.xlsx'
data = pd.read_excel(file_path, sheet_name='Notas')

# Configuração do título
st.title("Dashboard de Notas - 9º Ano")

# Gráfico 1: Média por Turma com cores diferentes e valores nas barras
st.header("Média de Notas por Turma")
media_por_turma = data.groupby('Turma')['total'].mean()
fig, ax = plt.subplots()
colors = sns.color_palette("Set2", len(media_por_turma))
bars = media_por_turma.plot(kind='barh', color=colors, ax=ax)
ax.set_title("Média por Turma")
ax.set_ylabel("Média das Notas")
ax.set_xlabel("Turma")
# Adicionar valores no topo das barras
for bar in bars.patches:
    ax.text(bar.get_width(), bar.get_y()+bar.get_height()/2, f'{
            bar.get_width():.2f}', ha='right', va='center')
    # ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5,
    #       f'{bar.get_height():.2f}', ha='center', va='bottom')
ax.legend().remove()  # Remove a legenda para clareza
st.pyplot(fig)

# Gráfico 2: Notas empilhadas por aluno
st.header("Notas por Aluno")
turma_filtrada = st.selectbox(
    "Selecione uma Turma para Analisar", data['Turma'].unique())
dados_turma = data[data['Turma'] == turma_filtrada]
fig, ax = plt.subplots(figsize=(25, 10))
componentes = ['Inst_1', 'Inst_2', 'Inst_3', 'Comportamento', 'CAED']
dados_turma.set_index('Nome')[componentes].plot(
    kind='barh', stacked=True, ax=ax, colormap='Set2')
ax.set_title(f"Notas Empilhadas - Turma {turma_filtrada}")
ax.set_ylabel("Pontuação Total")
ax.set_xlabel("Alunos")
plt.xticks(rotation=45, ha='right')
# Mover a legenda para fora do gráfico
ax.legend(loc='upper left', bbox_to_anchor=(1, 1))
st.pyplot(fig)

# Gráfico 3: Distribuição Geral das Notas com filtro opcional por turma
st.header("Distribuição Geral das Notas")
turma_opcional = st.selectbox("Filtrar por Turma (opcional)", [
                              'Todas'] + list(data['Turma'].unique()))
if turma_opcional == 'Todas':
    notas_filtradas = data['total']
else:
    notas_filtradas = data[data['Turma'] == turma_opcional]['total']
fig, ax = plt.subplots()
sns.histplot(notas_filtradas, bins=10, kde=False, color='orange', ax=ax)
ax.set_title(f"Distribuição de Notas ({turma_opcional})")
ax.set_xlabel("Total de Notas")
ax.set_ylabel("Frequência")
ax.legend().remove()  # Remove a legenda para clareza
st.pyplot(fig)

# Gráfico de Radar: Desempenho Médio por Turma
# Filtro para o gráfico de radar
st.header("Desempenho Médio por Componente (Radar Chart)")
opcoes_turmas = ['Todas'] + list(data['Turma'].unique())
turma_radar = st.selectbox("Selecione a Turma (ou Todas):", opcoes_turmas)

fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
componentes = ['Inst_1', 'Inst_2', 'Inst_3', 'Comportamento', 'CAED']

if turma_radar == 'Todas':
    media_componentes = data.groupby('Turma')[componentes].mean()
    for turma, valores in media_componentes.iterrows():
        categorias = list(componentes)
        valores = list(valores)
        # Repetir o primeiro valor para fechar o gráfico
        valores += valores[:1]
        categorias += categorias[:1]
        angles = [n / float(len(categorias) - 1) * 2 *
                  pi for n in range(len(categorias))]
        ax.plot(angles, valores, label=f'Turma {turma}')
        ax.fill(angles, valores, alpha=0.25)
else:
    media_componentes = data[data['Turma'] == turma_radar][componentes].mean()
    categorias = list(componentes)
    valores = list(media_componentes)
    valores += valores[:1]  # Repetir o primeiro valor para fechar o gráfico
    categorias += categorias[:1]
    angles = [n / float(len(categorias) - 1) * 2 *
              pi for n in range(len(categorias))]
    ax.plot(angles, valores, label=f'Turma {turma_radar}')
    ax.fill(angles, valores, alpha=0.25)

angulos_formatados = [f"{valor:.2f}" for valor in angles]
valores_formatados = [f"{valor:.2f}" for valor in valores]

col1, col2 = st.columns(2, gap="small")

col1.write("Ângulos")
col1.write(angulos_formatados)

col2.write("Valores")
col2.write(valores_formatados)


# Configuração do radar
ax.set_yticks([])
ax.set_xticks(angles[:-1])
ax.set_xticklabels(componentes)
ax.set_title("Média por Componente")
ax.legend(loc='upper right', bbox_to_anchor=(
    1.3, 1))  # Legenda fora do gráfico
st.pyplot(fig)
# dash
