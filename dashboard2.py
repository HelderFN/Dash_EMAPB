# %%
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Carregando dados

# %%
df = pd.read_excel("Notas_9ano.xlsx", sheet_name="Notas")
print(df.head())

# configurando o titulo
st.markdown(
    f"""
    <h1 style="font-size:25px; color:black; text-align:left;">
        Notas por aluno
    </h1>
    """,
    unsafe_allow_html=True
)

# Grafico 1: Notas por aluno, comparando com a media

col1, col2 = st.columns([1, 2])

with col1:
    st.write("")
    st.write("")
    st.write("")
    st.write("")
    turma = st.selectbox("Selecione a  turma", df['Turma'].unique())

    aluno_turma = df[df['Turma'] == turma]['Nome'].unique()

    aluno = st.selectbox("Selecione um aluno", aluno_turma)

    dados_df = df[(df['Turma'] == turma) & (df['Nome'] == aluno)]


with col2:

    # st.header(f"Notas: {str(aluno).capitalize()}")
    cabecalho = f"{str(aluno)}"
    st.markdown(
        f"""
    <h1 style="font-size:25px; color:black; text-align:center;">
        {cabecalho}
    </h1>
    """,
        unsafe_allow_html=True
    )
    media = df.iloc[:, 2:9].groupby(by='Turma').mean()

    dados_aluno = dados_df[dados_df['Nome'] == aluno].iloc[0, 4:9]
    fig, ax = plt.subplots(1, 1, figsize=(12, 6))

    ax.bar(dados_aluno.index, dados_aluno.values, color='green')
    ax.set_ylim(0, 35)
    ax.grid(which='major', axis='y')
    ax.set_xticklabels(dados_aluno.index, fontdict={'size': 20})
    ax.set_yticklabels(list(range(0, 40, 5)), fontdict={"size": 20})
    st.pyplot(fig)

    # Grafico de medias por turma

st.markdown(
    """
    <h1 style="font-size:25px; color:black; text-align:left;">
    Média de notas por turma
    </h1>
    """,
    unsafe_allow_html=True
)

bim = st.selectbox("Selecione o bimestre",
                   list(df['Bimestre'].unique()))

fig, ax = plt.subplots()
categorias = list(media.columns)
categorias.remove('Bimestre')
x = np.arange(len(media.index))

largura = 0.15

for i, categoria in enumerate(categorias):
    ax.bar(
        x + i * largura,
        media[categoria],
        largura,
        label=categoria
    )
ax.set_title(f"Médias das Notas Por Turma - {bim}º bimestre", fontsize=16)
ax.set_xticks(x+largura*(len(categorias)-1)/2)
ax.set_xticklabels(media.index)
ax.set_ylabel("Média")
ax.legend()
plt.xticks(rotation=45)
plt.tight_layout()
st.pyplot(fig)
