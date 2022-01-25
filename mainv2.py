import pandas as pd
from tkinter import filedialog as fd


def isNaN(num):
    return num != num


colunas_manter = ['Circuito', 'Descrição', 'Fases', 'In - R', 'In - S', 'In - T', 'Ip']
nome_tabela = fd.askopenfilename(title="Escolha o arquivo dos circuitos", filetypes=[("Arquivos excel", "*.xlsx")])
circuitos_df = pd.read_excel(nome_tabela, header=3)
circuitos_df.dropna(subset=['Circuito'], inplace=True)
circuitos_df.drop(index=circuitos_df.index[-1], axis=1, inplace=True)
circuitos_df.drop(index=circuitos_df.index[-1], axis=1, inplace=True)


nomes_colunas = circuitos_df.columns.tolist()
deletar_coluna = []
for coluna in nomes_colunas:
    if not coluna in colunas_manter:
        deletar_coluna.append(coluna)

circuitos_df.drop(columns=deletar_coluna, inplace=True)
circuitos_df.reset_index(drop=True, inplace=True)

informacoes =[]
for coluna in colunas_manter:
    coluna1 = [x for x in circuitos_df[coluna]]
    informacoes.append(coluna1)

valor_ip = informacoes[-1]

for i,corrente in enumerate(informacoes[-4]):
    if not isNaN(corrente):
        informacoes[-4][i] = valor_ip[i]

for i,corrente in enumerate(informacoes[-3]):
    if not isNaN(corrente):
        informacoes[-3][i] = valor_ip[i]


for i,corrente in enumerate(informacoes[-2]):
    if not isNaN(corrente):
        informacoes[-2][i] = valor_ip[i]

novos_circuitos_df = pd.DataFrame(informacoes).T

novos_circuitos_df.columns=colunas_manter

nova_tabela = fd.asksaveasfilename(title="Salvar arquivo", defaultextension=".xlsx", filetypes=[("Arquivos excel", "*.xlsx")])

novos_circuitos_df.to_excel(nova_tabela)
