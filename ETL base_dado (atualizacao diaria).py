# importando bibliotecas

import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog
import time
import os

# cria uma janela vazia
root = tk.Tk()
root.withdraw()

# abre uma caixa de diálogo para selecionar o arquivo
pedidos = filedialog.askopenfilename(title="Selecione arquivo Pedidos")
fechamentos = filedialog.askopenfilenames(title="Selecione os arquivos Fechamentos", filetypes=(("Arquivos de Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
nome_arquivos_fechamento = tuple(os.path.basename(f) for f in list(fechamentos))

# lendo arquivo excel - Pedidos
print('Tratando a base Pedidos')
df_pedidos = pd.read_excel(pedidos, sheet_name='Pedidos', usecols=['Empresas do grupo', 'Natureza do processo', 'Polo_Processual', 'Pasta_Situacao', 'Pasta_Motivoencerramento',
                                                                   'Numeroprocesso', 'Objeto', 'Advogadoexterno_Nome', 'Criado_Em', 'Risco_Remoto', 'Pasta_Encerradoem', 'Pasta_Websijur',
                                                                   'Datadistribuicao', 'Valor_Pedido'])

df_pedidos = df_pedidos.drop(df_pedidos.index[-1])
df_pedidos['Objeto'] = df_pedidos['Objeto'].str.strip()
df_pedidos['chave'] = df_pedidos['Pasta_Websijur'].str.extract(r'\.(\d+)/')

# exportando arquivo em excel
df_pedidos.to_excel('Pedidos.xlsx', sheet_name='Pedidos', index=False)
del df_pedidos
print('Pedidos.xlsx')

print('Processo encerrado!')
time.sleep(10)