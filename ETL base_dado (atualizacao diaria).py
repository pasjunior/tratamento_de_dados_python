# importando bibliotecas
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog
import time
import os
import locale


# cria uma janela vazia
root = tk.Tk()
root.withdraw()

# abre uma caixa de diálogo para selecionar o arquivo
pedidos = filedialog.askopenfilename(title="Selecione arquivo Pedidos")
fechamentos = filedialog.askopenfilenames(title="Selecione os arquivos Fechamentos", filetypes=(("Arquivos de Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
nome_arquivos_fechamento = tuple(os.path.basename(f) for f in list(fechamentos))

# definir locale para o formato brasileiro
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# formata números no padrão brasileiro
def formatar_numero(numero):
    try:
        numero = float(numero)
        return locale.currency(numero, grouping=True)
    except ValueError:
        return ''

# lendo arquivo excel - Pedidos
print('Tratando a base Pedidos')
df_pedidos = pd.read_excel(pedidos, sheet_name='Pedidos', usecols=['Empresas do grupo', 'Natureza do processo', 'Polo_Processual', 'Pasta_Situacao', 'Pasta_Motivoencerramento',
                                                                   'Numeroprocesso', 'Objeto', 'Advogadoexterno_Nome', 'Criado_Em', 'Risco_Remoto', 'Pasta_Encerradoem', 'Pasta_Websijur',
                                                                   'Datadistribuicao', 'Valor_Pedido'])

df_pedidos = df_pedidos.drop(df_pedidos.index[-1])
df_pedidos['Objeto'] = df_pedidos['Objeto'].str.strip()
df_pedidos['chave'] = df_pedidos['Pasta_Websijur'].str.extract(r'\.(\d+)/')
df_pedidos['Criado_Em'] = df_pedidos['Criado_Em'].dt.strftime('%d/%m/%y')
df_pedidos['Numeroprocesso'] = df_pedidos['Numeroprocesso'].astype(str)

# exportando arquivo em excel
#df_pedidos.to_excel('Pedidos.xlsx', sheet_name='Pedidos', index=False)
df_pedidos.to_parquet('Pedidos.parquet', engine='pyarrow')
del df_pedidos
print('Pedidos.xlsx')

# lendo arquivo excel - Fechamentos
print('Tratando as bases de Fechamento')
df_fechamento = pd.DataFrame()
for fechamento, nome_arquivo in zip(fechamentos, nome_arquivos_fechamento):
    df_fechamento = pd.read_excel(fechamento, sheet_name='BASE', usecols=['EMPRESA', 'COMPETÊNCIA', 'NÚMERO PROCESSO', 'PASTA', 'NATUREZA', 'POLO', 'OBJETO', 'PROGNÓSTICO PREDOMINANTE',
                                                                          'PEDIDO ATUALIZADO', 'PROVÁVEL MÊS ANTERIOR', 'PROVÁVEL ATUALIZADO', 'MOTIVO DO ENCERRAMENTO', 'TIPO',
                                                                          'ADVOGADO EXTERNO', 'VALOR ATUALIZAÇÃO PROVÁVEL', 'DATA DISTRIBUIÇÃO'])
    df_fechamento = df_fechamento.apply(lambda x: x.astype(str))
    df_fechamento['Cod'] = df_fechamento['EMPRESA'].str.extract(r'\- (\d+)')
    if not df_fechamento['PASTA'].isnull().all():
        df_fechamento['chave'] = df_fechamento['PASTA'].str.extract(r'\.(\d+)/')
    else:
        df_fechamento['chave'] = ''
    df_fechamento[['VALOR ATUALIZAÇÃO PROVÁVEL', 'PEDIDO ATUALIZADO', 'PROVÁVEL ATUALIZADO', 'PROVÁVEL MÊS ANTERIOR']] = \
        df_fechamento[['VALOR ATUALIZAÇÃO PROVÁVEL', 'PEDIDO ATUALIZADO', 'PROVÁVEL ATUALIZADO', 'PROVÁVEL MÊS ANTERIOR']].applymap(formatar_numero)
    df_fechamento[['VALOR ATUALIZAÇÃO PROVÁVEL', 'PEDIDO ATUALIZADO', 'PROVÁVEL ATUALIZADO', 'PROVÁVEL MÊS ANTERIOR','DATA DISTRIBUIÇÃO']] = \
        df_fechamento[['VALOR ATUALIZAÇÃO PROVÁVEL', 'PEDIDO ATUALIZADO', 'PROVÁVEL ATUALIZADO', 'PROVÁVEL MÊS ANTERIOR', 'DATA DISTRIBUIÇÃO']].apply(lambda x: x.fillna(0))
    df_fechamento[['VALOR ATUALIZAÇÃO PROVÁVEL', 'PEDIDO ATUALIZADO', 'PROVÁVEL ATUALIZADO', 'PROVÁVEL MÊS ANTERIOR','DATA DISTRIBUIÇÃO']] = \
        df_fechamento[['VALOR ATUALIZAÇÃO PROVÁVEL', 'PEDIDO ATUALIZADO', 'PROVÁVEL ATUALIZADO', 'PROVÁVEL MÊS ANTERIOR', 'DATA DISTRIBUIÇÃO']].apply(lambda x: x.str.replace('-', ''))
    df_fechamento['DATA DISTRIBUIÇÃO'] = pd.to_datetime(df_fechamento['DATA DISTRIBUIÇÃO'])
    df_fechamento['DATA DISTRIBUIÇÃO'] = df_fechamento['DATA DISTRIBUIÇÃO'].dt.strftime('%d/%m/%y')
    df_fechamento = df_fechamento.rename(columns={
        'NÚMERO PROCESSO': 'NUMERO PROCESSO',
        'COMPETÊNCIA': 'COMPETENCIA',
        'PROGNÓSTICO PREDOMINANTE': 'PROGNOSTICO PREDOMINANTE',
        'PROVÁVEL MÊS ANTERIOR': 'PROVAVEL MES ANTERIOR',
        'PROVÁVEL ATUALIZADO': 'PROVAVEL ATUALIZADO',
        'VALOR ATUALIZAÇÃO PROVÁVEL': 'VALOR ATUALIZACAO PROVAVEL',
        'DATA DISTRIBUIÇÃO': 'DATA DISTRIBUICAO'
    })
    df_fechamento['NUMERO PROCESSO'] = df_fechamento['NUMERO PROCESSO'].replace(to_replace='\n', value=' ', regex=True)
    #nome_arquivo_saida = os.path.splitext(nome_arquivo)[0] + '_output.xlsx'
    #df_fechamento.to_excel(nome_arquivo, encoding='utf8', sheet_name='BASE', index=False)
    nome_arquivo_saida = os.path.splitext(nome_arquivo)[0] + '.parquet'
    df_fechamento.to_parquet(nome_arquivo, engine='pyarrow')
    print(nome_arquivo)
del df_fechamento

print('Processo encerrado!')
time.sleep(10)