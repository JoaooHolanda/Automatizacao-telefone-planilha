import re
import datetime
import pandas as pd
import ctypes

def Corretor(numero):
    numero = str(numero)
    
    padrao = r"(\d)\1{5,}"
    
    if re.search(padrao, numero):
        return '0'
    
    if len(numero) > 13:
        return numero
    
    if len(numero) == 13:
        return numero[2:]
    
    if len(numero) == 12:
        if str(numero)[:2] == "55" or str(numero)[:2] == str(numero)[2:]:
            return numero[:2] + '9' + numero[2:]
        else:
            return numero
    
    if len(numero) == 11:
        if numero[2] not in {'1', '2', '3', '4', '5'}:
            return numero[:2] + '9' + numero[2:]
        else:
            return numero
    
    if len(numero) == 10:
        return numero[:2] + '9' + numero[2:]
    
    if len(numero) == 9:
        return '85' + numero
    
    if len(numero) == 8:
        return '859' + numero
    
    if len(numero) <= 7:
        return '0'
    
    if len(numero) == 0:
        return '0'

def exibir_alerta(titulo, mensagem):
    ctypes.windll.user32.MessageBoxW(0, mensagem, titulo, 0)

try:
    tabela = pd.read_excel('Base_A_Tratar.xlsx')
except FileNotFoundError:
    exibir_alerta('Arquivo não encontrado', 'Por gentileza despejar na pasta um arquivo com o nome: Mailing a tratar.xlsx')

linha = 0

while linha < len(tabela):
    nome = int(tabela.loc[linha, 'CD_PESSOA_FISICA'])

    try:
        number = int(tabela.loc[linha, 'TEL_CELULAR'])
    except ValueError:
        number = 0

    tabela.loc[tabela['CD_PESSOA_FISICA'] == nome, 'TEL_CELULAR'] = Corretor(number)
    linha += 1

tabela = tabela[(tabela['TEL_CELULAR'] != '0') & (tabela['TEL_CELULAR'] != 0)]

data_atual = datetime.datetime.now().date()
tabela.to_excel(f"{data_atual}_tratado.xlsx", index=False)
