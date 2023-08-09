import re
import datetime
import pandas as pd
import ctypes
import tkinter as tk

def Corretor(numero):
    numero = str(numero)
    
    padrao = r"(\d)\1{5,}"
    
    if re.search(padrao, numero):
        return '0'
    
    if numero.startswith('55'):
        return numero[2:]
    
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
            return numero
    
    if len(numero) == 10:
        if numero[2] not in {'1', '2', '3', '4', '5'}:
            return numero[:2] + '9' + numero[2:]
        else:
            #tira os telefones fixos esse de baixo
            # numero = 0
            return numero
    
    if len(numero) == 9:
        return '85' + numero
    
    if len(numero) == 8:
        return '859' + numero
    
    if len(numero) <= 7:
        return '0'
    
    if len(numero) == 0:
        return '0'

def capturar_dados():
    janela.quit()  # Fecha a janela

# Cria a janela principal
janela = tk.Tk()
janela.title("Interface de Entrada de Dados")

# Cria uma variável para armazenar a entrada
entrada_var = tk.StringVar()

# Cria um rótulo para o título
titulo_label = tk.Label(janela, text="Nome do Arquivo")
titulo_label.pack()

# Cria um rótulo para a descrição
descricao_label = tk.Label(janela, text="Digite abaixo o nome que você deseja:")
descricao_label.pack()

# Cria um campo de entrada
entrada_entry = tk.Entry(janela, textvariable=entrada_var)
entrada_entry.pack()

# Cria um botão para capturar os dados
capturar_botao = tk.Button(janela, text="Capturar Dados", command=capturar_dados)
capturar_botao.pack()


# Inicia o loop principal da interface gráfica
janela.mainloop()


def exibir_alerta(titulo, mensagem):
    ctypes.windll.user32.MessageBoxW(0, mensagem, titulo, 0)

try:
    tabela = pd.read_excel('Base_A_Tratar.xlsx')
except FileNotFoundError:
    exibir_alerta('Arquivo não encontrado', 'Por gentileza despejar na pasta um arquivo com o nome: Mailing a tratar.xlsx')

linha = 0

while linha < len(tabela):
    nome = str(tabela.loc[linha, 'CD_PESSOA_FISICA'])

    try:
        number = int(tabela.loc[linha, 'TEL_CELULAR'])
    except ValueError:
        number = 0

    tabela.loc[tabela['CD_PESSOA_FISICA'] == nome, 'TEL_CELULAR'] = Corretor(number)
    linha += 1

tabela = tabela[(tabela['TEL_CELULAR'] != '0') & (tabela['TEL_CELULAR'] != 0)]

data_atual = datetime.datetime.now().date()
tabela.to_excel(f"{entrada_var.get()}_tratado.xlsx", index=False)
