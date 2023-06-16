#FUNCOES DE CORRECAO DE NUMERO
#Created by João Holanda - SimCo
#Datetime - 15/06/23
import re
#funcao pra adicionar o 9 do digito em numeros com somente 10 charach!
def Corretor(numero):
    numero = str(numero)

    padrao = r"(\d)\1{5,}"
    
    # Verifica se o padrão é encontrado no número
    if re.search(padrao, numero):
        numero = 0 
        return numero
  

    
    if len(numero) > 13:
            numero = numero
            return numero
    
    if len(numero) == 13:
         numero = numero[2:]
         return numero
    #caso ele possuo 12 carac
    if len(numero) == 12:
        #ex: 55 8599999999 - caso possua o identificador do pais
        if str(numero)[:2] == "55" or str(numero)[:2] == str(numero)[2:]:
            numero = numero[2:]
            numero = numero[:2] + '9' + numero[2:]
        else:
            return numero
    
    #caso em que não precisa mudar nada
    if len(numero) == 11:
         return numero
    
    #validando se possui realmente 10 carachteres
    if len(numero) == 10:
        #validando se é um numero de celular e não de telefone fixo!
        if numero[2] != '1' and  numero[2] != '2' and numero[2] != '3' and numero[2] != '4' and numero[2] != '5': 
            if len(numero) >= 2:  # Verifica se o número tem pelo menos dois dígitos
                numero = numero[:2] + '9' + numero[2:]
            return numero
        
        else:
            numero = 0
            return numero

    #validando se possui realmente 10 carachteres
    if len(numero) == 9:
        numero = '85' + numero
        return numero      
    elif len(numero) == 8 :
            numero = '859' + numero
            return numero       
    
    
    #ilegivel
    if len(numero) <= 7:
        numero = '0'
        return numero
     
     #caso esteja vazio igualar a 0
    if len(numero) == 0:
         numero = 0


    


import pandas as pd
tabela = pd.read_excel('Planilha_A_editar.xlsx')


linha = 0
validador = 1

while(validador != 0):
        
        nome = int(tabela.loc[linha,'CD_PESSOA_FISICA'])
        #verificador pra se a proxima linha de numero for vazia e não for a ultima ele iguala a 0!
        try:    
            number = int(tabela.loc[linha,'TEL_CELULAR'])

        except ValueError:
            number = 0
        #pega o numero que foi tratado e atualiza o campo com o numero corrigido, caso não precise ele só repete!
   
        tabela.loc[tabela['CD_PESSOA_FISICA'] == nome , 'TEL_CELULAR'] = Corretor(number)
        
        
        linha+=1

        
             
        #quando ele esta na ultima linha ele muda o validador pra 0, tirando ele do laço e finalizando
        if linha == len(tabela):
             tabela = tabela[tabela['TEL_CELULAR'] != '0']
             tabela = tabela[tabela['TEL_CELULAR'] != 0]

             #salvando excel como novo nome
             tabela.to_excel("Pacientes 09.06 a 15.06 JUNHO.xlsx",index=False) 
             validador = 0
        


else:
     quit()

