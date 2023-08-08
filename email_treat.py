import pandas as pd
import re

# Função para verificar se um email é inválido
def is_invalid_email(email):

    email = str(email)
    # Verificar se o email possui um domínio inválido
    if not re.match(r"[^@]+@[^@]+\.[^@]+(\.br)?$", email):
        return True
    
    # Verificar se o email contém um número de telefone
    if re.match(r".*\d{3}.*\d{3}.*\d{4}.*", email):
        return True
    
    # Verificar se o email contém palavras proibidas sem o caractere '@'
    if re.match(r".*(naotem|nao)$", email, re.IGNORECASE):
        return True
    
    # Verificar se o email contém 4 letras repetidas
    if re.match(r".*(\w)\1{3}.*", email):
        return True
    
    return False

# Função para corrigir domínios escritos incorretamente
def fix_domain(email):
    if '@' in email:
        local_part, domain = email.split('@')
        domain = domain.lower().replace('gmial', 'gmail')  # Exemplo de correção específica
        return f"{local_part}@{domain}"
    return email

# Carregar a base de dados em Excel usando o Pandas
df = pd.read_excel('teste.xlsx')

# Verificar se o email possui os domínios necessários e removê-los caso não estejam presentes
required_domains = ['hotmail', 'clinicasim', 'gmail', 'yahoo', 'outlook']  # Domínios necessários
df = df[df['email'].str.contains('|'.join(required_domains), case=False, na=False)]

# Limpar os emails inválidos
df = df[~df['email'].apply(is_invalid_email)]

# Corrigir domínios escritos incorretamente
df['email'] = df['email'].apply(fix_domain)

# Remover emails duplicados
df.drop_duplicates(subset='email', inplace=True)

# Exibir o resultado da limpeza
print(f"Total de emails válidos: {len(df)}")

# Exportar a base de dados limpa para um novo arquivo Excel
df.to_excel('limpa.xlsx', index=False)

# Exibir uma janela no Windows indicando que a limpeza foi concluída
import ctypes
ctypes.windll.user32.MessageBoxW(0, "A limpeza dos emails inválidos foi concluída!", "Limpeza Concluída", 1)
