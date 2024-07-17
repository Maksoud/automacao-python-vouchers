import pandas as pd
import requests
from bs4 import BeautifulSoup
import random
import time

########################################

# Funções para verificar o conteúdo da página
def check_voucher_status(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Verifique o conteúdo da página aqui
        if "sucesso" in soup.text:
            return "válido"
        elif "voucher já utilizado" in soup.text:
            return "utilizado"
        elif "número de tentativas excedida" in soup.text:
            return "excedida"
        else:
            # You cannot redeem this offer because you currently have an active Premium subscription.
            return "erro"
    except Exception as e:
        print(f"Erro ao acessar {url}: {e}")
        return "erro"

########################################

# Carregar os dados do Excel
df = pd.read_excel('vouchers.xlsx')

########################################

# Função para processar os vouchers
def process_vouchers(df):

    # Selecionar índices das linhas que ainda não possuem status
    rows_to_process = df[df['Status'].isna()].index.tolist()

    ####################################
    
    while rows_to_process:

        # Selecionar uma linha aleatória
        row_index = random.choice(rows_to_process)
        
        # Obter o link do voucher
        url = df.at[row_index, 'Link']
        
        # Verificar o status do voucher
        status = check_voucher_status(url)
        
        if status == "excedida":
            print("Número de tentativas excedida. Parando a execução.")
            break
        
        # Atualizar o status na planilha
        df.at[row_index, 'Status'] = status
        print(f"Voucher {url} status: {status}")
        
        # Remover a linha processada da lista
        rows_to_process.remove(row_index)
        
        # Salvar o progresso no Excel
        df.to_excel('vouchers.xlsx', index=False)
        
        # Esperar um tempo aleatório para evitar bloqueio
        time.sleep(random.uniform(1, 5))

########################################

# Adicionar coluna de status se não existir
if 'Status' not in df.columns:
    df['Status'] = None

########################################

# Processar os vouchers
process_vouchers(df)