import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import random
import time

########################################

# Configurar o Selenium
chrome_options = Options()
# chrome_options.add_argument("--headless")  # Executar o Chrome no modo headless (opcional)
# chrome_options.add_argument("--disable-gpu")
# chrome_options.add_argument("--no-sandbox")
# chrome_options.add_argument("--disable-dev-shm-usage")
winuser = "renee"  # Nome de usuário do Windows
chrome_options.add_argument(f"--user-data-dir=C:/Users/{winuser}/AppData/Local/Google/Chrome/User Data")  # Caminho para o diretório de dados do usuário do Chrome
chrome_options.add_argument("--profile-directory=Default")  # Diretório do perfil padrão

service = ChromeService(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

########################################

# Funções para verificar o conteúdo da página
def check_voucher_status(url):
    try:

        # Acesse a página do voucher
        driver.get(url)
        page_source = driver.page_source

        if "Something went wrong. Please try using username and password." or "Ocorreu um erro. Tente usar o nome de usuário e a senha." in page_source:
            print("Aguardando 15s para realização do login...")
            time.sleep(15)  # Aguarda o carregamento da página

        # Leia novamente a página após os 40 segundos reservados para o login
        driver.get(url)
        page_source = driver.page_source
        # print(page_source)

        ################################

        # Verifique o conteúdo da página aqui
        if "sucesso" in page_source:
            return "válido"
        elif "Oferta indisponível" in page_source:
            # Oferta indisponível
            # Esta oferta já foi resgatada. Use um link de
            # promoção válido ou entre em contato com o 
            # Suporte ao Cliente para obter mais ajuda.
            # oferta_indisponivel.png
            return "utilizado"
        elif "número de tentativas excedida" in page_source:
            return "excedida"
        elif "You cannot redeem this offer because you currently have an active Premium subscription." in page_source:
            # You cannot redeem this offer because you 
            # currently have an active Premium 
            # subscription.
            # something_went_wrong.png
            return "premium"
        else:
            print(f"Status desconhecido para {url}")
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
            print("Número de tentativas excedida. Parando a execução...")
            break
        elif status == "login":
            print("Erro! Necessário realizar login. Parando a execução...")
            break
        elif status == "premium":
            print("Erro! Já possui a assinatura Premium. Parando a execução...")
            break
        elif status == "erro":
            print("Erro ao acessar a página. Parando a execução...")
            break
        
        # Atualizar o status na planilha
        df.at[row_index, 'Status'] = status
        print(f"Voucher {url} status: {status}")
        
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

# Fechar o driver
driver.quit()