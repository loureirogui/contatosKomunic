import traceback
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from docx import Document
import openpyxl
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import re
import unicodedata
from art import text2art, tprint  # Importa as funções da biblioteca art

tprint("Contatos\n    Komunic", font="starwars")

print("Seja bem-vindo à gambiarra (Brincadeiras à parte) desenvolvida para cadastrar contatos em massa no sistema Komunic.\n"
      "Para utilizar, será necessário uma planilha onde a coluna A1 contenha os nomes e a coluna B1 os números!\n"
      "Lembre-se de remover caracteres especiais do telefone, mantendo apenas números.\n"
      "Depois é só colocar a planilha na mesma pasta desse código com o nome (contatos.xlsx)\n"
      "Uma janela do Edge aparecerá e iniciará o processo de cadastro dos contatos que possuirem whatsapp válido")


perguntaSeguranca = input("Realizou as instruções acima?\n")

# Solicita as credenciais de acesso ao Komunic
emailLogin = input("Qual o email de login da Komunic?\n")
senhaLogin = input("Qual o senha de login da Komunic?\n")
print("Iniciando processo de inclusão dos contatos. Aguarde por gentileza...")



# Configura as opções do Edge
edge_options = Options()
edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])


edge_driver = webdriver.Edge(    service=Service(EdgeChromiumDriverManager().install()), options=edge_options)

# Abre o link desejado
url = f"https://app.komunic.net/login"
edge_driver.get(url)

# Carrega o arquivo .xlsx
workbook = openpyxl.load_workbook('contatosTeste.xlsx')

# Seleciona a planilha ativa (a primeira planilha aberta por padrão)
sheet = workbook.active

def format_phone_number(phone):
    phone = ''.join(filter(str.isdigit, phone))  # Remove caracteres não numéricos
    return phone

def normalize_name(name):#Função para remover caracteres especiais do nome evitando erro de cadastro na komunic
    """
    Remove acentos, caracteres especiais e retorna apenas letras e números no nome.
    """
    if not name:
        return ""
    name = unicodedata.normalize('NFKD', name)
    name = ''.join([c for c in name if not unicodedata.combining(c)])

    name = re.sub(r'[^a-zA-Z0-9 ]', '', name)

    name = name.strip()
    return name


# Lógica de login
try:
    email_input = WebDriverWait(edge_driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, '//*[@id="email"]'))
    )

    email_input.send_keys(emailLogin)
except Exception:
    print("Erro ao inserir o e-mail no campo de login:")
    

try:
    senha_input = WebDriverWait(edge_driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, '//*[@id="password"]'))
    )

    senha_input.send_keys(senhaLogin)
except Exception:
    print("Erro ao inserir a senha no campo de senha:")
    

try:
    login_button = WebDriverWait(edge_driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div[2]/form/div[4]/button'))
    )
    
    login_button.click()
except Exception:
    print("Erro ao clicar no botão de login:")
    
# Espera 2 segundos após clicar no botão de login, tempo necessário para o selenium poder efetuar o login para prosseguirmos ao cadastro de contato
time.sleep(2)

url = f"https://app.komunic.net/contacts"
edge_driver.get(url)


for row in sheet.iter_rows(min_row=2, min_col=1, max_col=11):
        
    nomentratado_value = row[0].value # Coluna 1 nomes
    fone_value = row[1].value       # Coluna 2: fones

    if fone_value:
        formatted_number = fone_value
        nome_value = normalize_name(nomentratado_value)  # Normaliza o nome
        edge_driver.get(url)
        time.sleep(2)
        
        #Iniciando cadastro de novo contato
        try:
            NovoContato_button = WebDriverWait(edge_driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div[2]/main/header/div/div[2]/button[2]'))
            )
            NovoContato_button.click()
            time.sleep(2)
        except Exception:
            print("Erro ao clicar no botão de novo contato:")
        
        # Envia o número de telefone formatado para o campo
        try:
            telefoneWhatsApp = WebDriverWait(edge_driver, 2).until(
                EC.element_to_be_clickable((By.ID, 'connection_key'))
            )
            
            telefoneWhatsApp.send_keys(formatted_number)
            
        except Exception as e:
            print(f"Erro ao inserir número do contato; {nome_value}; {fone_value}")
        
        time.sleep(0.8)

        try: # Clica no botão de verificar telefone
            verificaWhatsApp = WebDriverWait(edge_driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/div[1]/div[2]/div/div[4]/button'))
            )
            
            verificaWhatsApp.click()
            time.sleep(0.5)
        except:
            print(f"Erro ao clicar no botão de verificar o whatsapp")
        

        time.sleep(0.5)

        try: #Aguarda pelo erro informando que não foi possível encontrar contato no whatsapp
            error_element = WebDriverWait(edge_driver, 1).until(
            EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div[2]/div[1]/div[2]/div/div[5]/div/p'))
            )
            if error_element:
                print("Número de telefone não possui whatsapp. Nome:" + nome_value + ': Telefone: ' + formatted_number)
                continue  # Pular para a próxima linha se o alerta for encontrado
            else:
                pass
        
        except Exception: # Não encontrando o erro, continua o processamento
          pass

        
        try: #Clicar no botão para validar o telefone
            confirmaCadastro = WebDriverWait(edge_driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div[2]/div[2]/button[2]'))
            )
            confirmaCadastro.click()
            
        except Exception as e:
            print(f"Erro ao clicar no botão de verificar o whatsapp: {e}")


        try: # Insere o nome do contato
            
            nomeContato = WebDriverWait(edge_driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'name'))
            )
            nomeContato.send_keys(nome_value)
            
        except Exception as e:
            print(f"Erro ao digitar o nome do contato: {e}")
            
        
        #TRECHO COMENTADO POIS LÓGICA DE VINCULAÇÃO DO CONTATO EM UMA ORGANIZAÇÃO ESTÁ EM DESENVOLVIMENTO PARA GARANTIR ASSERTIVIDADE NO PROCESSO
        # try:
        #     element = WebDriverWait(edge_driver, 20).until(
        #         EC.element_to_be_clickable((By.XPATH, '//*[@id="select_person"]'))
        #     )
        #     element.send_keys(nome_value)
        #     time.sleep(2)
        #     # Simula o pressionamento da seta para baixo e Enter
        #     actions = ActionChains(edge_driver)
        #     actions.send_keys(Keys.ARROW_DOWN).perform()
        #     actions.send_keys(Keys.ENTER).perform()

        # except Exception as e:
        #     print(f"Erro ao selecionar o nome do contato: {e}")
        # teste = input('breakpoint')
 
        try:
            save_button = WebDriverWait(edge_driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/div[2]/div/div[2]/button[2]'))
            )
            # Clique no botão de login
            save_button.click()
            time.sleep(1.5)
        except Exception:
            print("Erro ao clicar no botão de novo contato:")

        
print('Processo de cadastro concluido com sucesso. Até logo!')