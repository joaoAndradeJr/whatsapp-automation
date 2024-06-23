import pywhatkit as kit
import time
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException
import os

# Caminho para o perfil do usuário do Chrome onde você está logado no WhatsApp Web
chrome_profile_path = r'C:\\Users\\Joao\\AppData\\Local\\Google\\Chrome\\User Data\\Default'

# Verificação do perfil do Chrome
try:
    with open(chrome_profile_path + '/Preferences', 'r') as f:
        pass
except FileNotFoundError:
    print(f"Erro: o perfil do Chrome em '{chrome_profile_path}' não foi encontrado.")
    print("Certifique-se de ajustar 'chrome_profile_path' para o caminho correto do seu perfil do Chrome.")
    exit(1)

# Configuração do ChromeOptions
chrome_options = Options()
chrome_options.add_argument(f'--user-data-dir={chrome_profile_path}')
chrome_options.add_argument('--profile-directory=Default')  # Nome do diretório do perfil

# Configuração do Service do ChromeDriver
# Como o ChromeDriver está integrado a partir da versão 115 do Chrome, não é necessário especificar o caminho do executável
service = Service()

# Inicialização do WebDriver
try:
    driver = webdriver.Chrome(service=service, options=chrome_options)
except WebDriverException as e:
    print(f"Erro ao iniciar o WebDriver: {e}")
    exit(1)

# Carregar os dados da planilha
print("Carregando a planilha...")
try:
    df = pd.read_excel('C:\\Users\\Joao\\Documents\\Meus-Projetos\\Pasta1.xlsx')  # Substitua pelo caminho correto
    print("Planilha carregada com sucesso.")
except Exception as e:
    print(f"Erro ao carregar a planilha: {e}")
    if driver:
        driver.quit()
    exit(1)

# Abre o WhatsApp Web na primeira guia
driver.get('https://web.whatsapp.com')

# Espera até que o usuário esteja logado no WhatsApp Web
# input("Após o login no WhatsApp Web, pressione Enter para continuar...")

# Função para calcular o próximo horário de envio
def get_next_time(current_time):
    next_minute = current_time.minute + 1
    if next_minute == 60:
        return current_time.replace(hour=current_time.hour + 1, minute=0)
    else:
        return current_time.replace(minute=next_minute)

# Horário de envio inicial
current_time = datetime.now().replace(second=0, microsecond=0) + timedelta(minutes=2)  # Arredonda para o minuto mais próximo
print(f"Horário de envio inicial: {current_time.hour}:{current_time.minute}")

try:
    for index, row in df.iterrows():
        phoneNumber = row['Número']
        name = row['Nome']
        message = row['Mensagem']

        print(f"Enviando mensagem para {phoneNumber} ({name}) às {current_time.hour}:{current_time.minute}")
        
        try:
            kit.sendwhatmsg(f"+55{phoneNumber}", message, current_time.hour, current_time.minute)
            
        except Exception as e:
            print(f"Erro ao enviar mensagem para {phoneNumber}: {e}")

        # Espera antes de enviar a próxima mensagem
        print("Esperando 30 segundos antes de enviar a próxima mensagem...")
        time.sleep(30)  # Espera 30 segundos antes de enviar a próxima mensagem
        
        # Incrementa o tempo de envio para o próximo número
        current_time = get_next_time(current_time)

except KeyboardInterrupt:
    print("\nProcesso interrompido.")
finally:
    # Fechar todas as guias abertas e o navegador
    try:
        while len(driver.window_handles) > 0:
            driver.switch_to.window(driver.window_handles[-1])
            driver.close()
        driver.quit()
    except WebDriverException as e:
        print(f"Erro ao fechar o WebDriver: {e}")

    # Fecha todas as instâncias do Chrome
    os.system("taskkill /f /im chrome.exe")
    print("Mensagens enviadas com sucesso e todas as janelas do Chrome foram fechadas.")
