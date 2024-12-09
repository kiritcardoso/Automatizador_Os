'''from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import pandas as pd
import time

# Configuração de variáveis
EMAIL = "jhonathancardosol.2017@gmail.com"
SENHA = "Electabuzz1"
SITE_LOGIN = "https://service.telecontrol.com.br"
SITE_BUSCA = (
    "https://service.telecontrol.com.br/os?data_inicial=04/12/2023&data_final=04/12/2024&contexto_data=digitacao"
)  # A data está como string
EXCEL_PATH = "Maquinas_of_comp.xlsx"
CHROME_DRIVER_PATH = "./chromedriver.exe"  # Caminho do ChromeDriver

# Configurando o Selenium
service = Service(CHROME_DRIVER_PATH)
driver = webdriver.Chrome(service=service)
wait = WebDriverWait(driver, 15)


def login_site():
    """Realiza login no site."""
    try:
        driver.get(SITE_LOGIN)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']"))).send_keys(EMAIL)
        driver.find_element(By.CSS_SELECTOR, "input[type='password']").send_keys(SENHA)
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()    # Certifique-se de que redirecionou para a página correta
        print("Login realizado com sucesso.")
    except TimeoutException:
        print("Erro ao realizar o login. Verifique as credenciais ou o site.")
        driver.quit()
        exit()


def buscar_dados(numero_os):
    """Busca os dados de um número de cadastro no site."""
    driver.get(SITE_BUSCA)
    try:
        search_box = wait.until(EC.presence_of_element_located((By.ID, "os_numero")))
        search_box.clear()
        search_box.send_keys(numero_os)
        search_box.send_keys(Keys.RETURN)
        time.sleep(3)  # Pausa para carregar os resultados

        modelo = driver.find_element(By.XPATH, "//td[@data-label='Produto']").text
        preco = driver.find_element(By.XPATH, "//td[@data-label='Total']").text
        status = driver.find_element(By.XPATH, "//td[@data-label='Status']").text
        return modelo, preco, status
    except NoSuchElementException:
        print(f"Número OS {numero_os} não encontrado.")
        return None, None, None
    except TimeoutException:
        print(f"Timeout ao buscar dados para o número OS {numero_os}.")
        return None, None, None


def atualizar_excel():
    """Atualiza o Excel com os dados coletados do site."""
    try:
        df = pd.read_excel(EXCEL_PATH)

        # Garantir que as colunas necessárias existam
        if "Modelo da Maquina" not in df.columns:
            print("Coluna 'Modelo da Maquina' não encontrada no Excel.")
            df["Modelo da Maquina"] = ""  # Criar a coluna caso não exista

        # Garantir que as colunas tenham os tipos corretos
        df["Status"] = df["Status"].astype(str)
        df["Preço"] = df["Preço"].astype(str)
        df["Modelo da Maquina"] = df["Modelo da Maquina"].astype(str)

        for index, row in df.iterrows():
            maquina = row.get("Maquinas")
            if pd.isna(maquina) or not str(maquina).isdigit():
                continue  # Pular se a célula estiver vazia ou não for um número

            modelo, preco, status = buscar_dados(maquina)
            if modelo and preco and status:
                df.loc[index, "Status"] = status
                df.loc[index, "Preço"] = preco
                df.loc[index, "Modelo da Maquina"] = modelo
                print(f"Dados atualizados para a máquina: {maquina}")
            else:
                print(f"Não foi possível atualizar dados para a máquina: {maquina}")

        df.to_excel(EXCEL_PATH, index=False)
        print("Excel atualizado com sucesso.")
    except FileNotFoundError:
        print("Arquivo Excel não encontrado. Verifique o caminho especificado.")
    except Exception as e:
        print(f"Ocorreu um erro ao atualizar o Excel: {e}")


# Executando o script
try:
    login_site()
    atualizar_excel()
finally:
    driver.quit()'''
import pandas as pd

# Caminho para o arquivo Excel
EXCEL_PATH = "Maquinas_of_comp.xlsx"

# Carregar e exibir a tabela
try:
    df = pd.read_excel(EXCEL_PATH)
    print(df)  # Exibe a tabela completa
except FileNotFoundError:
    print("Arquivo Excel não encontrado. Verifique o caminho.")
except Exception as e:
    print(f"Ocorreu um erro ao carregar o Excel: {e}")