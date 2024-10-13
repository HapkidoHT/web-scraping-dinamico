import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

# Configuração inicial
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Abre o site
driver.get("https://www.fenabrave.org.br/relatorios/rel_MaisVendidos.asp")

# Aguarda a página carregar
time.sleep(5)

# Seleciona o segmento "Auto"
segmento_select = Select(driver.find_element(By.NAME, "segmento"))
segmento_select.select_by_value("3")# 1- Auto, 3- Caminhão
time.sleep(2)

# Seleciona o período "Junho/2024"
#Selecionarno período desejado
periodo_select = Select(driver.find_element(By.NAME, "periodo"))
periodo_select.select_by_value("6,2024")
time.sleep(5)

# Habilita o mapa clicando no botão "OK", se estiver presente
try:
    dismiss_button = driver.find_element(By.CLASS_NAME, "dismissButton")
    dismiss_button.click()
    time.sleep(3)
except Exception as e:
    print("Botão de habilitar mapa não encontrado ou já foi clicado:", e)

# Lista de estados com suas coordenadas
#Coordenadaa cartesianas do processo
estados = {
    "Amazonas": "left: -159px; top: -105px;",
    "Mato Grosso": "left: 24px; top: -33px;",
    "Rio Grande do Sul": "left: -26px; top: 155px;",
    "Acre": "left: 84px; top: 33px;",
    "Alagoas": "left: -35px; top: -121px;",
    "Amapá": "left: 76px; top: -16px;",
    "Bahia": "left: 145px; top: -37px;",
    "Ceará": "left: 248px; top: -74px;",
    "Distrito Federal": "left: 55px; top: -69px;",
    "Espírito Santo": "left: 114px; top: -97px;",
    "Goiás": "left: -38px; top: -113px;",
    "Maranhão": "left: -22px; top: -99px;",
    "Mato Grosso do Sul": "left: 80px; top: -142px;",
    "Minas Gerais": "left: 87px; top: -201px;",
    "Pará": "left: 171px; top: -217px;",
    "Paraíba": "left: 330px; top: -5px;",
    "Paraná": "left: 38px; top: -35px;",
    "Pernambuco": "left: 342px; top: -43px;",
    "Piauí": "left: 7px; top: -53px;",
    "Rio Grande do Norte": "left: 327px; top: -84px;",
    "Rondônia": "left: -254px; top: 85px;",
    "Roraima": "left: 108px; top: -98px;",
    "Santa Catarina": "left: 28px; top: -45px;",
    "São Paulo": "left: 167px; top: -34px;",
    "Sergipe": "left: 13px; top: 106px;",
    "Rio de Janeiro": "left: 38px; top: -24px;",
    "Tocantins": "left: -55px; top: -180px;"
}

# Estados inicialmente habilitados
estados_iniciais = ["Amazonas", "Mato Grosso", "Rio Grande do Sul"]

# Data de raspagem
data_atual = time.strftime("%d/%m/%Y")

# Lista para armazenar os dados
dados = []

# Função para raspar os dados de um estado
def raspar_estado(estado):
    try:
        estado_element = driver.find_element(By.CSS_SELECTOR, f"div[title='{estado}']")
        driver.execute_script("arguments[0].click();", estado_element)
        time.sleep(3)  # Aguarda a atualização dos dados

        # Raspa os dados das marcas e quantidades
        linhas = driver.find_elements(By.ID, "linha")
        for linha in linhas:
            marca = linha.find_element(By.CSS_SELECTOR, "#marca span").text
            quantidade = linha.find_element(By.CSS_SELECTOR, "#contBarra #val").text
            dados.append([marca, quantidade, estado, data_atual])

        # Aguarda um pouco antes de passar para o próximo estado
        time.sleep(2)
    except Exception as e:
        print(f"Erro ao processar estado {estado}: {e}")

# Raspa os dados dos estados inicialmente habilitados
for estado in estados_iniciais:
    raspar_estado(estado)

# Função para aumentar o zoom no mapa
def aumentar_zoom():
    try:
        zoom_in_button = driver.find_element(By.XPATH, "//button[@aria-label='Aumentar o zoom']")
        zoom_in_button.click()
        time.sleep(2)  # Ajusta conforme necessário
    except Exception as e:
        print("Erro ao tentar aumentar o zoom:", e)

# Lista de estados processados para evitar duplicação
estados_processados = set(estados_iniciais)

# Função para verificar se um estado é visível
def estado_visivel(estado):
    try:
        driver.find_element(By.CSS_SELECTOR, f"div[title='{estado}']")
        return True
    except:
        return False

# Loop para aumentar o zoom e raspar os dados dos demais estados
while len(estados_processados) < len(estados):
    for estado in estados:
        if estado not in estados_processados and estado_visivel(estado):
            raspar_estado(estado)
            estados_processados.add(estado)
    aumentar_zoom()

# Fecha o navegador
driver.quit()

# Cria um DataFrame com os dados
df = pd.DataFrame(dados, columns=["Marca", "Quantidade", "Estado", "Data"])

# Salva o DataFrame em um arquivo Excel
df.to_excel("dados_veiculos_por_estado.xlsx", index=False)

print("Raspagem e salvamento concluídos com sucesso!")