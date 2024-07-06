import json
import schedule
import time
from datetime import date
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep as sl
from datetime import datetime
import os

# Configuracoes do driver
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_argument('--headless') # Esconde o navegador
driver = webdriver.Chrome(service=service, options=options)

# Carregar dados do JSON
with open('Info.json') as f:
    config = json.load(f)

cpf = config["cpf"]
password = config["password"]
acoes_horarios = config["horarios"]

# Carregar Planilha
if not os.path.exists('tabelaponto.xlsx'):
    df = pd.DataFrame(columns=["Data", "Entrada1", "Saida1", "Entrada2", "Saida2", "Horas Trabalhadas"])
    df.to_excel('tabelaponto.xlsx', index=False)
else:
    df = pd.read_excel('tabelaponto.xlsx')


# Contador de ações executadas
contador_acoes = 0


def calcular_horas_trabalhadas(entrada1, saida1, entrada2, saida2):
    # Datetime
    entrada1 = datetime.strptime(entrada1, "%H:%M")
    saida1 = datetime.strptime(saida1, "%H:%M")
    entrada2 = datetime.strptime(entrada2, "%H:%M")
    saida2 = datetime.strptime(saida2, "%H:%M")

    periodo1 = saida1 - entrada1
    periodo2 = saida2 - entrada2

    horas_trabalhadas = periodo1 + periodo2
    return horas_trabalhadas


def salvartabela():
    global df
    entrada1 = acoes_horarios["Entrada1"]
    saida1 = acoes_horarios["Saida1"]
    entrada2 = acoes_horarios["Entrada2"]
    saida2 = acoes_horarios["Saida2"]

    horas_trabalhadas = calcular_horas_trabalhadas(entrada1, saida1, entrada2, saida2)

    novo_registro = {
        "Data": date.today().strftime("%d/%m/%Y"),
        "Entrada1": entrada1,
        "Saida1": saida1,
        "Entrada2": entrada2,
        "Saida2": saida2,
        "Horas Trabalhadas": str(horas_trabalhadas)
    }

    df = pd.concat([df, pd.DataFrame([novo_registro])], ignore_index=True)
    df.to_excel('tabelaponto.xlsx', index=False)
    print(f"Novo registro na tabela: {novo_registro}")


def _esperar_botao_carregar(identificador):
    return WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, identificador)))


def registro_populis(tipo):
    driver.get(r"https://dell.populisservicos.com.br/populisII-web/paginas/publicas/login.xhtml")

    login = _esperar_botao_carregar('//*[@id="formLogin:usuarioInput"]')
    login.send_keys(cpf)

    senha = _esperar_botao_carregar('//*[@id="formLogin:senhaInput"]')
    senha.send_keys(password)

    bt_login = _esperar_botao_carregar('//*[@id="formLogin:btnLogin"]/span')
    bt_login.click()

    bt_marcacoes = _esperar_botao_carregar('/html/body/div[2]/div[1]/div[1]/div[1]')
    bt_marcacoes.click()
    sl(1)

    if tipo == "E":
        bt_entrada = _esperar_botao_carregar('//*[@id="marcacaoEntrarBtn"]')
        bt_entrada.click()

    if tipo == "S":
        bt_saida = _esperar_botao_carregar('//*[@id="marcacaoSairBtn"]')
        bt_saida.click()


def executa_acao(nome_acao):
    global contador_acoes
    if nome_acao in acoes_horarios:
        horario = acoes_horarios[nome_acao]
        tipo = "E" if "Entrada" in nome_acao else "S"
        registro_populis(tipo)
        print(f"Novo registro no populis Horário de {nome_acao}: {horario}")
        contador_acoes += 1


def agendar_acoes(acoes):
    for nome_acao, horario in acoes.items():
        hora, minuto = horario.split(":")
        schedule.every().day.at(f"{hora}:{minuto}").do(lambda n=nome_acao: executa_acao(n))


# Chamada de funcoes
agendar_acoes(acoes_horarios)
salvartabela()

# Verificar horarios
while contador_acoes < 4:
    schedule.run_pending()
    time.sleep(1)

print("Todas as ações foram executadas. Finalizando o script.")
driver.quit()
