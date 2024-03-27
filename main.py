from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import pyodbc
import openpyxl
import shutil
from datetime import datetime
import datetime

def conexao_banco(driver='SQL Server', server='NBSTI-003', database='automacao', username='marcos', passowrd='Marcos@2024', trusted_connection='yes'):
    string_conexao = f"DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={passowrd};TRUSTED_CONNECTION={trusted_connection}"
    conexao = pyodbc.connect(string_conexao)
    cursor = conexao.cursor()
    return conexao, cursor
conexao, cursor = conexao_banco()

query_insert = "INSERT INTO TESTE_C(NUM_PARC, ATENDENTE, STATUS_ATEND, INI_ATEND, FIM_ATEND) VALUES (?, ?, ?, ?, ?)"
query_insert_t = "INSERT INTO TESTE_C(NUM_PARC, ATENDENTE, STATUS_ATEND, INI_ATEND) VALUES (?, ?, ?, ?)"

#ABRIR NAVEGADOR E FAZER DOWNLOAD DO XLSX

navegador = webdriver.Chrome()

navegador.get("xxxx")
campo_mail = navegador.find_element(by=By.XPATH, value='//*[@id="inputEmail"]')
campo_mail.send_keys('xxxx')
campo_pssws = navegador.find_element(by=By.XPATH, value='//*[@id="inputPassword"]')
campo_pssws.send_keys('xxx')

login = navegador.find_element(by=By.XPATH, value='//*[@id="submit"]').click()
sleep(5)
navegador.get("https://ipchat.com.br/dashboard")
sleep(5)
excel = navegador.find_element(by=By.XPATH, value='//*[@id="example_wrapper"]/div[1]/button[1]').click()
sleep(5)

#MOVE XLSX PARA ARQUIVO DO PROJETO 

origem = r"C:\Users\marcos.silva\Downloads\xlsx.xlsx"
destino =r"C:\Users\marcos.silva\Documents\project_mdk"

shutil.move(origem, destino)


#LEITURA DO XLSX E INSERT NO BANCO

xlsx = openpyxl.load_workbook('xlsx.xlsx')
sheet_dados = xlsx['Sheet1']

for linha in sheet_dados.iter_rows(min_row=3):

    if linha[0].value is not None:
        num_parceiro = linha[0].value
        atendente = linha[1].value
        status = linha[2].value
        ini_atend =linha[4].value
        fim_atend =linha[5].value

        if ini_atend != '0000-00-00 00:00:00':
           fim_atend =datetime.datetime.strptime(linha[4].value, '%Y-%m-%d %H:%M:%S')
        else:
            fim_atend = datetime.datetime.strptime('2000-01-01 01:01:01', '%Y-%m-%d %H:%M:%S')

        if fim_atend != '0000-00-00 00:00:00':
           fim_atend =datetime.datetime.strptime(linha[5].value, '%Y-%m-%d %H:%M:%S')
        else:
            fim_atend = datetime.datetime.strptime('2000-01-01 01:01:01', '%Y-%m-%d %H:%M:%S')

    print(fim_atend)
