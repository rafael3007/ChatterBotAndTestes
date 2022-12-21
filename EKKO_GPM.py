##########     EKKO_GPM V.0.1     ###########
#      ELABORADO POR FLÁVIO FONSECA        #
############################################

from time import sleep
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pandas as pd
from datetime import date, timedelta
import datetime
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def getRelatorios():
    ############## DEFINIR OS RELATORIOS COMO 1 PARA O DOWNLOAD SER EXECUTADO ###############
    baixar_servico = 1
    baixar_turno = 1
    baixar_faturamento = 1

    #########################################################################################

    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)

    service = Service(ChromeDriverManager().install())
    browser = webdriver.Chrome(service=service)

    browser.get('https://ecoeletrica.gpm.srv.br/index.php')
    browser.maximize_window()

    ################################# DATAS ##############################

    today_date = date.today()

    gap_servico = timedelta(-5)
    gap_faturamento = timedelta(-10)
    gap_turno = timedelta(-2)
    dia_anterior = timedelta(-1)

    #DATA RELATÓRIO DE FATURAMENTO

    data_inicio_faturamento = datetime.datetime.strptime(str(gap_faturamento + today_date), "%Y-%m-%d").strftime("%d-%m-%Y")
    data_final_faturamento = datetime.datetime.strptime(str(dia_anterior + today_date), "%Y-%m-%d").strftime("%d-%m-%Y")

    #DATA RELATÓRIO DE SERVIÇO

    data_inicio_servico = datetime.datetime.strptime(str(gap_servico + today_date), "%Y-%m-%d").strftime("%d-%m-%Y")
    data_final_servico = datetime.datetime.strptime(str(dia_anterior + today_date), "%Y-%m-%d").strftime("%d-%m-%Y")

    #DATA RELATÓRIO DE TURNO

    data_inicio_turno = datetime.datetime.strptime(str(gap_turno + today_date), "%Y-%m-%d").strftime("%d-%m-%Y")
    data_final_turno = datetime.datetime.strptime(str(dia_anterior + today_date), "%Y-%m-%d").strftime("%d-%m-%Y")

    #LOGIN

    #WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="idLogin"]')))
    browser.find_element('xpath','//*[@id="idLogin"]').send_keys("RAFAEL.SAMPAIO")
    browser.find_element('xpath','//*[@id="idSenha"]').send_keys("ECO@2012")
    browser.find_element('xpath','//*[@id="form_login"]/input[5]').click()

    #########################################################################################

        ########################################################################
        #                    ENTRANDO NA PÁGINA DE SERVIÇOS                    #
        ########################################################################

    print("Estou sendo executado!")



    if baixar_servico == 1:
        browser.get('https://ecoeletrica.gpm.srv.br/gpm/geral/consulta_servico.php')
        #INSERINDO DATA
        WebDriverWait(browser, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="data_inicial"]')))
        browser.find_element('xpath','//*[@id="data_inicial"]').click()
        browser.find_element('xpath','//*[@id="data_inicial"]').send_keys(data_inicio_servico)
        WebDriverWait(browser, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="data_final"]')))
        browser.find_element('xpath','//*[@id="data_final"]').click()
        browser.find_element('xpath','//*[@id="data_final"]').send_keys(data_final_servico)
        #ESCOLHENDO COORDENADOR
        WebDriverWait(browser, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form"]/table/tbody/tr[6]/td[4]/div')))
        browser.find_element('xpath','//*[@id="form"]/table/tbody/tr[6]/td[4]/div').click()
        browser.find_element('xpath','//*[@id="form"]/table/tbody/tr[6]/td[4]/div/input').send_keys("ISRAEL IURY JACINTO DE SOUZA - ECM427527")
        browser.find_element('xpath','//*[@id="form"]/table/tbody/tr[6]/td[4]/div').click()
        WebDriverWait(browser, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form"]/table/tbody/tr[6]/td[4]/div/div[2]/div[2]')))
        browser.find_element('xpath','//*[@id="form"]/table/tbody/tr[6]/td[4]/div/div[2]/div[2]').click()
        #PESQUISAR
        WebDriverWait(browser, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form"]/div/input')))
        browser.find_element('xpath','//*[@id="form"]/div/input').click()
        #EXPORTANDO PARA EXCEL
        WebDriverWait(browser, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form"]/div[2]/span/a/button')))
        browser.find_element('xpath','//*[@id="form"]/div[2]/span/a/button').click()
        print("Baixei Serviços!")

        ########################################################################
        #                  ENTRANDO NA PÁGINA DE FATURAMENTO                   #
        ########################################################################

    if baixar_faturamento == 1:
        browser.get('https://ecoeletrica.gpm.srv.br/gpm/geral/grafico_mov_fatur_obra.php')
        #INSERINDO DATA
        WebDriverWait(browser, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="data_inicial"]')))
        browser.find_element('xpath','//*[@id="data_inicial"]').click()
        browser.find_element('xpath','//*[@id="data_inicial"]').send_keys(data_inicio_faturamento)
        WebDriverWait(browser, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="data_final"]')))
        browser.find_element('xpath','//*[@id="data_final"]').click()
        browser.find_element('xpath','//*[@id="data_final"]').send_keys(data_final_faturamento)
        #CLICANDO EM DETALHADO
        browser.find_element('xpath','//*[@id="id_form"]/div/table/tbody/tr[2]/td[2]/label[2]/input').click()
        #CLICANDO NOS CONTRATOS
        WebDriverWait(browser, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="contrato_chosen"]')))
        browser.find_element('xpath','//*[@id="contrato_chosen"]').click()
        browser.find_element('xpath','//*[@id="contrato_chosen"]/ul/li/input').send_keys("MANUTENÇÃO - COELBA")
        WebDriverWait(browser, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="contrato_chosen"]/div')))
        browser.find_element('xpath','//*[@id="contrato_chosen"]/div').click()
        browser.find_element('xpath','//*[@id="contrato_chosen"]').click()
        browser.find_element('xpath','//*[@id="contrato_chosen"]/ul/li/input').send_keys("CONSTRUÇÃO - COELBA")
        WebDriverWait(browser, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="contrato_chosen"]/div')))
        browser.find_element('xpath','//*[@id="contrato_chosen"]/div').click()
        #PESQUISAR
        WebDriverWait(browser, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="id_form"]/div/div/input')))
        browser.find_element('xpath','//*[@id="id_form"]/div/div/input').click()
        #EXPORTANDO PARA EXCEL
        WebDriverWait(browser, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="id_form"]/div[3]/div/span/span/span/input')))
        browser.find_element('xpath','//*[@id="id_form"]/div[3]/div/span/span/span/input').click()
        print("Baixei Faturamento!")

        ########################################################################
        #                      ENTRANDO NA PÁGINA DE TURNO                     #
        ########################################################################

    if baixar_turno == 1:
        browser.get('https://ecoeletrica.gpm.srv.br/gpm/geral/consulta_turno.php')
        #INSERINDO DATA
        WebDriverWait(browser, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="data_inicial"]')))
        browser.find_element('xpath','//*[@id="data_inicial"]').click()
        browser.find_element('xpath','//*[@id="data_inicial"]').send_keys(data_inicio_turno)
        WebDriverWait(browser, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="data_final"]')))
        browser.find_element('xpath','//*[@id="data_final"]').click()
        browser.find_element('xpath','//*[@id="data_final"]').send_keys(data_final_turno)

        #ESCOLHENDO COORDENADOR
        WebDriverWait(browser, 60).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form[4]/table/tbody/tr[4]/td[2]/div')))
        browser.find_element('xpath','/html/body/form[4]/table/tbody/tr[4]/td[2]/div').click()
        WebDriverWait(browser, 60).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form[4]/table/tbody/tr[4]/td[2]/div/input')))
        browser.find_element('xpath','/html/body/form[4]/table/tbody/tr[4]/td[2]/div/input').send_keys("ISRAEL IURY JACINTO DE SOUZA - ECM427527")
        browser.find_element('xpath','/html/body/form[4]/table/tbody/tr[4]/td[2]/div').click()
        WebDriverWait(browser, 60).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form[4]/table/tbody/tr[4]/td[2]/div/div[2]')))
        browser.find_element('xpath','/html/body/form[4]/table/tbody/tr[4]/td[2]/div/div[2]').click()
        #PESQUISAR
        WebDriverWait(browser, 60).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form[4]/div/input')))
        browser.find_element('xpath','/html/body/form[4]/div/input').click()
        #EXPORTANDO PARA EXCEL

        #/html/body/form[4]/div[1]/img
        WebDriverWait(browser, 60).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form[4]/div[4]/a[1]/img')))
        browser.find_element('xpath','/html/body/form[4]/div[4]/a[1]/img').click()
        print("Baixei Turnos!")

    while True:
        sleep(1)

if __name__ == "__main__":
    getRelatorios()