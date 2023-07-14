import time
from selenium import webdriver
import datetime
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import Select
import openpyxl
from selenium.webdriver import ActionChains
import os
import shutil
import locale

locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')

# Formatação de datas
hoje = datetime.date.today()
hoje_formatado = hoje.strftime('%d/%m/%Y')
hoje_extenso = hoje.strftime('%d de %B de %Y')
ontem = hoje - datetime.timedelta(days=1)
ontem_formatado = ontem.strftime('%d/%m/%Y')
ontem_extenso = ontem.strftime('%d de %B de %Y')
ultima_semana = hoje - datetime.timedelta(days=7)
semana_formatada = ultima_semana.strftime('%d/%m/%Y')
ontem_dia = ontem.strftime('%d')
ontem_mes = ontem.strftime('%m')
ontem_ano = ontem.strftime('%Y')

navegador = webdriver.Chrome()
navegador.maximize_window()

# Abrir Página do Sistema
navegador.get('https://gps.performance-al.com.br/login/login.xhtml;JSESSIONID=8234d2cf-bf9f-45c6-8f05-3378b158a217')

# Inserir Credenciais e logar
usuario = navegador.find_element(By.XPATH, '//*[@id="j_username"]')
usuario.send_keys('aalmeida')
senha = navegador.find_element(By.XPATH, '//*[@id="j_password"]')
senha.send_keys('Endicon2023.', Keys.TAB, Keys.ENTER)
time.sleep(1)


# Acessar a aba de notificações
navegador.get('https://gps.performance-al.com.br/secure/manager/groups/#/')
time.sleep(1)

# Escolher grupo
escolher_grupo = navegador.find_element(By.XPATH, '/html/body/div[3]/div[2]/div[2]/input')
#escolher_grupo.send_keys('2.3 EQUAT NORDESTE - PARAGOMINAS')
#escolher_grupo.send_keys('2.2 EQUAT NORDESTE - CAPANEMA')
#escolher_grupo.send_keys('2.4 EQUAT NORDESTE - ABAETETUBA')
#escolher_grupo.send_keys('2.1 EQUAT NORDESTE - CASTANHAL')
escolher_grupo.send_keys('3.1 CELPE - RECIFE')
#escolher_grupo.send_keys('3.2 CELPE - CARPINA')
time.sleep(1)

botao_editar = navegador.find_element(By.XPATH, '/html/body/div[3]/div[2]/div[3]/table/tbody/tr/td[5]/a')
botao_editar.click()
time.sleep(1)

selecionar_todos = navegador.find_element(By.XPATH, '/html/body/div[3]/div[2]/div[2]/div[1]/div[3]/table/thead/tr/th[1]/input')
selecionar_todos.click()
excluir_todos = navegador.find_element(By.XPATH, '/html/body/div[3]/div[2]/div[2]/div[1]/div[3]/table/thead/tr/th[3]/button')
excluir_todos.click()
time.sleep(5)
navegador.refresh()
time.sleep(2)

#aba_veiculos = navegador.find_element_by_xpath('/html/body/div[3]/div[2]/div[2]/div[1]')

# 3. Acessar Planilha de Cadastro
workbook = openpyxl.load_workbook('C:\\temp\\atualizar_cadastro.xlsx')
planilha = workbook.active

# 4. Criar iteração do cadastros
for coluna in planilha.iter_rows(min_row=2, values_only=True):
    veiculo = coluna[0]

    buscar_veiculos = navegador.find_element(By.XPATH, '/html/body/div[3]/div[2]/div[2]/div[1]/div[2]/div/div[1]/input')
    buscar_veiculos.send_keys(veiculo)
    buscar_veiculos.send_keys(Keys.ENTER)
    time.sleep(2)

    print('veículo ', veiculo, 'incluido com sucesso.')


incluir = navegador.find_element(By.XPATH, '/html/body/div[3]/div[2]/div[2]/div[1]/div[2]/button')
incluir.click()
time.sleep(10)




