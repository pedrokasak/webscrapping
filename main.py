from selenium import webdriver
from selenium.webdriver.chrome.options import  Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.expected_conditions import _find_element
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import *
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import openpyxl
import smtplib
import os
from email.message import EmailMessage
import re

class Scrappy :
# Configurando Chrome
        

    def iniciar(self) :
        self.informe_email_senha()
        self.coleta_dados_site()
        self.criar_planilha_excel()
        self.enviar_planilha_email()

# informar email para para o envio de informações
    def informe_email_senha(self):
        self.email = input('INFORME O SEU EMAIL PARA RECEBER O SEU RELATÓRIO DO SITE\n')
        self.email.lower()
        self.senha = input('DIGITE A SUA SENHA\n')
        padrao = re.search(
            r"^[a-zA-Z0-9._-]+@[a-zA-Z0-9]+\.[a-zA-Z\.a-zA-Z]{1,6}$",self.email)
        if padrao:
            print('Email Válido')
        else:
            print('Informe um email Válido')
            self.informe_email_senha()
# acessando o site para coleta de dados
    def coleta_dados_site(self) :
        chrome_options = Options()
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_options.add_argument('--lang=pr-BR')        
        chrome_options.add_argument('--disable-notifications')
        self.driver = webdriver.Chrome(ChromeDriverManager().install())
        self.driver.set_window_size(800,700)
        self.link = 'https://www.buscape.com.br/'        
        print(self.driver.title)
        self.lista_titulo_celular = []
        self.lista_nome_celulares = []
        self.lista_preco_celulares = []
        self.driver.get(self.link)
        celular = self.driver.find_element_by_xpath("//a[@class='HotLinks_Label__3VWTs']")
        sleep(2)
        print(celular)

# acessando o site para coleta de dados
        for p in range(15):
            item = 1
            lista_nomes = self.driver.find_element_by_xpath('//*[@id="resultArea"]/div[2]/div/ul/li[1]/article/span/a/span/span[2]/span[1]') 
            self.lista_nome_celulares.append(lista_nomes[0].text)
            sleep(2)
            lista_precos = self.driver.find_element_by_xpath('//*[@id="resultArea"]/div[2]/div/ul/li[1]/article/span/a/span/span[2]/span[2]/span[1]')
            self.lista_preco_celulares.append(lista_precos[0].text)
            item += 1
            sleep(2)

            try:
                botao_proximo = self.driver.find_element_by_xpath('//*[@id="resultArea"]/div[4]/div/ul/li[9]/a')
                botao_proximo.click()
                print(f'\u001b[32m{"Navegando para Proxima Página"}\u001b[0m')
                sleep(2)

            except NoSuchElementException :
                print(f'\u001b[33m{"Naão há mais Páginas"}\u001b[0m')
                print(f'\u001b[32m{"Escaneamento concluído"}\u001b[0m')
                self.driver.quit()

# criando uma planilha no excel
    def criar_planilha_excel(self):
        index= 2
        planilha = openpyxl.workbook()
        celulares = planilha['Sheet']
        celulares.title = 'Celulares'
        celulares['A1'] = 'Nome'
        celulares['B1'] = 'Preço'

        for nome, preco in zip(self.lista_nome_celulares,self.lista_preco_celulares):
            celulares.cell(column=1, row=index, value=nome)
            celulares.cell(column=2, row=index, value=nome)
            index +=1

        planilha.save("planilha_de_preços.xlsx")
        print(f'\u001b[32m{"Planilha gerada com Sucesso"}\u001b[0m')

# enviando a planilha por email
    def enviar_planilha_email(self,endereco,senha):

        msg = EmailMessage()
        msg['Subject'] = 'Planilha de Preços de celulares e Smartphones'
        msg['From'] = endereco
        msg['To'] = self.email
        msg.set_content('olá sua Planilha Chegou') 
        arquivos = ["planilha_de_preços.xlsx"]
        for arquivo in arquivos:
            with open(arquivo,'rb') as arq:
                dados = arq.read()
                nome_arquivo = arq.name
            msg.add_attachment(dados,maintype='application', subtype='octet-stream', filename=nome_arquivo) 

        server = smtplib.SMTP('imap.gmail.com',port=587)  
        server.ehlo() 
        server.starttls()
        server.login(endereco, senha, initial_response_ok=True)    
        server.send_message(msg)
        print(f'\u001b[32m{"Email Enviado com Sucesso"}\u001b[0m')

        server.quit()

start = Scrappy()
start.iniciar()
