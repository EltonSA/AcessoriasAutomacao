from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import openpyxl

class AcessoriasBoot:
    def __init__(self, usuario, senha):
        self.usuario = usuario
        self.senha = senha
        self.driver = webdriver.Firefox()

#inicia o login no site da acessorias
    def login(self):
        driver = self.driver
        driver.get("https://app.acessorias.com")

        #login
        campo_user = driver.find_element(By.XPATH, "//input[@name='mailAC']")
        campo_user.clear()
        campo_user.send_keys(self.usuario)
        time.sleep(2)

        campo_senha = driver.find_element(By.XPATH, "//input[@name='passAC']")
        campo_senha.clear()
        campo_senha.send_keys(self.senha)
        campo_senha.send_keys(Keys.RETURN)

        # Aguarde alguns segundos para permitir que a página carregue após o login
        time.sleep(2)

        # Clica na aba empresa
        campo_empresa = driver.find_element(By.XPATH, "//a[@href='sysmain.php?m=4']")
        campo_empresa.click()
        time.sleep(2)

        # Loop para buscar e interagir com empresas
        empresas = self.carregar_cnpjs("C:\\Users\\elton.santos\\Desktop\\projetos\\AutomacaoAcessorias\\empresas.xlsx")

        for cnpj in empresas:
            campo_busca = driver.find_element(By.XPATH, "//input[@name='searchString']")
            campo_busca.clear()
            campo_busca.send_keys(cnpj)
            print(cnpj)
            campo_busca.send_keys(Keys.RETURN)
            time.sleep(2)

            #//*[@id='divEmpresas']/div[1]
            campo_empresaId = driver.find_element(By.XPATH, "//*[@id='divEmpresas']/div[1]")
            campo_empresaId.click()
            time.sleep(2)

            clicks1 = ["//*[@id='CttEdit_0_1']/button[1]", "//*[@id='CttEdit_0_2']/button[1]", "//*[@id='CttEdit_0_3']/button[1]", "//*[@id='CttEdit_0_4']/button[1]", "//*[@id='CttEdit_0_5']/button[1]"]
            clicks2 = ["//*[@id='divCtt_0_1']/div[1]/div/span[1]/button", "//*[@id='divCtt_0_2']/div[1]/div/span[1]/button", "//*[@id='divCtt_0_3']/div[1]/div/span[1]/button", "//*[@id='divCtt_0_4']/div[1]/div/span[1]/button", "//*[@id='divCtt_0_5']/div[1]/div/span[1]/button"]
            clicks3 = ["//*[@id='selDpZ_0_1']/a[1]", "//*[@id='selDpZ_0_2']/a[1]", "//*[@id='selDpZ_0_3']/a[1]", "//*[@id='selDpZ_0_4']/a[1]" ,"//*[@id='selDpZ_0_5']/a[1]"]
            clicks4 = ["//*[@id='CttSave_0_1']/button[1]", "//*[@id='CttSave_0_2']/button[1]", "//*[@id='CttSave_0_3']/button[1]", "//*[@id='CttSave_0_4']/button[1]", "//*[@id='CttSave_0_5']/button[1]"]
            
            #Clica no icone para editar o contato
            for x in clicks1:
                try:
                    click01 = driver.find_element(By.XPATH, x)
                    click01.click()
                    time.sleep(1)
                except:
                    print("Abrindo os departamentos")
            
            #Clica no chuveirinho para abrir os departamentos
            for y in clicks2:
                try:
                    click02 = driver.find_element(By.XPATH, y)
                    click02.click()
                    time.sleep(1)
                except:
                    print("Marca os a caixa de todos os departamentos")

            #Marca todos os departamentos dos contatos        
            for z in clicks3:
                try:
                    click03 = driver.find_element(By.XPATH, z)
                    click03.click()
                    time.sleep(3)
                except:
                    print("Clica em salvar os dados")

            #Clica em salvar os dados    
            for s in clicks4:
                try:
                    click04 = driver.find_element(By.XPATH, s)
                    click04.click()
                    time.sleep(3)
                except:
                    print("Sainda para buscar a nova empresa")
                    #clica em voltar para o campor de busca
            campo_voltar = driver.find_element(By.XPATH, "//*[@id='navFim']/div[2]/button[2]")
            campo_voltar.click()
            time.sleep(4)  


    def carregar_cnpjs(self, caminho_planilha):
        wb = openpyxl.load_workbook(caminho_planilha)
        ws = wb.active

        cnpjs = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            cnpj = row[0]  # Supondo que a primeira coluna contenha os CNPJs
            cnpjs.append(cnpj)

        return cnpjs
        
# Criar uma instância da classe AcessoriasBoot
boot = AcessoriasBoot('Usuario', 'Senha')
boot.login()