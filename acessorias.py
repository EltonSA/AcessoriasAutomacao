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

            #Clica no campo de busca da empresa
            #//*[@id='divEmpresas']/div[1]
            campo_empresaId = driver.find_element(By.XPATH, "//*[@id='divEmpresas']/div[1]")
            campo_empresaId.click()
            time.sleep(2)

            #Entra dentro do contato na empresa
            campo_contato1 = driver.find_element(By.XPATH, "//*[@id='CttEdit_0_1']/button[1]")
            campo_contato1.click()
            time.sleep(2)
            #Clica para editar o contato
            #//*[@id="divCtt_0_1"]/div[1]/div/span[1]/button
            campo_contato2 = driver.find_element(By.XPATH, "//*[@id='divCtt_0_1']/div[1]/div/span[1]/button")
            campo_contato2.click()
            time.sleep(2)

            #Marca a caixa de todos os departamentos
            #//*[@id="selDpZ_0_1"]/a[1]
            campo_contato3 = driver.find_element(By.XPATH, "//*[@id='selDpZ_0_1']/a[1]")
            campo_contato3.click()
            time.sleep(2)

            #clica em salvar as alterações
            #//*[@id="CttSave_0_1"]/button[1]
            campo_salvar = driver.find_element(By.XPATH, "//*[@id='CttSave_0_1']/button[1]")
            campo_salvar.click()
            time.sleep(5)

            #clica em voltar para o campor de busca
            #//*[@id="navFim"]/div[2]/button[2]
            campo_voltar = driver.find_element(By.XPATH, "//*[@id='navFim']/div[2]/button[2]")
            campo_voltar.click()
            time.sleep(5)


    def carregar_cnpjs(self, caminho_planilha):
        wb = openpyxl.load_workbook(caminho_planilha)
        ws = wb.active

        cnpjs = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            cnpj = row[0]  # Supondo que a primeira coluna contenha os CNPJs
            cnpjs.append(cnpj)

        return cnpjs
        
# Criar uma instância da classe AcessoriasBoot
boot = AcessoriasBoot('contabilidade_ramos@hotmail.com', 'fiscal11')
#boot = AcessoriasBoot('usuario', 'senha')
boot.login()