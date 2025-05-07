from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import flet as ft
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

pasta_raiz = os.path.abspath(os.getcwd())
diretorio_downloads = os.path.join(pasta_raiz)

def tela_login(page: ft.Page):
    page.title = "Login Futurista"
    page.bgcolor = "#0a0a0a"  # Fundo escuro
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    def realizar_login(e):
        #usuario = campo_usuario.value
        #senha = campo_senha.value
        mensagem.value = "Realizando login..."
        mensagem.color = "yellow"
        page.update()

        # Iniciar Selenium
        # service = Service(ChromeDriverManager().install())
        service = Service(executable_path='chromedriver.exe')
        options = webdriver.ChromeOptions()
        #options.add_argument("--headless")  # Executa o Chrome em segundo plano (modo headless)
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        prefs = {
            "download.default_directory": diretorio_downloads,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_setting_values.automatic_downloads": 1  # Permite múltiplos downloads
        }
        options.add_experimental_option("prefs", prefs)
        global navegador  # Declare navegador como global aqui
        navegador = webdriver.Chrome(service=service, options=options)
        navegador.set_window_size(1600, 900)
        navegador.get(r"https:\\dwbi.correios.com.br\MicroStrategy\servlet\mstrWeb")

        try:
            navegador.find_element(By.XPATH, '//*[@id="Uid"]').send_keys('84198842')

            navegador.find_element(By.XPATH, '//*[@id="Pwd"]').send_keys('19041982')

            # Clica no Botão entrar
            navegador.find_element(By.XPATH, '//*[@id="3054"]').click()

            navegador.find_element(By.XPATH,
                                   '//*[@id="projects_ProjectsStyle"]/table/tbody/tr[18]/td[2]/div/table/tbody/tr/td[2]/a').click()

            # Meus relatórios

            WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH,
                                                                           '/html/body/div[4]/table/tbody/tr[2]/td[1]/div/div/div/div[3]/div[1]/div/div[1]/ul/div[2]/li/span/a/span'))).click()

###################

            #####################################------------      PACN     -----------------###################
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            import time

            # PACN
            # Meus relatórios
            # -- navegador.find_element(By.XPATH,'/html/body/div[4]/table/tbody/tr[2]/td[1]/div/div/div/div[3]/div[1]/div/div[1]/ul/div[2]/li/span/a/span').click()

            # WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, ''))).click()
            # PACN
            WebDriverWait(navegador, 20).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="FolderIcons"]/tbody/tr[9]/td[2]/div/table/tbody/tr/td[2]/a'))).click()
            # navegador.find_element(By.XPATH,'//*[@id="FolderIcons"]/tbody/tr[9]/td[2]/div/table/tbody/tr/td[2]/a').click()

            # navegador.find_element(By.XPATH,'//*[@id="id_mstr54"]/table/tbody/tr[5]/td[2]').click() # 5 - selecione Familia
            WebDriverWait(navegador, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="id_mstr54"]/table/tbody/tr[5]/td[2]'))).click()

            # <<
            WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH,
                                                                           '/html/body/div[2]/table/tbody/tr[2]/td[2]/div[2]/div[1]/div/table/tbody/tr[1]/td[2]/div/div/div[6]/span/div[2]/div[1]/table/tbody/tr[3]/td[2]/div[5]/div/img'))).click()
            # navegador.find_element(By.XPATH, '//*[@id="id_mstr222"]/img').click()

            # escolha de sedex
            WebDriverWait(navegador, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="id_mstr217ListContainer"]/div[6]/div'))).click()
            # navegador.find_element(By.XPATH, '//*[@id="id_mstr217ListContainer"]/div[6]/div').click()

            # >
            WebDriverWait(navegador, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="id_mstr219"]/img'))).click()
            # navegador.find_element(By.XPATH,'//*[@id="id_mstr219"]/img').click()

            # campo data
            data = '12/03/2025'

            time.sleep(2)

            try:
                # Espera até que o elemento esteja presente
                element = WebDriverWait(navegador, 20).until(
                    EC.presence_of_element_located((By.XPATH,
                                                    '/html/body/div[2]/table/tbody/tr[2]/td[2]/div[2]/div[1]/div/table/tbody/tr[1]/td[2]/div/div/div[6]/span/div[2]/div[1]/table/tbody/tr[2]/td/div/table/tbody/tr/td[1]/div/table/tbody/tr/td[1]/div/input'))
                )
                # Clica no elemento
                element.click()
                # Insere a data
                element.send_keys(data)
                print("Data inserida com sucesso!")
            except TimeoutException:
                print(
                    "Elemento não encontrado dentro do tempo limite.")  # navegador.find_element(By.XPATH,'//*[@id="id_mstr252_txt"]').send_keys(data)

            time.sleep(2)
            # Pesquisar a data
            try:
                # Espera até que o elemento esteja presente
                element = WebDriverWait(navegador, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="id_mstr252"]/table/tbody/tr/td[2]/div/input'))
                )
                # Clica no elemento
                element.click()
                # Insere a data

                print("Clicou no botão pesquisar com Sucesso!")
            except TimeoutException:
                print(
                    "Elemento não encontrado dentro do tempo limite.")  # navegador.find_element(By.XPATH,'//*[@id="id_mstr252_txt"]').send_keys(data)
            # WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="id_mstr252"]/table/tbody/tr/td[2]/div/input'))).click()
            # navegador.find_element(By.XPATH,'//*[@id="id_mstr252"]/table/tbody/tr/td[2]/div/input').click()

            # >>
            WebDriverWait(navegador, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="id_mstr257"]/img'))).click()
            # navegador.find_element(By.XPATH,'//*[@id="id_mstr257"]/img').click()

            # Executar Relatorio
            WebDriverWait(navegador, 40).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="id_mstr227"]'))).click()
            # navegador.find_element(By.XPATH,'//*[@id="id_mstr227"]').click()

            # Esperar o objeto aparecer
            WebDriverWait(navegador, 40).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tbExport"]'))).click()
            # navegador.find_element(By.XPATH,'//*[@id="tbExport"]').click()

            # Exportar
            # WebDriverWait(navegador, 40).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="3131"]'))).click()
            # navegador.find_element(By.XPATH,'//*[@id="3131"]').click()
            navegador.switch_to.window(navegador.window_handles[-1])
            try:
                # Espera até que o elemento esteja presente
                element = WebDriverWait(navegador, 58).until(
                    EC.presence_of_element_located((By.ID, 3131))
                )
                # Clica no elemento
                element.click()
                # Insere a data

                print("Sucesso!")
            except TimeoutException:
                print("Elemento não encontrado dentro do tempo limite.")
                # navegador.find_element(By.XPATH,'//*[@id="id_mstr252_txt"]').send_keys(data)

            time.sleep(1)
            # Logo - inicio
            WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mstrLogo"]'))).click()
            # navegador.find_element(By.XPATH,'//*[@id="mstrLogo"]').click()

            ###################





            # Esperar a página carregar e verificar login
            WebDriverWait(navegador, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="menu"]/a[1]')))
            mensagem.value = "✅ Login realizado com sucesso!"
            mensagem.color = "green"
            page.update()
            time.sleep(2)
            #navegador.quit()
        except:
            mensagem.value = "❌ Erro no login! Verifique suas credenciais."
            mensagem.color = "red"
            page.update()
            navegador.quit()

    robo_imagem = ft.Image(src="https://media.istockphoto.com/id/1958063552/pt/foto/digital-ai-newsletter-concept-ai-powered-email-marketing-automation-marketing-automation-email.jpg?s=1024x1024&w=is&k=20&c=WAhfwZu2XNWU8zDfyOm2RPeAkQMxCoCptKpY2ASbIP4=", width=150, height=150)
    titulo = ft.Text("Acesso Robótico DW", size=24, color="#00ffcc", weight=ft.FontWeight.BOLD)
    #campo_usuario = ft.TextField(label="Usuário", width=300, bgcolor="#1a1a1a", color="white")
    #campo_senha = ft.TextField(label="Senha", password=True, width=300, bgcolor="#1a1a1a", color="white")
    botao_login = ft.ElevatedButton("Entrar", on_click=realizar_login, bgcolor="#00ffcc", color="#0a0a0a")
    mensagem = ft.Text(value="", size=16, weight=ft.FontWeight.BOLD)

    page.add(
        ft.Column([
            robo_imagem,
            titulo,
            #campo_usuario,
           #campo_senha,
            botao_login,
            mensagem
        ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER)
    )




if __name__ == "__main__":
    ft.app(target=tela_login)

