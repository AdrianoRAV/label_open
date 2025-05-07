import os
import time
import flet as ft
import base64
import os
import time
import shutil
import pandas as pd
import flet as ft
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import fitz
import re
#from Excel.SRO_OPEN.main import usuario
#from Excel.app import usuario
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from utils_pdf import diretorio_downloads, pdf_to_excelucub
#from main import user_mat
complementar = '*'
#user_mat = usuario
def deletar_arquivo(nome_arquivo):
    # Verifique se o arquivo existe e então o delete
    if os.path.isfile(nome_arquivo):
        os.remove(nome_arquivo)
    else:
        # Informe o usuário se o arquivo não existir
        print(f"Erro: arquivo {nome_arquivo} não encontrado")

def esperar_download(diretorio, nome_arquivo):
    caminho_arquivo = os.path.join(diretorio, nome_arquivo)
    while not os.path.exists(caminho_arquivo):
        time.sleep(1)  # Espera 1 segundo antes de verificar novamente
    # Adicionalmente, você pode verificar se o arquivo ainda está em uso
    while True:
        try:
            with open(caminho_arquivo, 'rb'):
                break
        except IOError:
            time.sleep(1)

def mensagem_sucesso(nome):
    texto = ft.Text(
        nome,
        color=ft.colors.GREEN,
        size=20,
        weight=ft.FontWeight.BOLD,
        italic=True
    )
    return texto


def esperar_download_arquivo(caminho_arquivo, timeout=60):
    """
    Espera até que o arquivo seja encontrado no diretório especificado.
    :param caminho_arquivo: Caminho completo para o arquivo esperado.
    :param timeout: Tempo máximo para esperar pelo download (em segundos).
    """
    start_time = time.time()
    while True:
        if os.path.exists(caminho_arquivo):
            print(f"Arquivo {caminho_arquivo} encontrado.")
            return True
        elif time.time() - start_time > timeout:
            print("Tempo limite atingido para o download do arquivo.")
            return False
        time.sleep(1)  # Espera um segundo antes de verificar novamente


def iniciar_navegador(usuario, senha):
    global wait
    try:
        #service = Service(ChromeDriverManager().install())
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
        navegador.get("https://sroweb.correios.com.br")
        # Aguarda até que o campo de username esteja disponível para interagir
        #                                   Login SRO                                                    #
        wait = WebDriverWait(navegador, 10)
        wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="username"]'))).send_keys(usuario)
        navegador.find_element(By.XPATH, '//*[@id="password"]').send_keys(senha)
        navegador.find_element(By.NAME, 'submitBtn').click()

        #                                Acessa o menu de expedição e exibe estações                                        #

        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menu"]/a[1]'))).click()  # Menu
        navegador.find_element(By.XPATH, "//a[contains(text(), 'Expedição')]").click()  # Expedição
        navegador.find_element(By.XPATH, "//a[contains(text(), ' Expedição Simultânea')]").click()  # Expedição simultânea
        if navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes" and text() = "Ocultar estações"]'):
            navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]')
        else:
            navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()

        return navegador
    except Exception as e:
        print(f"Erro ao iniciar o navegador: {e}")
        return None


def abrir_rotulos_plano2(cod, nomepdf, nomeexcel, nomeplano, celula,page):
    # Selecionando o botão "Rótulos PDF" pelo texto
    navegador.find_element(By.XPATH, f'//*[@id="{cod}" and text()="Rótulos PDF"]').click()

    z = navegador.find_element(By.XPATH, '//*[@id="Z"]')
    navegador.execute_script("arguments[0].click()", z)
    # navegador.find_element(By.XPATH, '//*[@id="Z"]').click()  # Célula Z
    mudar_tipo_cdl_mala("MLA 04")
    # Clica no checkbox
    button0 = navegador.find_element(By.XPATH, '//*[@id="input-checkbox-termica"]')
    navegador.execute_script("arguments[0].click()", button0)

    # Confirma se está marcado o checkbox
    checkbox = navegador.find_element(By.XPATH, '//*[@id="input-checkbox-termica"]')
    if not checkbox.is_selected():
        navegador.find_element(By.XPATH, '//*[@id="input-checkbox-termica"]').click()

    # Clica no botão para download de caixetas e malas
    button = navegador.find_element(By.XPATH, '//*[@id="button-gerar-todos-rotulos-caixetas"]')
    navegador.execute_script("arguments[0].click();", button)
    navegador.back()

    # Caminho do arquivo PDF que esperamos baixar
    pdf_nome = nomepdf
    pdf_caminho = os.path.join(diretorio_downloads,
                               pdf_nome)  # Certifique-se de que `diretorio_downloads` está definido

    if esperar_download_arquivo(pdf_caminho):
        # Clica na célula
        navegador.find_element(By.XPATH, f'//*[@id="{cod}"]').click()
        navegador.find_element(By.XPATH, celula).click()
        # Converte PDF para Excel
        pdf_to_excelucub(pdf_nome, nomeexcel)
        complemento = usuario + complementar
        navegador.find_element(By.XPATH, '//*[@id="matricula"]').send_keys(
            complemento + Keys.ENTER)
        registrar_rotulo(nomeexcel)
        deletar_arquivo(nomeexcel)

    else:
        print("Erro: Arquivo PDF não foi baixado a tempo.")

    # Navegação e limpeza de arquivos
    navegador.get('https://sroweb.correios.com.br/app/expedicao/expedicaosimultanea/index.php')
    if navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes" and text() = "Ocultar estações"]'):
        navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]')
    else:
        navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()
    # navegador.find_element(By.XPATH,                #                       '//*[@id="link-exibir-ocultar-estacoes"]').click()  # Exibir estações
    deletar_arquivo(pdf_nome)
    deletar_arquivo(nomeexcel)
    page.add(mensagem_sucesso(f"Rótulos {nomeplano} aberto com sucesso!"))
    page.update()


def mudar_tipo_cdl_mala(nome):
    # Espera até que os elementos estejam carregados (substitua 'modellist-rotulos' com o identificador correto)
    wait = WebDriverWait(navegador, 10)
    elementos = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "modellist-rotulos")))

    # Itera por todos os elementos encontrados e altera o texto quando necessário
    for elemento in elementos:
        if elemento.get_attribute("value") == "CDL G":
            elemento.clear()  # Limpa o campo antes de inserir o novo valor
            elemento.send_keys(nome)  # Insere o novo valor


def registrar_rotulo(nome):
    tabela = pd.read_excel(nome)

    for lacres in tabela['Lacre']:

        print(lacres)
        while (navegador.find_element(By.XPATH, '//*[@id="sub-mensagem_princial"]').text == 'Aguarde...'):
            time.sleep(1)
        navegador.find_element(By.XPATH, '//*[@id="matricula"]').send_keys(lacres)
        navegador.find_element(By.XPATH, '//*[@id="matricula"]').send_keys(Keys.ENTER)

        while (navegador.find_element(By.XPATH,
                                      '//*[@id="sub-mensagem_princial"]').text == 'Aguarde...'):
            time.sleep(1)
        navegador.find_element(By.XPATH, '//*[@id="matricula"]').send_keys(lacres)
        navegador.find_element(By.XPATH, '//*[@id="matricula"]').send_keys(Keys.ENTER)


def resto_rotulos(nome_arquivo):
    # Encontra todos os elementos <div> com a classe 'texto-esquerda' que contém o texto 'Lacre: '
    lacre_elements = navegador.find_elements(By.XPATH,
                                             "//div[@class='texto-esquerda' and contains(text(), 'Lacre: ')]")

    # Extrai os textos dos códigos de lacre encontrados, filtrando apenas os que começam com 'UC' ou 'UB'
    lacres = []
    for element in lacre_elements:
        texto = element.text
        # Verifica se o texto contém 'Lacre: ' e aplica o regex para encontrar lacres que começam com UC ou UB
        match = re.search(r'\b(U[CB]\d{9})\b', texto)  # Regex para 'UC' ou 'UB' seguidos de 9 dígitos
        if match:
            lacres.append(match.group(1))  # Adiciona o código encontrado na lista

    # Verifica se lacres foram encontrados
    if lacres:
        # Salva os lacres em um DataFrame do pandas
        df = pd.DataFrame(lacres, columns=["Lacre"])

        # Exporta o DataFrame para um arquivo Excel usando openpyxl como motor
        df.to_excel(nome_arquivo, index=False, engine='openpyxl')

        print("Códigos de lacre salvos no arquivo CODIGOS.xlsx")
    else:
        print("Nenhum código de lacre encontrado.")


def abrir_rotulos_plano(cod, nomepdf, nomeexcel, nomeplano,page):
    # Selecionando o botão "Rótulos PDF" pelo texto
    navegador.find_element(By.XPATH, f'//*[@id="{cod}" and text()="Rótulos PDF"]').click()
    z = navegador.find_element(By.XPATH, '//*[@id="Z"]')
    navegador.execute_script("arguments[0].click()", z)
    # navegador.find_element(By.XPATH, '//*[@id="Z"]').click()  # Célula Z
    mudar_tipo_cdl_mala("MLA 04")
    # Clica no checkbox
    button0 = navegador.find_element(By.XPATH, '//*[@id="input-checkbox-termica"]')
    navegador.execute_script("arguments[0].click()", button0)

    # Confirma se está marcado o checkbox
    checkbox = navegador.find_element(By.XPATH, '//*[@id="input-checkbox-termica"]')
    if not checkbox.is_selected():
        navegador.find_element(By.XPATH, '//*[@id="input-checkbox-termica"]').click()

    # Clica no botão para download de caixetas e malas
    button = navegador.find_element(By.XPATH, '//*[@id="button-gerar-todos-rotulos-caixetas"]')
    navegador.execute_script("arguments[0].click();", button)
    navegador.back()

    # Caminho do arquivo PDF que esperamos baixar
    pdf_nome = nomepdf
    pdf_caminho = os.path.join(diretorio_downloads,
                               pdf_nome)  # Certifique-se de que `diretorio_downloads` está definido

    if esperar_download_arquivo(pdf_caminho):

        navegador.find_element(By.XPATH, f'//*[@id="{cod}"]').click()
        navegador.find_element(By.XPATH, '//*[@id="cards-estacao-celulas"]/div[8]/a/h3').click()
        pdf_to_excelucub(pdf_nome, nomeexcel)
        complemento = usuario + complementar
        navegador.find_element(By.XPATH, '//*[@id="matricula"]').send_keys(complemento + Keys.ENTER)
        registrar_rotulo(nomeexcel)
        deletar_arquivo(nomeexcel)

    else:
        print("Erro: Arquivo PDF não foi baixado a tempo.")

    # Navegação e limpeza de arquivos
    navegador.get('https://sroweb.correios.com.br/app/expedicao/expedicaosimultanea/index.php')

    if navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes" and text() = "Ocultar estações"]'):
        navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]')
    else:
        navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()
    deletar_arquivo(pdf_nome)
    deletar_arquivo(nomeexcel)
    page.add(mensagem_sucesso(f"Rótulos {nomeplano} aberto com sucesso!"))
    page.update()

    # navegador.find_element(By.XPATH,'//*[@id="link-exibir-ocultar-estacoes"]').click()   # Exibir estações
    # deletar_arquivo(pdf_nome)
    # deletar_arquivo(nomeexcel)
    # page.add(mensagem_sucesso(f"Rótulos {nomeplano} aberto com sucesso!"))
    # page.update()


def fechar_rotulos_plano2(cod, nomeplano, nomeexcel, nomeexcel2, celula,page):
    navegador.find_element(By.XPATH, f'//*[@id="{cod}"]').click()
    navegador.find_element(By.XPATH, celula).click()
    time.sleep(1)
    navegador.find_element(By.XPATH, '//*[@id="btn-objetos"]').click()
    alert_elements = navegador.find_elements(By.XPATH, '//*[@id="modal-alertas-corpo"]/p')
    if alert_elements:
        page.add(mensagem_sucesso(f"Rótulos {nomeplano} já estavam fechados!"))
        navegador.get(
            'https://sroweb.correios.com.br/app/expedicao/expedicaosimultanea/index.php')
        if navegador.find_element(By.XPATH,
                                  '//*[@id="link-exibir-ocultar-estacoes" and text() = "Ocultar estações"]'):
            navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]')
        else:
            navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()

        page.update()

    else:
        resto_rotulos(nomeexcel)
        navegador.find_element(By.XPATH, '//*[@id="modal-objetos"]/section/header/a').click()
        time.sleep(1)

        complemento = usuario + complementar
        navegador.find_element(By.XPATH, '//*[@id="matricula"]').send_keys(
            complemento + Keys.ENTER)
        registrar_rotulo(nomeexcel)
        navegador.find_element(By.XPATH, '//*[@id="btn-objetos"]').click()

        # Re-verifica o alerta para garantir que o lacre foi fechado com sucesso
        alert_elements = navegador.find_elements(By.XPATH, '//*[@id="modal-alertas-corpo"]/p')
        if alert_elements:
            page.add(mensagem_sucesso(f"Rótulos {nomeplano} fechado com sucesso!"))
            navegador.get(
                'https://sroweb.correios.com.br/app/expedicao/expedicaosimultanea/index.php')
            if navegador.find_element(By.XPATH,
                                      '//*[@id="link-exibir-ocultar-estacoes" and text() = "Ocultar estações"]'):
                navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]')
            else:
                navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()

            page.update()
        else:
            # Se o fechamento ainda não foi concluído, tenta novamente com novo arquivo
            resto_rotulos(nomeexcel2)
            navegador.get(
                'https://sroweb.correios.com.br/app/expedicao/expedicaosimultanea/index.php')
            if navegador.find_element(By.XPATH,
                                      '//*[@id="link-exibir-ocultar-estacoes" and text() = "Ocultar estações"]'):
                navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]')
            else:
                navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()
            navegador.find_element(By.XPATH, f'//*[@id="{cod}"]').click()
            navegador.find_element(By.XPATH,
                                   celula).click()
            complemento = usuario + complementar
            navegador.find_element(By.XPATH, '//*[@id="matricula"]').send_keys(
                complemento + Keys.ENTER)
            registrar_rotulo(nomeexcel2)
            deletar_arquivo(nomeexcel2)

            # Finaliza removendo arquivos e adicionando mensagem final
            navegador.get(
                'https://sroweb.correios.com.br/app/expedicao/expedicaosimultanea/index.php')
            if navegador.find_element(By.XPATH,
                                      '//*[@id="link-exibir-ocultar-estacoes" and text() = "Ocultar estações"]'):
                navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]')
            else:
                navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()
            deletar_arquivo(nomeexcel)
            page.add(mensagem_sucesso(f"Rótulos {nomeplano} fechado com sucesso!"))
            page.update()

def turno1():
    # Clica no botão
    navegador.find_element(By.XPATH, '//*[@id="tabs2"]/div/div/button[1]').click()
    # Verifica se o link está exibido com o texto correto
    try:
        if navegador.find_element(By.XPATH,
                                  '//*[@id="link-exibir-ocultar-estacoes" and text()="Ocultar estações"]'):
            print("Link já exibido")
    except NoSuchElementException:
        # Se o elemento não for encontrado, clica no link
        navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()
def turno2():
    # Clica no botão
    navegador.find_element(By.XPATH, '//*[@id="tabs2"]/div/div/button[2]').click()
    # Verifica se o link está exibido com o texto correto
    try:
        if navegador.find_element(By.XPATH,
                                  '//*[@id="link-exibir-ocultar-estacoes" and text()="Ocultar estações"]'):
            print("Link já exibido")
    except NoSuchElementException:
        # Se o elemento não for encontrado, clica no link
        navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()
def turno3():
    # Clica no botão
    navegador.find_element(By.XPATH, '//*[@id="tabs2"]/div/div/button[3]').click()
    # Verifica se o link está exibido com o texto correto
    try:
        if navegador.find_element(By.XPATH,
                                  '//*[@id="link-exibir-ocultar-estacoes" and text()="Ocultar estações"]'):
            print("Link já exibido")
    except NoSuchElementException:
        # Se o elemento não for encontrado, clica no link
        navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()


def fechar_rotulos_plano(cod, nomeplano, nomeexcel, nomeexcel2,page):
    navegador.find_element(By.XPATH, f'//*[@id="{cod}"]').click()
    navegador.find_element(By.XPATH, '//*[@id="cards-estacao-celulas"]/div[8]/a/h3').click()
    time.sleep(1)
    navegador.find_element(By.XPATH, '//*[@id="btn-objetos"]').click()
    alert_elements = navegador.find_elements(By.XPATH, '//*[@id="modal-alertas-corpo"]/p')
    if alert_elements:
        page.add(mensagem_sucesso(f"Rótulos {nomeplano} já estavam fechados!"))
        navegador.get(
            'https://sroweb.correios.com.br/app/expedicao/expedicaosimultanea/index.php')
        navegador.find_element(By.XPATH,
                               '//*[@id="link-exibir-ocultar-estacoes"]').click()

        page.update()

    else:
        resto_rotulos(nomeexcel)
        navegador.find_element(By.XPATH, '//*[@id="modal-objetos"]/section/header/a').click()
        time.sleep(1)

        complemento = usuario + complementar
        navegador.find_element(By.XPATH, '//*[@id="matricula"]').send_keys(
            complemento + Keys.ENTER)
        registrar_rotulo(nomeexcel)
        navegador.find_element(By.XPATH, '//*[@id="btn-objetos"]').click()

        # Re-verifica o alerta para garantir que o lacre foi fechado com sucesso
        alert_elements = navegador.find_elements(By.XPATH, '//*[@id="modal-alertas-corpo"]/p')
        if alert_elements:
            page.add(mensagem_sucesso(f"Rótulos {nomeplano} fechado com sucesso!"))
            navegador.get(
                'https://sroweb.correios.com.br/app/expedicao/expedicaosimultanea/index.php')
            if navegador.find_element(By.XPATH,
                                      '//*[@id="link-exibir-ocultar-estacoes" and text() = "Ocultar estações"]'):
                navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]')
            else:
                navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()

            page.update()
        else:
            # Se o fechamento ainda não foi concluído, tenta novamente com novo arquivo
            resto_rotulos(nomeexcel2)
            navegador.get(
                'https://sroweb.correios.com.br/app/expedicao/expedicaosimultanea/index.php')
            navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()
            navegador.find_element(By.XPATH, f'//*[@id="{cod}"]').click()
            navegador.find_element(By.XPATH,
                                   '//*[@id="cards-estacao-celulas"]/div[8]/a/h3').click()
            complemento = usuario + complementar
            navegador.find_element(By.XPATH, '//*[@id="matricula"]').send_keys(
                complemento + Keys.ENTER)
            registrar_rotulo(nomeexcel2)
            deletar_arquivo(nomeexcel2)

            # Finaliza removendo arquivos e adicionando mensagem final
            navegador.get(
                'https://sroweb.correios.com.br/app/expedicao/expedicaosimultanea/index.php')
            if navegador.find_element(By.XPATH,
                                      '//*[@id="link-exibir-ocultar-estacoes" and text() = "Ocultar estações"]'):
                navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]')
            else:
                navegador.find_element(By.XPATH, '//*[@id="link-exibir-ocultar-estacoes"]').click()

            deletar_arquivo(nomeexcel)
            page.add(mensagem_sucesso(f"Rótulos {nomeplano} fechado com sucesso!"))
            page.update()
