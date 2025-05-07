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
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from utils_pdf import diretorio_downloads, pasta_raiz , renomear_pdf , pdf_to_excelucub
from utils import deletar_arquivo , esperar_download , mensagem_sucesso , esperar_download_arquivo
navegador = None
usuario = None
user_mat = None

from utils import iniciar_navegador, abrir_rotulos_plano2, turno1, turno2, turno3, fechar_rotulos_plano2

# Função principal do aplicativo
def main(page: ft.Page):
    global navegador  # Declare como global
    global usuario  # Declare como global
    global user_mat
    page.title = "Gestão de Fechamento e Abertura de Rótulos"
    page.scroll = ft.ScrollMode.AUTO
    #page.window.height = 900
    #page.window.width = 550
    # Função de login

    """
    def realizar_login(e):
        global usuario
        usuario = usuario_input.value
        senha = senha_input.value

        if usuario and senha:
            page.clean()  # Limpa a tela de login
            navegador = iniciar_navegador(usuario, senha)  # Inicia o navegador com Selenium
            if navegador:
                carregar_painel_informacoes(navegador)
                user_mat = usuario
                print(f"O usuario é {user_mat}")
            else:
                login_feedback.value = "Erro de login: usuário ou senha incorretos  Tente novamente."
                page.add(login_feedback)  # Adiciona a mensagem de feedback à página
                page.update()
                time.sleep(3)
                page.clean()
                main(page)
        else:
            login_feedback.value = "Por favor, insira as credenciais!"
            page.add(login_feedback)
            page.update()
    """

    def realizar_login(e):
        global usuario
        usuario = campo_usuario.value
        senha = campo_senha.value
        mensagem.value = "Realizando login..."
        mensagem.color = "yellow"
        page.update()

        # Iniciar Selenium
        service = Service(ChromeDriverManager().install())
        options = webdriver.ChromeOptions()
        navegador = webdriver.Chrome(service=service, options=options)
        navegador.get("https://sroweb.correios.com.br")

        try:
            navegador.find_element(By.XPATH, '//*[@id="username"]').send_keys(usuario)
            navegador.find_element(By.XPATH, '//*[@id="password"]').send_keys(senha)
            navegador.find_element(By.XPATH,
                                   '/html/body/main/div[1]/div/section/section/div/div/form/div/div[2]/button').click()

            # Esperar a página carregar e verificar login
            WebDriverWait(navegador, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="menu"]/a[1]')))
            mensagem.value = "✅ Login realizado com sucesso!"
            mensagem.color = "green"
            page.update()
            time.sleep(2)
            navegador.quit()
        except:
            mensagem.value = "❌ Erro no login! Verifique suas credenciais."
            mensagem.color = "red"
            page.update()
            navegador.quit()

    robo_imagem = ft.Image(
        src="https://media.istockphoto.com/id/1958063552/pt/foto/digital-ai-newsletter-concept-ai-powered-email-marketing-automation-marketing-automation-email.jpg?s=1024x1024&w=is&k=20&c=WAhfwZu2XNWU8zDfyOm2RPeAkQMxCoCptKpY2ASbIP4=",
        width=150, height=150)
    titulo = ft.Text("Acesso Robótico", size=24, color="#00ffcc", weight=ft.FontWeight.BOLD)
    campo_usuario = ft.TextField(label="Usuário", width=300, bgcolor="#1a1a1a", color="white")
    campo_senha = ft.TextField(label="Senha", password=True, width=300, bgcolor="#1a1a1a", color="white")
    botao_login = ft.ElevatedButton("Entrar", on_click=realizar_login, bgcolor="#00ffcc", color="#0a0a0a")
    mensagem = ft.Text(value="", size=16, weight=ft.FontWeight.BOLD)

    # Função para carregar o painel de informações após o login
    def carregar_painel_informacoes(navegador):
        blocos = [
            {"titulo": "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_1_P1_A_SQ", "celula_task": "task1", "rotulos_task": "task2"},
            {"titulo": "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_2_P1_B_SQ", "celula_task": "task3", "rotulos_task": "task4"},
            {"titulo": "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_3_P1_A_TQSS", "celula_task": "task5", "rotulos_task": "task6"},
            {"titulo": "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_4_P1_B_TQSS", "celula_task": "task7", "rotulos_task": "task8"},
            {"titulo": "E1_CTCE_BHE_1_IMP_SAP_PCT_SDX_4_P1_A_SQ", "celula_task": "task9", "rotulos_task": "task10"},
            {"titulo": "E1_CTCE_BHE_1_IMP_SAP_PCT_SDX_5_P1_B_SQ", "celula_task": "task11", "rotulos_task": "task12"},
            {"titulo": "E1_CTCE_BHE_2_IMP_SAP_PCT_SDX_3_P1_A_TQSS", "celula_task": "task13", "rotulos_task": "task14"},
            {"titulo": "E1_CTCE_BHE_2_IMP_SAP_PCT_SDX_4_P1_B_TQSS", "celula_task": "task15", "rotulos_task": "task16"},
            {"titulo": "E1_CTCE_BHE_1_IMP_SAP_PCT_SDX_6_P2_TQSS", "celula_task": "task17", "rotulos_task": "task18"},
            {"titulo": "E1_CTCE_BHE_3_IMP_SAP_PCT_SDX_2_P2_SQ", "celula_task": "task19", "rotulos_task": "task20"},
        ]

        def criar_bloco(bloco):
            feedback_text = ft.Text("", size=16, color=ft.colors.RED, weight=ft.FontWeight.BOLD, )
            def Abrir(e):
                lottie_animacao = None
                try:
                    lottie_animacao = animacao(page)
                    feedback_text.value = f"Abrindo: {bloco['titulo']}..."
                    page.add( feedback_text , lottie_animacao)
                    page.update()

                    match bloco["titulo"]:

                        case "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_1_P1_A_SQ":
                            turno1()
                            abrir_rotulos_plano2(13573, "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_1_P1_A_SQ_Z_TODOS_ROTULOS.pdf","pacp1asq.xlsx", "PAC P1 A SQ","//h3[contains(text(), 'Abertura 1 - Z')]",page)

                            #abrir_rotulos_plano(13573,"S1_CTCE_BHE_1_IMP_SAP_GO_PAC_1_P1_A_SQ_Z_TODOS_ROTULOS.pdf","pacp1asq.xlsx","PAC P1 A SQ")
                        case "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_2_P1_B_SQ":
                            turno1()
                            abrir_rotulos_plano2(13575,"S1_CTCE_BHE_1_IMP_SAP_GO_PAC_2_P1_B_SQ_Z_TODOS_ROTULOS.pdf","pacp1bsq.xlsx","PAC P1 B SQ","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_3_P1_A_TQSS":
                            turno1()
                            abrir_rotulos_plano2(13578,"S1_CTCE_BHE_1_IMP_SAP_GO_PAC_3_P1_A_TQSS_Z_TODOS_ROTULOS.pdf","pacp1ATQSS.xlsx","PAC P1 A TQSS","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_4_P1_B_TQSS":
                            turno1()
                            abrir_rotulos_plano2(13579,"S1_CTCE_BHE_1_IMP_SAP_GO_PAC_4_P1_B_TQSS_Z_TODOS_ROTULOS.pdf","pacp1BTQSS.xlsx","PAC P1 B TQSS","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "E1_CTCE_BHE_1_IMP_SAP_PCT_SDX_4_P1_A_SQ":
                            turno1()
                            abrir_rotulos_plano2(13588,"E1_CTCE_BHE_1_IMP_SAP_PCT_SDX_4_P1_A_SQ_Z_TODOS_ROTULOS.pdf","SEDEXP1ASQ.xlsx","SEDEX P1 A SQ","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "E1_CTCE_BHE_1_IMP_SAP_PCT_SDX_5_P1_B_SQ":
                            turno1()
                            abrir_rotulos_plano2(13591,"E1_CTCE_BHE_1_IMP_SAP_PCT_SDX_5_P1_B_SQ_Z_TODOS_ROTULOS.pdf","SEDEXP1BSQ.xlsx","SEDEX P1 B SQ","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "E1_CTCE_BHE_2_IMP_SAP_PCT_SDX_3_P1_A_TQSS":
                            turno2()
                            abrir_rotulos_plano2(13593,"E1_CTCE_BHE_2_IMP_SAP_PCT_SDX_3_P1_A_TQSS_Z_TODOS_ROTULOS.pdf","SEDEXP1ATQSS.xlsx","SEDEX P1 A TQSS","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "E1_CTCE_BHE_2_IMP_SAP_PCT_SDX_4_P1_B_TQSS":
                            turno2()
                            abrir_rotulos_plano2(13594,"E1_CTCE_BHE_2_IMP_SAP_PCT_SDX_4_P1_B_TQSS_Z_TODOS_ROTULOS.pdf","SEDEXP1BTQSS.xlsx","SEDEX P1 B TQSS","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "E1_CTCE_BHE_1_IMP_SAP_PCT_SDX_6_P2_TQSS":
                            turno1()
                            abrir_rotulos_plano2(13599,"E1_CTCE_BHE_1_IMP_SAP_PCT_SDX_6_P2_TQSS_Z_TODOS_ROTULOS.pdf","SEDEXP2TQSS.xlsx","SEDEX P2  TQSS","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "E1_CTCE_BHE_3_IMP_SAP_PCT_SDX_2_P2_SQ":
                            turno3()
                            abrir_rotulos_plano2(13601,"E1_CTCE_BHE_3_IMP_SAP_PCT_SDX_2_P2_SQ_Z_TODOS_ROTULOS.pdf","SEDEXP2SQ.xlsx","SEDEX P2 SQ","//h3[contains(text(), 'Abertura 1 - Z')]",page)



                except Exception as error:
                    print(f"Erro ao gerar rótulo PDF: {error}")
                finally:
                    # Removendo a animação caso ela tenha sido adicionada com sucesso
                    if lottie_animacao and lottie_animacao in page.controls:
                        page.controls.remove(lottie_animacao)
                    if feedback_text and feedback_text in page.controls:
                        page.controls.remove(feedback_text)
                    # feedback_text.value = f"Concluído: {bloco['titulo']}."
                    page.update()

            def clicar_elemento(xpath):
                try:
                    elemento = WebDriverWait(navegador, 10).until(
                        EC.element_to_be_clickable((By.XPATH, xpath))
                    )
                    elemento.click()
                except (NoSuchElementException, TimeoutException):
                    print(f"Elemento não encontrado: {xpath}")


            def animacao(page):
                # Carrega e codifica o conteúdo do arquivo JSON em Base64
                with open("../16-04-2025/esperando.json", "r") as lottie_file:
                    animation_data = lottie_file.read()
                    animation_base64 = base64.b64encode(animation_data.encode()).decode()

                # Retorna a animação Lottie para ser usada no page.add
                return ft.Lottie(
                    src_base64=animation_base64,  # Use `src_base64` para animação local
                    reverse=False,
                    animate=True,
                    width=200,
                    height=200
                )

            def fechar_rotulo_pdf(e):
                lottie_animacao = None
                try:
                    lottie_animacao = animacao(page)
                    feedback_text.value = f"Fechando: {bloco['titulo']}..."
                    page.add( feedback_text,lottie_animacao)
                    page.update()

                    match bloco["titulo"]:

                        case "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_1_P1_A_SQ":
                            turno1()
                            fechar_rotulos_plano2(13573,"PAC PLANO 1 A SQ",'plano1sq.xlsx',"plano1sq_.xlsx","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_2_P1_B_SQ":
                            turno1()
                            fechar_rotulos_plano2(13575,"PAC PLANO 1 B SQ","planobsq.xlsx","planobsq_.xlsx","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_3_P1_A_TQSS":
                            turno1()
                            fechar_rotulos_plano2(13578,"PAC PLANO 1 A TQSS",'plano1tqss.xlsx',"plano1tqss_.xlsx","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "S1_CTCE_BHE_1_IMP_SAP_GO_PAC_4_P1_B_TQSS":
                            turno1()
                            fechar_rotulos_plano2(13579,"PAC PLANO 1 B TQSS",'plano1BTQSS.xlsx',"plano1BTQSS_.xlsx","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "E1_CTCE_BHE_1_IMP_SAP_PCT_SDX_4_P1_A_SQ":
                            turno1()
                            fechar_rotulos_plano2(13588,"SEDEX PLANO 1 A SQ",'plano1sqSEDEX.xlsx',"plano1sqSEDEX_.xlsx","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "E1_CTCE_BHE_1_IMP_SAP_PCT_SDX_5_P1_B_SQ":
                            turno1()
                            fechar_rotulos_plano2(13591,"SEDEX PLANO 1 B SQ",'plano1sqBSEDEX.xlsx',"plano1sqBSEDEX_.xlsx","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "E1_CTCE_BHE_2_IMP_SAP_PCT_SDX_3_P1_A_TQSS":
                            turno2()
                            fechar_rotulos_plano2(13593,"SEDEX PLANO 1 A TQSS",'plano1TQSSEDEX.xlsx',"plano1TQSSEDEX_.xlsx","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "E1_CTCE_BHE_2_IMP_SAP_PCT_SDX_4_P1_B_TQSS":
                            turno2()
                            fechar_rotulos_plano2(13594,"SEDEX PLANO 1 B TQSS",'plano1TQSSBSEDEX.xlsx',"plano1TQSSBSEDEX_.xlsx","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "E1_CTCE_BHE_1_IMP_SAP_PCT_SDX_6_P2_TQSS":
                            turno1()
                            fechar_rotulos_plano2(13599,"SEDEX PLANO 2 TQSS",'plano2TQSSSEDEX.xlsx',"plano2TQSSSEDEX_.xlsx","//h3[contains(text(), 'Abertura 1 - Z')]",page)
                        case "E1_CTCE_BHE_3_IMP_SAP_PCT_SDX_2_P2_SQ":
                            turno3()
                            fechar_rotulos_plano2(13601,"SEDEX PLANO 2 SQ",'plano2sqSEDEX.xlsx',"plano2sqSEDEX_.xlsx","//h3[contains(text(), 'Abertura 1 - Z')]",page)

                except Exception as error:
                    print(f"Erro ao gerar rótulo PDF: {error}")
                finally:
                    # Removendo a animação caso ela tenha sido adicionada com sucesso
                    if lottie_animacao and lottie_animacao in page.controls:
                        page.controls.remove(lottie_animacao)
                    if feedback_text and feedback_text in page.controls:
                        page.controls.remove(feedback_text)
                    page.update()
            # Função para criar cada bloco
            return ft.Container(
                #height=150,
                #width=300,
                content=ft.Column(
                    controls=[
                        ft.Text(bloco["titulo"], size=20, weight="bold"),
                        ft.Row(
                            controls=[
                                ft.ElevatedButton("Abrir Rótulos", on_click=Abrir),
                                ft.ElevatedButton("Fechar Rótulos", on_click=fechar_rotulo_pdf),
                            ],
                            alignment=ft.MainAxisAlignment.CENTER,
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                padding=10,
                bgcolor=ft.colors.LIGHT_BLUE_100,
                border_radius=10,
            )
        grindview = ft.GridView(
            expand=True,
            run_spacing=1.0,
            runs_count=3,
            spacing=1.0,

            )
            # Adiciona os blocos ao painel

        #for bloco in blocos:
         #   grindview.controls.append(criar_bloco(bloco))
          #  page.add(grindview)
        for bloco in blocos:
            page.add(criar_bloco(bloco))
            #grindview.controls.append(criar_bloco(bloco))
            #page.add(grindview)

    # Tela de login
    usuario_input = ft.TextField(label="Usuário", password=False)
    senha_input = ft.TextField(label="Senha", password=True, can_reveal_password=True)
    login_button = ft.ElevatedButton(text="Login", on_click=realizar_login)
    login_feedback = ft.Text(value="", color=ft.colors.RED)

    # Layout de login
    login_layout = ft.Column(
        controls=[usuario_input, senha_input, login_button, login_feedback],
        alignment=ft.MainAxisAlignment.CENTER,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER

    )
    page.add(login_layout,)
    page.update()

# Executa o aplicativo
ft.app(target=main)