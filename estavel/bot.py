# Funcao main
from botcity.core import DesktopBot
from login_lge import login
import constants
from config import webbot_config
from app import common
import os
from botcity.web import WebBot
from botcity.maestro import *

from modules.gerp import gerp_web
from modules.gerp_production_desktop import dowload_files_from_gerp
from utils import helpers

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    # Runner passes the server URL, the ID of the task being executed,
    # the access token, and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    desktop_bot = DesktopBot()
    webbot = webbot_config.webbot_config(WebBot())

    username = constants.USERNAME
    password = constants.PASSWORD
    employee_id = constants.EMPLOYEE_ID
    personal_id = constants.PERSONAL_ID

    # Realizar login no sistema
    login.login(webbot, username, password, employee_id, personal_id)
    webbot.wait(5000)

    orgs = ['SP BOMM General User(NW8)']
    app_name = 'Issue History Inquiry'
    
    attachments = []
    
    for i, org in enumerate(orgs):
        latest_file = None

        # Iniciar processo no GERP Web
        try:
            gerp_web(webbot, org, app_name)
            helpers.keep_download_gerp_executable(webbot, i + 1)
            executable_gerp = webbot.get_last_created_file()

            # Verificar se o arquivo gerado é o esperado
            if executable_gerp[-15:] != "frmservlet.jnlp":
                print(f"Arquivo inesperado: {executable_gerp}, encerrando.")
                return
            
            desktop_bot.execute(executable_gerp)
            helpers.pass_warning_open_java_app(desktop_bot)
            helpers.activate_gerp_window(desktop_bot)
            
            # Baixar os arquivos e verificar se houve sucesso
            att = dowload_files_from_gerp(desktop_bot, org, latest_file)
            if att:
                print(f"Arquivo baixado para {org}: {att}")
                attachments.append(att)
            else:
                print(f"Nenhum arquivo baixado para {org}")
            
        except Exception as e:
            print(f"Erro ao processar {org}: {e}")
        
        # Esperar e fechar a janela do GERP
        desktop_bot.wait(4000)
        os.system("taskkill /f /im jp2launcher.exe")

    # Enviar email com anexos
    print(f"Anexos preparados: {attachments}")
    
    if attachments:
        is_test = '1'  # ou use a lógica para obter 'is_test' dinamicamente
        common.send_email_is_test(is_test, attachments)
        print("Email enviado com sucesso.")
    else:
        print("Nenhum anexo encontrado. Email não enviado.")
    
    # Finalizar a tarefa
    # common.send_log_smart_office(61, is_test)

if __name__ == '__main__':
    try:
        main()
    except Exception as ex:
        print(f"Erro na execução do bot: {ex}")
    finally:
        os.system("taskkill /f /im jp2launcher.exe")
        os.system("taskkill /f /im msedgedriver.exe")
        os.system("taskkill /f /im msedge.exe")
