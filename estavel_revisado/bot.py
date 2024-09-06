# Importações
from botcity.core import DesktopBot
from login_lge import login
import constants
from config import webbot_config
from app import common
import os
from botcity.web import WebBot
from botcity.maestro import *
from modules.gerp import gerp_web
from modules.gerp_production_desktop import download_files_from_gerp
from utils import helpers
from utils.file_manager import prepare_export_folder, move_file_to_export, zip_files

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    # Inicialização do Maestro e captura da execução
    maestro = BotMaestroSDK.from_sys_args()
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    # Inicializando o bot de desktop e web
    desktop_bot = DesktopBot()
    webbot = webbot_config.webbot_config(WebBot())

    # Constantes para login
    username = constants.USERNAME
    password = constants.PASSWORD
    employee_id = constants.EMPLOYEE_ID
    personal_id = constants.PERSONAL_ID

    # Login no sistema
    login.login(webbot, username, password, employee_id, personal_id)
    webbot.wait(5000)

    # Lista de organizações a serem processadas
    orgs = [
        'SP BOMM General User(NW8)',
    ]

    app_name = 'Issue History Inquiry'

    # Preparar a pasta de exportação
    prepare_export_folder()

    # Lista para armazenar os caminhos dos arquivos baixados
    attachments = []

    # Processamento de cada organização
    for i in range(len(orgs)):
        
        # Abrir o GERP e iniciar o processo de download
        gerp_web(webbot, orgs[i], app_name)
        helpers.keep_download_gerp_executable(webbot, i+1)
        executable_gerp = webbot.get_last_created_file()

        if executable_gerp[-15:] != "frmservlet.jnlp":
            return  # Caso o arquivo não seja o esperado, sai da função

        # Executa o arquivo do GERP
        desktop_bot.execute(executable_gerp)
        helpers.pass_warning_open_java_app(desktop_bot)
        helpers.activate_gerp_window(desktop_bot)

        # Download e renomear o arquivo
        downloaded_file = download_files_from_gerp(desktop_bot, orgs[i])
        
        if downloaded_file:
            # Mover o arquivo baixado para a pasta export_data
            moved_file = move_file_to_export(downloaded_file)
            if moved_file:
                attachments.append(moved_file)
        
        desktop_bot.wait(4000)
        os.system("taskkill /f /im jp2launcher.exe")

    # Zipa os arquivos baixados
    zip_file = zip_files()

    # Verifica se o arquivo zip foi criado e se há anexos
    if zip_file and zip_file.exists():
        is_test = '1'
        # Adiciona o arquivo zip à lista de anexos
        attachments.append(zip_file)
        
        # Envia o e-mail com os anexos
        common.send_email_is_test(is_test, attachments)
    
    # Fechando o Chrome no final
    os.system("taskkill /f /im msedgedriver.exe")
    os.system("taskkill /f /im msedge.exe")


if __name__ == '__main__':
    try:
        main()
    except Exception as ex:
        print(ex)
    finally:
        os.system("taskkill /f /im jp2launcher.exe")
        os.system("taskkill /f /im msedgedriver.exe")
        os.system("taskkill /f /im msedge.exe")
