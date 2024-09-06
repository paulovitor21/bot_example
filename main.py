############### main #################
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
from app.gerenciar_arquivos import prepare_export_folder, zip_files

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():

    # Preparar a pasta export_data
    prepare_export_folder("export_data")
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
    # orgs = [
    #     'SP BOMM General User(NW1)',
    #     'SP BOMM General User(NW4)', 
    #     'SP BOMM General User(NW6)', 
    #     'SP BOMM General User(NW7)', 
    #     'SP BOMM General User(NW8)', 
    #     'SP BOMM General User(NWK)'
    # ]
    app_name = 'Issue History Inquiry'
    
    attachments = []
    for i in range(len(orgs)):
        latest_file = None
        gerp_web(webbot, orgs[i], app_name)
        helpers.keep_download_gerp_executable(webbot, i+1)
        executable_gerp = webbot.get_last_created_file()
        
        if executable_gerp[-15:] != "frmservlet.jnlp": return
        desktop_bot.execute(executable_gerp)
        
        helpers.pass_warning_open_java_app(desktop_bot)
        helpers.activate_gerp_window(desktop_bot)
    
        att = dowload_files_from_gerp(desktop_bot, orgs[i], latest_file)
        print("ARQUIVOS => ", att)
        attachments.append(att)
        
        desktop_bot.wait(4000)
        os.system("taskkill /f /im jp2launcher.exe")
    
     # Cria o arquivo zip
    zip_path = zip_files()
    
    # Envia email se o arquivo zip foi criado
    if zip_path and os.path.exists(zip_path):
        is_test = '1'
        common.send_email_is_test(is_test, [zip_path])

    # Zipar os arquivos da pasta export_data
    # zip_file = zip_files("export_data")
    # attachments.append(zip_file)
    
    # is_test = '1'    
    # common.send_email_is_test(is_test, attachments)
    
    # Finalizar a tarefa
    # common.send_log_smart_office(61, is_test)

if __name__ == '__main__':
    try:
        main()
    except Exception as ex:
        print(f"Erro na execução do bot: {ex}")
    finally:
        os.system("taskkill /f /im jp2launcher.exe")

