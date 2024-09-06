from botcity.core import DesktopBot
from utils.ClassBot import ClassBot
from pathlib import Path
from utils import helpers
import shutil

def dowload_files_from_gerp(desktop_bot: DesktopBot, orgs, latest_file):
    bot = ClassBot()

    # Wait for GERP to open
    desktop_bot.wait(40000)
    
    try:
        # Check and click "OK" button if present
        bot.wait_for_element("button_ok", waiting_time=10000)
        bot.click_button("button_ok")
    except Exception:
        pass
        
    # Search and click the "Find" button
    bot.wait_for_element("button_find", waiting_time=10000)
    bot.click_button("button_find")
    bot.wait(6000)

    try:
        bot.wait_for_element("error_invalid_number", waiting_time=8000)
        return
    except Exception as e:
        print(e)

    bot.wait(4000)

    try:
        bot.wait_for_element("total_sales_amount", waiting_time=20000)

        bot.click_button("button_file")
        for _ in range(6):
            desktop_bot.type_down()
        desktop_bot.enter()
        
        bot.wait_for_element("excel_export_img", waiting_time=300000)
                    
        bot.wait(6000)
        bot.wait_for_element("salvar_como", waiting_time=5000)
        bot.enter() 
        desktop_bot.wait(5000)
        helpers.keep_download_gerp_results_chrome(desktop_bot)
            
        downloads_path = str(Path.home() / "Downloads")
        latest_file = bot.get_last_file(downloads_path)
        print("SHEET => ", latest_file)

        # Renomear o arquivo baixado conforme a organização
        renamed_file = rename_file_by_org(latest_file, orgs, "export_data")
        return renamed_file
    except Exception as e:
        print(e)


def rename_file_by_org(file_path, org_name, folder_name="export_data"):
    """Renomeia o arquivo baixado conforme a organização e move para a pasta export_data."""
    try:
        # Extrair a extensão do arquivo
        file_extension = Path(file_path).suffix
        # Criar um novo nome baseado na organização
        new_file_name = f"{org_name}{file_extension}"
        # Criar o novo caminho para o arquivo
        destination = Path(folder_name) / new_file_name
        
        # Mover e renomear o arquivo
        shutil.move(file_path, destination)
        return destination
    except Exception as e:
        print(f"Erro ao renomear o arquivo {file_path} para {org_name}: {e}")
        return None

