from botcity.core import DesktopBot
from utils.ClassBot import ClassBot
from pathlib import Path
from utils import helpers

def dowload_files_from_gerp(desktop_bot: DesktopBot, orgs, latest_file):
    """
    Downloads files from the GERP system, automating navigation and manipulation of interface elements using the bot configured for the graphical interface (DesktopBot).

    Parameters:
    - desktop_bot (DesktopBot): Instance of the bot for automating the graphical interface.
    - orgs (str): Name of the organization currently being processed.
    - latest_file (str): Name of the last file downloaded, used for renaming and verification.

    Returns:
    - str: The path to the downloaded and renamed file.
    - None: If an error occurs during the process or the organization is not found.
    """
    
    bot = ClassBot()

    # Wait for the GERP window to open
    desktop_bot.wait(40000)
    
    try:
        # Check and click "OK" button if present
        bot.wait_for_element("button_ok", waiting_time=10000)
        bot.click_button("button_ok")
    except Exception as e:
        print(f"Erro ao clicar no botão OK: {e}")
    
    try:
        # Search and click the "Find" button
        bot.wait_for_element("button_find", waiting_time=10000)
        bot.click_button("button_find")
        bot.wait(6000)

        # Check for an error message related to an invalid number
        try:
            bot.wait_for_element("error_invalid_number", waiting_time=8000)
            print("Número inválido encontrado, abortando.")
            return None
        except Exception:
            pass

        bot.wait(4000)

        # Wait for the "total sales amount" element to load
        bot.wait_for_element("total_sales_amount", waiting_time=20000)

        # Navigate the menu to export data as Excel
        bot.click_button("button_file")
        for _ in range(6):
            desktop_bot.type_down()
        desktop_bot.enter()

        # Wait for the export to Excel to complete
        bot.wait_for_element("excel_export_img", waiting_time=300000)

        # Wait for the "save as" window in Windows Explorer to proceed
        bot.wait(6000)
        bot.wait_for_element("salvar_como", waiting_time=5000)
        bot.enter() 
        desktop_bot.wait(5000)

        # Handle file download completion in Chrome
        helpers.keep_download_gerp_results_chrome(desktop_bot)
        
        # Find the most recent downloaded file
        downloads_path = str(Path.home() / "Downloads")
        latest_file = bot.get_last_file(downloads_path)
        
        if not latest_file:
            print("Nenhum arquivo encontrado na pasta de downloads.")
            return None
        
        print(f"Arquivo baixado: {latest_file}")

        # Rename the downloaded file according to the current organization
        renamed_file = bot.rename_file(latest_file, downloads_path, orgs)
        if renamed_file:
            print(f"Arquivo renomeado para: {renamed_file}")
            return renamed_file
        else:
            print(f"Falha ao renomear o arquivo {latest_file}")
            return None

    except Exception as e:
        print(f"Erro durante o download de arquivos do GERP: {e}")
        return None
