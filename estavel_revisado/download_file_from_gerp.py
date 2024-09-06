import time
import shutil
from pathlib import Path
from botcity.core import DesktopBot
from utils import helpers
from utils.ClassBot import ClassBot

def check_download_complete(downloads_path, timeout=300):
    """Verifica se o download foi concluído verificando se o arquivo está presente na pasta de downloads e não está mais sendo escrito."""
    start_time = time.time()
    files_before = set(Path(downloads_path).glob("*"))
    
    while time.time() - start_time < timeout:
        files_after = set(Path(downloads_path).glob("*"))
        if files_after != files_before:
            # Arquivo novo ou alterado encontrado, verificar se está completamente baixado
            for file in files_after - files_before:
                if file.stat().st_size > 0:
                    # Verificar se o arquivo ainda está em processo de download
                    try:
                        with open(file, 'rb') as f:
                            if len(f.read()) == file.stat().st_size:
                                return file
                    except IOError:
                        continue
        time.sleep(5)
    return None

def download_files_from_gerp(desktop_bot: DesktopBot, orgs):
    bot = ClassBot()
    
    # Aguarda o GERP abrir
    desktop_bot.wait(40000)
    
    try:
        # Clica no botão OK, se presente
        bot.wait_for_element("button_ok", waiting_time=10000)
        bot.click_button("button_ok")
    except Exception as e:
        print(f"Botão OK não encontrado ou erro: {e}")
    
    bot.wait_for_element("button_find", waiting_time=10000)
    bot.click_button("button_find")
    bot.wait(6000)

    try:
        bot.wait_for_element("error_invalid_number", waiting_time=8000)
        print("Número inválido detectado.")
        return
    except Exception:
        pass  # Se não encontrar o erro, continua

    bot.wait(4000)

    try:
        bot.wait_for_element("total_sales_amount", waiting_time=20000)

        # Exportação para Excel
        bot.click_button("button_file")
        for _ in range(6):
            desktop_bot.type_down()
        desktop_bot.enter()

        bot.wait_for_element("excel_export_img", waiting_time=300000)

        bot.wait(6000)
        bot.wait_for_element("salvar_como", waiting_time=5000)
        bot.enter()

        # Verificar se o download foi concluído
        downloads_path = str(Path.home() / "Downloads")
        downloaded_file = check_download_complete(downloads_path)

        if downloaded_file:
            helpers.keep_download_gerp_results_chrome(desktop_bot)
            print(f"Arquivo baixado: {downloaded_file}")
            
            # Renomear e mover o arquivo baixado conforme a organização
            renamed_file = rename_file_by_org(downloaded_file, orgs, "export_data")
            return renamed_file
        else:
            print("Download não concluído no tempo especificado.")
            return None
    except Exception as e:
        print(f"Erro durante o download ou processamento: {e}")
        return None

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
