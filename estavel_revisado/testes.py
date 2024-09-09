import shutil
from pathlib import Path
from botcity.core import DesktopBot
from utils import helpers
from utils.ClassBot import ClassBot

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

        # Aqui a verificação de download foi removida
        helpers.keep_download_gerp_results_chrome(desktop_bot)
        print("Exportação concluída.")

        # Renomear e mover o arquivo baixado conforme a organização
        downloaded_file = Path.home() / "Downloads" / "nome_do_arquivo.xlsx"  # Exemplo de nome de arquivo
        renamed_file = rename_file_by_org(downloaded_file, orgs, "export_data")
        return renamed_file

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