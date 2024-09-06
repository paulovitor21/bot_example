####### funcao gerenciar arquivos ################################################
import shutil
import os
from zipfile import ZipFile
from pathlib import Path

def prepare_export_folder(folder_name="export_data"):
    """Verifica se a pasta export_data existe, caso contrário, cria.
       Exclui todos os arquivos da pasta antes de cada execução.
    """
    folder_path = Path(folder_name)
    
    # Verifica se a pasta existe, se não existir, cria
    if not folder_path.exists():
        folder_path.mkdir(parents=True)
    
    # Remove todos os arquivos na pasta
    for file in folder_path.glob("*"):
        try:
            if file.is_file():
                file.unlink()
        except Exception as e:
            print(f"Erro ao tentar remover {file}: {e}")

def move_file_to_export(file_path, folder_name="export_data"):
    """Move o arquivo baixado para a pasta export_data."""
    try:
        destination = Path(folder_name) / Path(file_path).name
        shutil.move(file_path, destination)
        return destination
    except Exception as e:
        print(f"Erro ao mover o arquivo {file_path} para {folder_name}: {e}")
        return None

def zip_files(folder_name="export_data", zip_name="export_data.zip"):
    """Cria um arquivo zip contendo todos os arquivos da pasta export_data e remove os arquivos XLS depois."""
    try:
        zip_path = Path(folder_name) / zip_name
        with ZipFile(zip_path, 'w') as zipf:
            # Adiciona os arquivos ao ZIP
            for file in Path(folder_name).glob("*.xls"):
                if file.is_file():
                    zipf.write(file, file.name)

        # Remover os arquivos XLS após a criação do ZIP
        for file in Path(folder_name).glob("*.xls"):
            try:
                file.unlink()  # Remove o arquivo XLS
            except Exception as e:
                print(f"Erro ao remover o arquivo {file}: {e}")

        return zip_path
    except Exception as e:
        print(f"Erro ao zipar arquivos: {e}")
        return None

        os.system("taskkill /f /im msedgedriver.exe")
        os.system("taskkill /f /im msedge.exe")