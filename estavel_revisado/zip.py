from zipfile import ZipFile
from pathlib import Path

def zip_files(folder_name="export_data", zip_name="export_data.zip", file_path=None):
    """
    Cria um arquivo zip contendo todos os arquivos da pasta export_data ou um único arquivo.
    Remove os arquivos XLS depois se forem de uma pasta.
    
    :param folder_name: Nome da pasta para compactar (não usado se `file_path` for especificado).
    :param zip_name: Nome do arquivo ZIP a ser criado.
    :param file_path: Caminho de um arquivo específico a ser compactado (opcional).
    :return: Caminho do arquivo ZIP criado ou None em caso de erro.
    """
    try:
        # Define o caminho de saída para o arquivo ZIP
        zip_path = Path(folder_name) / zip_name if not file_path else Path(file_path).parent / zip_name

        with ZipFile(zip_path, 'w') as zipf:
            if file_path:  # Se um arquivo específico foi passado
                file = Path(file_path)
                if file.is_file():
                    # Adiciona o arquivo ao ZIP com um nome relativo (para evitar diretórios complexos)
                    zipf.write(file, file.name)
            else:  # Se nenhum arquivo foi especificado, zipar todos os arquivos da pasta
                for file in Path(folder_name).glob("*.xls"):
                    if file.is_file():
                        zipf.write(file, file.name)

        # Se for uma pasta, remover os arquivos XLS após a criação do ZIP
        if not file_path:
            for file in Path(folder_name).glob("*.xls"):
                try:
                    file.unlink()  # Remove o arquivo XLS
                except Exception as e:
                    print(f"Erro ao remover o arquivo {file}: {e}")

        return zip_path
    except Exception as e:
        print(f"Erro ao zipar arquivos: {e}")
        return None