"""
Funções para listar arquivos de uma pasta local e para identificar arquivos
enviados por upload (arrastar e soltar), usados tanto para processamento
quanto para o cache (saber se um arquivo já foi processado antes).
"""
import hashlib
import os

SUPPORTED_EXTENSIONS = (
    '.pdf', '.docx', '.doc', '.xlsx', '.xls',
    '.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.webp'
)


def list_local_files(folder_path: str):
    """Lista os arquivos suportados dentro de uma pasta local (sem subpastas)."""
    if not os.path.isdir(folder_path):
        raise NotADirectoryError(f"Pasta não encontrada: {folder_path}")

    files = []
    for name in sorted(os.listdir(folder_path)):
        full_path = os.path.join(folder_path, name)
        if os.path.isfile(full_path) and name.lower().endswith(SUPPORTED_EXTENSIONS):
            files.append({
                "name": name,
                "path": full_path,
                "modified_time": str(os.path.getmtime(full_path)),
            })
    return files


def hash_bytes(data: bytes) -> str:
    """Gera um hash do conteúdo do arquivo, usado como 'versão' para o cache
    de arquivos enviados por upload (se o conteúdo não mudar, o hash não muda)."""
    return hashlib.sha256(data).hexdigest()
