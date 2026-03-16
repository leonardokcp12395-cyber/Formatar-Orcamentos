import sys
import os
from pathlib import Path

def get_app_dir():
    """
    Retorna o diretório onde o 'programa' está rodando.
    - Se for .EXE: Retorna a pasta onde o .exe está (onde queremos salvar configs/DB).
    - Se for Script: Retorna a pasta raiz do projeto.
    """
    if getattr(sys, 'frozen', False):
        # Se for EXE, pega a pasta do executável
        return Path(os.path.dirname(sys.executable))
    else:
        # Se for script, pega a pasta raiz do projeto (subindo 2 níveis de core/)
        return Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def get_resource_path(relative_path):
    """
    Retorna onde os arquivos internos (imagens, icones) estão.
    - Se for .EXE: Pega da pasta temporária (_MEIPASS).
    - Se for Script: Pega da pasta raiz.
    """
    if getattr(sys, 'frozen', False):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    
    return base_path / relative_path