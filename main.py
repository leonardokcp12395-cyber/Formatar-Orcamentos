import os
import sys
import shutil
import threading
import webbrowser
from pathlib import Path
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import uvicorn
from contextlib import asynccontextmanager

# Configura estrutura de pastas automaticamente
def setup_environment():
    root = Path(os.path.dirname(os.path.abspath(__file__)))

    # 1. Pastas Obrigatórias
    folders = ['config', 'core', 'ui', 'utils', 'Output', 'static']
    for f in folders:
        (root / f).mkdir(exist_ok=True)

    # 2. Move arquivos soltos para config
    config_files = ['profiles.json', 'settings.json', 'autocomplete.json', 'last_session.json']
    for cf in config_files:
        src = root / cf
        dst = root / 'config' / cf
        if src.exists() and not dst.exists():
            shutil.move(str(src), str(dst))

    # Garante que o Python encontre tudo
    sys.path.append(str(root))

setup_environment()

def open_browser():
    webbrowser.open("http://localhost:8000")

@asynccontextmanager
async def lifespan(app: FastAPI):
    # Executa a abertura do browser após um pequeno delay para garantir que o server está pronto
    timer = threading.Timer(1.0, open_browser)
    timer.start()

    # Roda o caçador de zumbis silenciosamente no início
    from utils.excel_killer import clean_zombie_excels
    clean_zombie_excels(force=True)

    yield

    # Limpeza na saída
    clean_zombie_excels(force=True)

app = FastAPI(title="Planify Web-Local", lifespan=lifespan)

# Monta arquivos estáticos
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", response_class=HTMLResponse)
async def read_root():
    index_path = Path("static/index.html")
    if index_path.exists():
        with open(index_path, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>Planify - Arquivo index.html não encontrado na pasta static/</h1>"

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    # Mock do retorno para Fase 1
    return JSONResponse(content={"status": "sucesso", "arquivo": file.filename})

def iniciar():
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=False)

if __name__ == "__main__":
    iniciar()
