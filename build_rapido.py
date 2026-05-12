import os
import subprocess

print("Iniciando build rapido do Planify com PyInstaller...")

# Instala o PyInstaller se não tiver
os.system("pip install pyinstaller")

# Montando o comando do PyInstaller
comando = [
    "pyinstaller",
    "--noconfirm",      # Substitui a pasta dist/ antiga sem perguntar
    "--onedir",         # Cria uma pasta com o executável (muito mais rápido para testar)
    "--windowed",       # Esconde o terminal preto do Python
    "--icon=assets/icon.ico",

    # Adicionando pastas de arquivos não-Python (imagens, jsons, planilhas)
    "--add-data=assets;assets",
    "--add-data=config;config",

    # Imports ocultos que o PyInstaller às vezes não acha sozinho
    "--hidden-import=pandas",
    "--hidden-import=customtkinter",
    "--hidden-import=tkinterdnd2",
    "--hidden-import=openpyxl",
    "--hidden-import=rapidfuzz",
    "--hidden-import=pydantic",
    "--hidden-import=psutil",

    "--name=Planify",
    "main.py"
]

# Executando o comando
print("Executando o PyInstaller (isso deve levar cerca de 1 minuto)...")
subprocess.run(" ".join(comando), shell=True)

print("\n" + "=" * 50)
print("BUILD CONCLUIDO!")
print("Abra a pasta 'dist/Planify' e rode o Planify.exe")
print("=" * 50)