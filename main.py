"""
SISORC ULTIMATE v3.0
Sistema de Orçamentação Automatizado
Entry Point - Verifica dependências e inicializa aplicação
"""

import sys
import subprocess
import os
from pathlib import Path

def verificar_e_instalar_dependencias():
    """Verifica e instala todas as dependências necessárias"""
    dependencias = [
        'pandas',
        'openpyxl',
        'customtkinter',
        'pillow'  # Requerido pelo CustomTkinter
    ]
    
    print("🔍 Verificando dependências...")
    
    for pacote in dependencias:
        try:
            __import__(pacote)
            print(f"✅ {pacote} já instalado")
        except ImportError:
            print(f"📦 Instalando {pacote}...")
            subprocess.check_call([
                sys.executable, 
                "-m", 
                "pip", 
                "install", 
                pacote
            ])
            print(f"✅ {pacote} instalado com sucesso")

def obter_caminho_base():
    """Retorna o caminho base da aplicação (útil para PyInstaller)"""
    if getattr(sys, 'frozen', False):
        # Rodando como executável compilado
        return Path(sys._MEIPASS)
    else:
        # Rodando como script Python
        return Path(__file__).parent

if __name__ == "__main__":
    # Configuração inicial
    BASE_DIR = obter_caminho_base()
    
    # Verifica dependências
    verificar_e_instalar_dependencias()
    
    # Importa e executa a aplicação
    print("\n🚀 Iniciando SISORC ULTIMATE...")
    
    try:
        from ui.main_window import SisorcApp
        
        app = SisorcApp()
        app.mainloop()
        
    except Exception as e:
        print(f"\n❌ ERRO CRÍTICO: {str(e)}")
        input("Pressione ENTER para sair...")
        sys.exit(1)