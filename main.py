import sys
import os
import shutil
import tkinter as tk
from pathlib import Path
import threading


# Configura estrutura de pastas automaticamente
def setup_environment():
    root = Path(os.path.dirname(os.path.abspath(__file__)))

    # 1. Pastas Obrigatórias
    folders = ['config', 'core', 'ui', 'utils', 'Output']
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


def show_splash():
    """Mostra uma tela de carregamento leve (Tkinter puro) enquanto carrega as libs pesadas"""
    splash = tk.Tk()
    splash.overrideredirect(True)  # Remove bordas da janela

    # Centraliza
    w, h = 400, 100
    ws = splash.winfo_screenwidth()
    hs = splash.winfo_screenheight()
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    splash.geometry(f'{w}x{h}+{int(x)}+{int(y)}')

    # Estilo
    splash.configure(bg='#2B2B2B')
    label = tk.Label(splash, text="🚀 Carregando Planify...\nPor favor, aguarde.",
                     fg='white', bg='#2B2B2B', font=("Arial", 12))
    label.pack(expand=True)

    # Barra de loading fake
    canvas = tk.Canvas(splash, height=5, width=400, bg="#444", highlightthickness=0)
    canvas.pack(side="bottom", fill="x")
    rect = canvas.create_rectangle(0, 0, 0, 5, fill="#00AA00", width=0)

    def animate(width=0):
        # CORREÇÃO: Verifica se a janela ainda existe antes de tentar animar
        if not canvas.winfo_exists():
            return

        if width < 400:
            width += 5
            canvas.coords(rect, 0, 0, width, 5)
            splash.after(20, lambda: animate(width))

    animate()
    splash.update()
    return splash


def iniciar():
    setup_environment()

    # --- INÍCIO TELA HIGH-DPI ---
    try:
        import ctypes
        # Habilita suporte a DPI alto no Windows 10/11 para evitar borrado
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    # --- FIM TELA HIGH-DPI ---

    # Mostra Splash Screen IMEDIATAMENTE
    splash = show_splash()

    try:
        # Importações pesadas acontecem AQUI
        import customtkinter
        import pandas
        import openpyxl

        # Importa a App Principal
        from ui.main_window import PlanifyApp

        # Destroi splash e abre app real
        if splash.winfo_exists():
            splash.destroy()

        # Roda o caçador de zumbis silenciosamente no início
        from utils.excel_killer import clean_zombie_excels
        clean_zombie_excels(force=True)

        app = PlanifyApp()
        app.mainloop()

        # Limpeza na saída
        clean_zombie_excels(force=True)

    except ImportError as e:
        # CORREÇÃO: winfo_exists() impede o TclError
        if 'splash' in locals() and splash.winfo_exists():
            splash.destroy()
        print("\n❌ ERRO CRÍTICO: Faltam bibliotecas.")
        print(f"Detalhe: {e}")
        print("\nPor favor, execute o comando abaixo no terminal para instalar as dependências necessárias:")
        print("pip install -r requirements.txt")
        input("\nENTER para sair...")
    except Exception as e:
        if 'splash' in locals() and splash.winfo_exists():
            splash.destroy()
        print(f"\n❌ Erro Inesperado: {e}")
        import traceback
        traceback.print_exc()
        input("\nENTER para sair...")


if __name__ == "__main__":
    iniciar()