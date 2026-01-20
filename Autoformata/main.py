import sys
import os

# Garante que o Python encontre as pastas do projeto (core, ui, utils)
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def iniciar():
    print("🚀 Iniciando SISORC...")
    
    try:
        # Tenta importar as bibliotecas. Se falhar, avisa o usuário.
        import customtkinter
        import pandas
        import openpyxl
        
        # Importa e inicia a interface
        from ui.main_window import SisorcApp
        
        app = SisorcApp()
        app.mainloop()
        
    except ImportError as e:
        print("\n❌ ERRO: Faltam bibliotecas necessárias.")
        print(f"Detalhe: {e}")
        print("\nPara corrigir, abra o terminal e digite:")
        print("pip install pandas openpyxl customtkinter pillow")
        input("\nPressione ENTER para sair...")
        
    except Exception as e:
        print(f"\n❌ Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
        input("\nPressione ENTER para sair...")

if __name__ == "__main__":
    iniciar()