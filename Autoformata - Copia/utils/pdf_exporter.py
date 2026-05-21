import os
import time
from pathlib import Path

# Tenta importar win32com para comunicação com Excel
try:
    import win32com.client
    import pythoncom # Necessário para rodar em threads
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

class PDFExporter:
    @staticmethod
    def converter_para_pdf(caminho_excel):
        """
        Converte um arquivo Excel (.xlsx) para PDF usando a API nativa do Windows.
        Retorna: (Sucesso: bool, CaminhoPDF: str, Mensagem: str)
        """
        if not HAS_WIN32:
            return False, "", "Biblioteca pywin32 não instalada ou erro de importação."

        # Garante caminhos absolutos e normalizados para o Windows
        try:
            path_obj = Path(caminho_excel).resolve()
            caminho_abs_excel = str(path_obj)
            caminho_pdf = str(path_obj.with_suffix('.pdf'))
        except Exception as e:
            return False, "", f"Erro de caminho: {e}"

        # 1. Espera tática: Dá tempo do disco liberar o arquivo recém-salvo
        time.sleep(2)

        excel = None
        wb = None
        try:
            # Inicializa contexto COM
            pythoncom.CoInitialize()
            
            # 2. Usa DispatchEx em vez de Dispatch
            # Isso força uma NOVA instância do Excel, evitando conflito com planilhas abertas
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            # Abre o workbook
            wb = excel.Workbooks.Open(caminho_abs_excel)
            
            # Garante que a primeira planilha está ativa
            wb.Worksheets(1).Activate()
            
            # Exporta para PDF (0 = xlTypePDF)
            wb.ExportAsFixedFormat(0, caminho_pdf)
            
            return True, caminho_pdf, "PDF gerado com sucesso"

        except Exception as e:
            # Tenta pegar o erro detalhado do Excel se disponível
            err_msg = str(e)
            if hasattr(e, 'excepinfo') and e.excepinfo:
                err_msg = f"{e.excepinfo[2]} (Código: {e.excepinfo[5]})"
            
            return False, "", f"Erro na conversão PDF: {err_msg}"
            
        finally:
            # Fecha tudo com cuidado extremo
            if wb:
                try: wb.Close(False)
                except: pass
            if excel:
                try: excel.Quit()
                except: pass
            
            # Libera memória COM
            try: pythoncom.CoUninitialize()
            except: pass