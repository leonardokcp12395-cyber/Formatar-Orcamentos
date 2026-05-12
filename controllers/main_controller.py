import os
import gc
import shutil
import threading
import json
import time
import queue
from pathlib import Path
import pandas as pd
import win32com.client
import pythoncom

from core.excel_handler import OrcamentoEngine
from core.database import DatabaseManager
from core.paths import get_app_dir

from utils.logger import Logger
from utils.config_manager import ConfigManager
from utils.autocomplete_manager import AutocompleteManager
from utils.smart_parser import SmartParser
from utils.pdf_exporter import PDFExporter
from utils.template_manager import TemplateManager


class MainController:
    """
    Controller MVC — NÃO detém referências directas à UI.
    Comunica com a View exclusivamente via queue.Queue.
    A View deve chamar schedule_queue_poll() após construção para iniciar o polling.
    """

    def __init__(self, ui_queue: queue.Queue, schedule_fn):
        """
        Args:
            ui_queue: Fila partilhada onde o controller publica eventos para a View.
            schedule_fn: Função da View para agendar callbacks na Main Thread (ex: self.after).
        """
        self.ui_queue = ui_queue
        self._schedule = schedule_fn

        self.logger = Logger("Planify")
        self.config_manager = ConfigManager()
        self.dados_config = self.config_manager.load_profiles()
        self.autocomplete = AutocompleteManager()
        self.template_manager = TemplateManager()

        db_path = get_app_dir() / 'planify_history.db'
        db_config = {'database': {'nome_arquivo': str(db_path)}}
        self.db_manager = DatabaseManager(db_config)

        self.sintetico_original_path = ""
        self.sintetico_limpo_path = ""
        self.modelo_path = ""

    def schedule_queue_poll(self):
        """Inicia o polling da fila UI. Chamar UMA VEZ após a View estar pronta."""
        self._schedule(100, self._process_ui_queue)

    def _process_ui_queue(self):
        """Processa todos os eventos pendentes na fila e reagenda-se."""
        try:
            while True:
                msg = self.ui_queue.get_nowait()
                # Publica o evento directamente — a View regista handlers no dict
                action = msg.get('action')
                handler = msg.get('_handler')
                if handler:
                    handler(msg)
        except queue.Empty:
            pass
        finally:
            self._schedule(100, self._process_ui_queue)

    # ──────────────────────────────────────────────
    #  SESSÃO
    # ──────────────────────────────────────────────

    def _get_session_path(self):
        return get_app_dir() / "config" / "last_session.json"

    def limpar_dados_sessao(self):
        self.sintetico_original_path = ""
        self.sintetico_limpo_path = ""
        try:
            sess_path = self._get_session_path()
            if sess_path.exists():
                os.remove(sess_path)
        except Exception as e:
            self.logger.error(f"Erro ao limpar sessão: {e}")

    def salvar_sessao_atual(self, data):
        try:
            sess_path = self._get_session_path()
            sess_path.parent.mkdir(parents=True, exist_ok=True)
            with open(sess_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False)
        except Exception as e:
            print(f"Erro ao salvar sessão: {e}")

    def carregar_ultima_sessao(self):
        path = self._get_session_path()
        if not path.exists():
            return {}
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Erro ao carregar sessão: {e}")
            return {}

    # ──────────────────────────────────────────────
    #  LEITURA SEGURA (Win32COM Blindado)
    # ──────────────────────────────────────────────

    def iniciar_leitura_segura(self, original_path, on_success, on_error):
        """
        Inicia limpeza do ficheiro SIPAC em background thread.
        Args:
            on_success: callable(msg_dict) chamado na Main Thread quando sucesso
            on_error: callable(msg_dict) chamado na Main Thread quando erro
        """
        self.sintetico_original_path = original_path
        threading.Thread(
            target=self._limpar_planilha_sipac,
            args=(original_path, on_success, on_error),
            daemon=True
        ).start()

    def _limpar_planilha_sipac(self, caminho_original, on_success, on_error):
        """Abre o ficheiro no próprio Excel invisível e guarda uma cópia limpa e sem erros."""
        temp_dir = get_app_dir() / "Output"
        temp_dir.mkdir(exist_ok=True)

        caminho_copia = str(temp_dir / "temp_original_desbloqueado.xlsx")
        try:
            shutil.copy2(caminho_original, caminho_copia)
        except Exception as e:
            self.ui_queue.put({
                'action': 'limpar_planilha_erro',
                '_handler': on_error,
                'erro_msg': f"Falha ao tirar bloqueio de segurança: {e}"
            })
            return

        caminho_limpo = str(temp_dir / "temp_sintetico_limpo.xlsx")
        if os.path.exists(caminho_limpo):
            try:
                os.remove(caminho_limpo)
            except Exception:
                pass

        excel = None
        wb = None
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            caminho_abs = os.path.abspath(caminho_copia)
            wb = excel.Workbooks.Open(
                caminho_abs, UpdateLinks=False, ReadOnly=True)
            wb.SaveAs(os.path.abspath(caminho_limpo), FileFormat=51)

            wb.Close(False)
            wb = None
            excel.Quit()
            excel = None

            self.sintetico_limpo_path = caminho_limpo
            self.ui_queue.put({
                'action': 'limpar_planilha_sucesso',
                '_handler': on_success,
                'path_limpo': caminho_limpo
            })
        except Exception as e:
            self.ui_queue.put({
                'action': 'limpar_planilha_erro',
                '_handler': on_error,
                'erro_msg': f"Erro COM do Windows: {str(e)}"
            })
        finally:
            # Win32COM Blindado: garante fecho no finally
            if wb:
                try:
                    wb.Close(False)
                except Exception:
                    pass
            if excel:
                try:
                    excel.Quit()
                except Exception:
                    pass
            # Liberta referências COM explicitamente
            del wb
            del excel
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            gc.collect()

    # ──────────────────────────────────────────────
    #  LEITURA DE COLUNAS
    # ──────────────────────────────────────────────

    def ler_colunas(self, line_num):
        if not self.sintetico_limpo_path:
            return []
        df = None
        try:
            df = pd.read_excel(self.sintetico_limpo_path,
                               header=line_num, nrows=5)
            cols = [str(c).strip() for c in df.columns if "Unnamed" not in str(c)]
            return cols
        except Exception as e:
            print(f"Erro ao ler colunas: {e}")
            return []
        finally:
            if df is not None:
                del df
                gc.collect()

    # ──────────────────────────────────────────────
    #  PREVIEW
    # ──────────────────────────────────────────────

    def carregar_preview(self, line_num, m_item, m_desc, m_cod, m_banco, m_unit,
                         on_success, on_error):
        if not self.sintetico_limpo_path:
            self.ui_queue.put({
                'action': 'carregar_preview_erro',
                '_handler': on_error,
                'erro_msg': 'Selecione o sintético primeiro.'
            })
            return

        threading.Thread(
            target=self._ler_dados_preview,
            args=(line_num, m_item, m_desc, m_cod, m_banco, m_unit, on_success, on_error),
            daemon=True
        ).start()

    def _ler_dados_preview(self, line_num, m_item, m_desc, m_cod, m_banco, m_unit,
                           on_success, on_error):
        df = None
        try:
            df = pd.read_excel(self.sintetico_limpo_path, header=line_num)
            df.columns = [str(c).strip() for c in df.columns]

            # Radar Inteligente: para quando encontra rodapé do orçamento
            palavras_parada = ["TOTAL SEM BDI", "TOTAL DO BDI",
                               "TOTAL GERAL", "VALOR GLOBAL", "CUSTO TOTAL"]

            dados_linhas = []

            for idx, row in df.iterrows():
                desc_val = str(row.get(m_desc, 'nan')).strip()
                desc_upper = desc_val.upper()

                if any(p in desc_upper for p in palavras_parada):
                    self.logger.info(
                        f"🛑 Fim do orçamento detetado pelo radar na linha {line_num + idx + 2}.")
                    break

                if desc_val == 'nan' or desc_val == '' or desc_val == 'None':
                    continue

                unit_val = row.get(m_unit, 0)

                dados_linhas.append({
                    'index_excel': line_num + idx + 2,
                    'item_val': row.get(m_item, ''),
                    'desc_val': desc_val,
                    'cod_val': row.get(m_cod, ''),
                    'banco_val': row.get(m_banco, ''),
                    'raw_row_data': row.to_dict(),
                    'unit_val': unit_val
                })

            self.ui_queue.put({
                'action': 'carregar_preview_sucesso',
                '_handler': on_success,
                'dados_linhas': dados_linhas
            })
        except Exception as e:
            self.ui_queue.put({
                'action': 'carregar_preview_erro',
                '_handler': on_error,
                'erro_msg': f"Ocorreu um erro ao carregar a tabela:\n{str(e)}"
            })
        finally:
            # Garbage Collection explícito
            if df is not None:
                del df
                gc.collect()

    # ──────────────────────────────────────────────
    #  GERAÇÃO DE ORÇAMENTO
    # ──────────────────────────────────────────────

    def gerar_orcamento(self, d, m, info, modelo_path, on_progress, on_success, on_error):
        threading.Thread(
            target=self._run_orcamento,
            args=(d, m, info, modelo_path, on_progress, on_success, on_error),
            daemon=True
        ).start()

    def _run_orcamento(self, d, m, p, modelo_path, on_progress, on_success, on_error):
        start_time = time.time()
        eng = OrcamentoEngine({})

        def progress_callback(pct):
            self.ui_queue.put({
                'action': 'gerar_orcamento_progresso',
                '_handler': on_progress,
                'percent': pct
            })

        ok, msg, extra_info = eng.gerar_excel_final(d, modelo_path, m, p, progress_callback)

        pdf_msg = ""
        if ok and p.get("gerar_pdf", 0) == 1:
            self.logger.info("Iniciando conversão para PDF...")
            ok_pdf, path_pdf, log_pdf = PDFExporter.converter_para_pdf(msg)
            if ok_pdf:
                self.logger.info(f"✅ PDF Gerado: {path_pdf}")
                pdf_msg = f"\n\nPDF também gerado:\n{path_pdf}"
            else:
                self.logger.error(f"Erro no PDF: {log_pdf}")

        end_time = time.time()
        duration = end_time - start_time

        if ok:
            try:
                dados_historico = {
                    'data_geracao': p.get('data'),
                    'nome_obra': p.get('nome_arquivo'),
                    'local': f"{p.get('campus')} - {p.get('setor')}",
                    'bdi': p.get('bdi'),
                    'valor_total': 0.0,
                    'arquivo_saida': msg,
                    'num_itens': len(d),
                    'num_titulos': sum(1 for x in d if x.get('_NIVEL_FORCADO') != 'ITEM'),
                    'duracao_processamento': round(duration, 2)
                }
                self.db_manager.inserir_orcamento(dados_historico)
                self.logger.info("✅ Histórico salvo no banco de dados.")
            except Exception as e:
                self.logger.error(f"Erro ao salvar histórico: {e}")

            self.ui_queue.put({
                'action': 'gerar_orcamento_sucesso',
                '_handler': on_success,
                'msg': msg,
                'duration': duration,
                'info_data': p,
                'raw_data': d,
                'pdf_msg': pdf_msg
            })
        else:
            self.ui_queue.put({
                'action': 'gerar_orcamento_erro',
                '_handler': on_error,
                'msg': msg
            })

    # ──────────────────────────────────────────────
    #  SMART PARSER
    # ──────────────────────────────────────────────

    def extrair_dados_texto(self, texto):
        return SmartParser.parse_whatsapp_text(texto, self.autocomplete)
