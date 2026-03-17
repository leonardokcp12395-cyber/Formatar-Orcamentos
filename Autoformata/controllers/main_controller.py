import os
import shutil
import threading
import json
import time
from pathlib import Path
import pandas as pd
import win32com.client
import pythoncom
import queue
import gc
from rapidfuzz import process, fuzz

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
    def __init__(self, view):
        self.view = view

        self.logger = Logger("SISORC")
        self.config_manager = ConfigManager()
        self.dados_config = self.config_manager.load_profiles()
        self.autocomplete = AutocompleteManager()
        self.template_manager = TemplateManager()

        db_path = get_app_dir() / 'sisorc_history.db'
        db_config = {'database': {'nome_arquivo': str(db_path)}}
        self.db_manager = DatabaseManager(db_config)

        self.sintetico_original_path = ""
        self.sintetico_limpo_path = ""
        self.modelo_path = ""

        # Queue to safely communicate from background threads to the main UI thread
        self.ui_queue = queue.Queue()
        self.view.after(100, self._process_ui_queue)

    def _process_ui_queue(self):
        try:
            while True:
                msg = self.ui_queue.get_nowait()
                action = msg.get('action')
                if action == 'limpar_planilha_sucesso':
                    self.view._on_limpar_planilha_sucesso(msg['path_limpo'])
                elif action == 'limpar_planilha_erro':
                    self.view._on_limpar_planilha_erro(msg['erro_msg'])
                elif action == 'carregar_preview_sucesso':
                    self.view._on_carregar_preview_sucesso(msg['dados_linhas'])
                elif action == 'carregar_preview_erro':
                    self.view._on_carregar_preview_erro(msg['erro_msg'])
                elif action == 'gerar_orcamento_sucesso':
                    self.view._on_gerar_orcamento_sucesso(
                        msg['msg'], msg['duration'], msg['info_data'], msg['raw_data'], msg['pdf_msg'])
                elif action == 'gerar_orcamento_erro':
                    self.view._on_gerar_orcamento_erro(msg['msg'])
                elif action == 'gerar_orcamento_progresso':
                    self.view._on_gerar_orcamento_progresso(msg['percent'])
        except queue.Empty:
            pass
        finally:
            self.view.after(100, self._process_ui_queue)

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

    def iniciar_leitura_segura(self, original_path):
        self.sintetico_original_path = original_path
        threading.Thread(target=self._limpar_planilha_sipac,
                         args=(original_path,), daemon=True).start()

    def _limpar_planilha_sipac(self, caminho_original):
        """Abre o ficheiro no próprio Excel invisível e guarda uma cópia limpa e sem erros"""
        temp_dir = get_app_dir() / "Output"
        temp_dir.mkdir(exist_ok=True)

        caminho_copia = str(temp_dir / "temp_original_desbloqueado.xlsx")
        try:
            shutil.copy2(caminho_original, caminho_copia)
        except Exception as e:
            self.ui_queue.put({'action': 'limpar_planilha_erro',
                              'erro_msg': f"Falha ao tirar bloqueio de segurança: {e}"})
            return

        caminho_limpo = str(temp_dir / "temp_sintetico_limpo.xlsx")
        if os.path.exists(caminho_limpo):
            try:
                os.remove(caminho_limpo)
            except:
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
            excel.Quit()

            self.sintetico_limpo_path = caminho_limpo
            self.ui_queue.put(
                {'action': 'limpar_planilha_sucesso', 'path_limpo': caminho_limpo})
        except Exception as e:
            self.ui_queue.put({'action': 'limpar_planilha_erro',
                              'erro_msg': f"Erro COM do Windows: {str(e)}"})
        finally:
            if wb:
                try:
                    wb.Close(False)
                except:
                    pass
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def ler_colunas(self, line_num):
        if not self.sintetico_limpo_path:
            return [], {}
        try:
            df = pd.read_excel(self.sintetico_limpo_path, header=line_num, nrows=5)
            cols = [str(c).strip() for c in df.columns if "Unnamed" not in str(c)]
            del df
            gc.collect()

            # Smart fuzzy matching
            expected_keys = {
                "ITEM": ["ITEM", "IT", "NUM"],
                "CODIGO": ["CODIGO", "COD", "CÓDIGO"],
                "BANCO": ["BANCO", "FONTE", "REF", "REFERENCIA", "REFERÊNCIA"],
                "DESCRICAO": ["DESCRICAO", "DESCRIÇÃO", "DESC", "SERVICO", "OBJETO", "DISCRIMINAÇÃO"],
                "UNID": ["UNID", "UND", "UNIDADE", "U."],
                "QUANT": ["QUANT", "QTD", "QUANTIDADE"],
                "UNIT": ["UNIT", "VALOR UNIT", "VALOR UNITÁRIO", "PREÇO", "PRECO UNIT"]
            }

            best_matches = {}
            for key, keywords in expected_keys.items():
                best_match = None
                best_score = 0
                for keyword in keywords:
                    match = process.extractOne(keyword, cols, scorer=fuzz.token_set_ratio)
                    if match:
                        col_name, score, _ = match
                        if score > best_score and score >= 60:
                            best_score = score
                            best_match = col_name
                if best_match:
                    best_matches[key] = best_match

            return cols, best_matches
        except Exception as e:
            self.logger.error(f"Erro ao ler colunas: {e}")
            return [], {}

    def carregar_preview(self, line_num, m_item, m_desc, m_cod, m_banco, m_unit):
        if not self.sintetico_limpo_path:
            self.ui_queue.put({'action': 'carregar_preview_erro',
                              'erro_msg': 'Selecione o sintético primeiro.'})
            return

        threading.Thread(target=self._ler_dados_preview, args=(
            line_num, m_item, m_desc, m_cod, m_banco, m_unit), daemon=True).start()

    def _ler_dados_preview(self, line_num, m_item, m_desc, m_cod, m_banco, m_unit):
        try:
            df = pd.read_excel(self.sintetico_limpo_path, header=line_num)
            df.columns = [str(c).strip() for c in df.columns]

            palavras_parada = ["TOTAL SEM BDI", "TOTAL DO BDI",
                               "TOTAL GERAL", "VALOR GLOBAL", "CUSTO TOTAL"]

            dados_linhas = []

            for idx, row in df.iterrows():
                desc_val = str(row.get(m_desc, 'nan')).strip()
                desc_upper = desc_val.upper()

                if any(p in desc_upper for p in palavras_parada):
                    self.logger.info(
                        f"🛑 Fim do orçamento detetado pelo radar na linha {line_num+idx+2}.")
                    break

                if desc_val == 'nan' or desc_val == '' or desc_val == 'None':
                    continue

                unit_val = row.get(m_unit, 0)

                # We need to pass row as a dict so it can be safely sent over queue
                dados_linhas.append({
                    'index_excel': line_num+idx+2,
                    'item_val': row.get(m_item, ''),
                    'desc_val': desc_val,
                    'cod_val': row.get(m_cod, ''),
                    'banco_val': row.get(m_banco, ''),
                    'raw_row_data': row.to_dict(),
                    'unit_val': unit_val
                })

            del df
            gc.collect()

            self.ui_queue.put(
                {'action': 'carregar_preview_sucesso', 'dados_linhas': dados_linhas})
        except Exception as e:
            self.ui_queue.put({'action': 'carregar_preview_erro',
                              'erro_msg': f"Ocorreu um erro ao carregar a tabela:\n{str(e)}"})

    def gerar_orcamento(self, d, m, info, modelo_path):
        threading.Thread(target=self._run_orcamento, args=(
            d, m, info, modelo_path), daemon=True).start()

    def _run_orcamento(self, d, m, p, modelo_path):
        start_time = time.time()
        eng = OrcamentoEngine({})

        def progress_callback(pct):
            self.ui_queue.put({'action': 'gerar_orcamento_progresso', 'percent': pct})

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
                'msg': msg,
                'duration': duration,
                'info_data': p,
                'raw_data': d,
                'pdf_msg': pdf_msg
            })
        else:
            self.ui_queue.put({
                'action': 'gerar_orcamento_erro',
                'msg': msg
            })

    def extrair_dados_texto(self, texto):
        return SmartParser.parse_whatsapp_text(texto, self.autocomplete)
