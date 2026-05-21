import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
from tkinterdnd2 import TkinterDnD, DND_FILES
import os
import queue
from pathlib import Path
from tkinter import ttk
import tkinter as tk

from controllers.main_controller import MainController
from ui.components.top_dashboard import TopDashboard
from ui.components.side_panel import SidePanel
from ui.components.config_panel import ConfigPanel
from ui.components.excel_preview import LevelSelector

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


# ═══════════════════════════════════════════════════════
#  JANELAS AUXILIARES (Modais)
# ═══════════════════════════════════════════════════════

class TemplateEditor(ctk.CTkToplevel):
    """Janela para Adicionar/Remover Modelos"""

    def __init__(self, parent, manager, callback_refresh):
        super().__init__(parent)
        self.title("⚙️ Gerenciador de Modelos")
        self.geometry("500x400")
        self.manager = manager
        self.callback_refresh = callback_refresh
        self.path_temp = ""

        self.transient(parent)
        self.grab_set()
        self._setup_ui()

    def _setup_ui(self):
        f_new = ctk.CTkFrame(self)
        f_new.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(f_new, text="Adicionar Novo Modelo",
                     font=("Arial", 12, "bold")).pack(pady=5)

        f_btn = ctk.CTkFrame(f_new, fg_color="transparent")
        f_btn.pack(fill="x")

        self.btn_file = ctk.CTkButton(
            f_btn, text="📂 Selecionar Excel (.xlsx)", command=self._pick_file)
        self.btn_file.pack(side="left", padx=5, fill="x", expand=True)

        self.lbl_file = ctk.CTkLabel(
            f_new, text="Nenhum arquivo selecionado", text_color="gray", font=("Arial", 10))
        self.lbl_file.pack(pady=2)

        f_inputs = ctk.CTkFrame(f_new, fg_color="transparent")
        f_inputs.pack(fill="x", pady=5)

        self.ent_name = ctk.CTkEntry(
            f_inputs, placeholder_text="Nome (ex: PADRAO 2026)")
        self.ent_name.pack(side="left", padx=5, fill="x", expand=True)

        self.ent_start = ctk.CTkEntry(
            f_inputs, width=80, placeholder_text="Linha Início")
        self.ent_start.insert(0, "25")
        self.ent_start.pack(side="right", padx=5)

        ctk.CTkButton(f_new, text="💾 Salvar Modelo", command=self._save,
                      fg_color="green").pack(fill="x", padx=5, pady=5)

        ctk.CTkLabel(self, text="Modelos Salvos:", font=(
            "Arial", 12, "bold")).pack(pady=(10, 5))
        self.scroll = ctk.CTkScrollableFrame(self)
        self.scroll.pack(fill="both", expand=True, padx=10, pady=10)
        self._load_list()

    def _pick_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p:
            self.path_temp = p
            self.lbl_file.configure(
                text=os.path.basename(p), text_color="lime")

    def _save(self):
        name = self.ent_name.get().strip()
        start = self.ent_start.get()
        if not name or not self.path_temp:
            return messagebox.showwarning("Aviso", "Selecione um arquivo e dê um nome.")

        ok, msg = self.manager.add_template(name, self.path_temp, start)
        if ok:
            messagebox.showinfo("Sucesso", msg)
            self.ent_name.delete(0, 'end')
            self.path_temp = ""
            self.lbl_file.configure(
                text="Nenhum arquivo selecionado", text_color="gray")
            self._load_list()
            self.callback_refresh()
        else:
            messagebox.showerror("Erro", msg)

    def _load_list(self):
        for w in self.scroll.winfo_children():
            w.destroy()
        names = self.manager.get_template_names()
        for n in names:
            f = ctk.CTkFrame(self.scroll, fg_color="transparent")
            f.pack(fill="x", pady=2)
            ctk.CTkLabel(f, text=n).pack(side="left", padx=5)
            ctk.CTkButton(f, text="🗑️", width=30, fg_color="red",
                          command=lambda x=n: self._del(x)).pack(side="right")

    def _del(self, name):
        if messagebox.askyesno("Confirmar", f"Deletar modelo '{name}'?"):
            self.manager.remove_template(name)
            self._load_list()
            self.callback_refresh()


class DatabaseEditor(ctk.CTkToplevel):
    def __init__(self, parent, manager, callback_refresh):
        super().__init__(parent)
        self.title("📝 Editor de Listas")
        self.geometry("500x600")
        self.manager = manager
        self.callback_refresh = callback_refresh
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self._setup_ui()

    def _setup_ui(self):
        top_frame = ctk.CTkFrame(self)
        top_frame.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(top_frame, text="Selecione a Lista para Editar:",
                     font=("Arial", 14, "bold")).pack(pady=5)

        self.cats = {
            "Campi (Campus)": "campus", "Setores": "setor", "Servidores": "servidor",
            "Elaboradores": "elaborador", "Estagiários": "estagiario", "Fiscais": "fiscal"
        }
        self.combo_cat = ctk.CTkComboBox(top_frame, values=list(
            self.cats.keys()), command=self._carregar_lista, width=300)
        self.combo_cat.pack(pady=5)

        self.scroll = ctk.CTkScrollableFrame(self, label_text="Itens Salvos")
        self.scroll.pack(fill="both", expand=True, padx=10, pady=5)
        ctk.CTkButton(self, text="Concluir", command=self.destroy,
                      fg_color="gray").pack(pady=10)
        self._carregar_lista(list(self.cats.keys())[0])

    def _carregar_lista(self, cat_friendly):
        for widget in self.scroll.winfo_children():
            widget.destroy()
        key = self.cats[cat_friendly]
        items = self.manager.get_list(key)
        if not items:
            ctk.CTkLabel(self.scroll, text="(Lista Vazia)",
                         text_color="gray").pack(pady=20)
            return
        for item in items:
            row = ctk.CTkFrame(self.scroll, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkButton(row, text="🗑️", width=40, fg_color="#C0392B", hover_color="#E74C3C",
                          command=lambda k=key, i=item: self._deletar_item(k, i)).pack(side="right", padx=5)
            ctk.CTkLabel(row, text=item, anchor="w").pack(
                side="left", padx=5, fill="x", expand=True)

    def _deletar_item(self, key, item):
        if self.manager.remove_value(key, item):
            self._carregar_lista(self.combo_cat.get())
            self.callback_refresh()


# ═══════════════════════════════════════════════════════
#  APLICAÇÃO PRINCIPAL
# ═══════════════════════════════════════════════════════

class PlanifyApp(ctk.CTk, TkinterDnD.DnDWrapper):
    """
    Janela principal Planify — Orquestrador MVC.
    Não acede directamente aos widgets internos dos componentes.
    Usa as APIs get_data()/set_data()/limpar_campos() de cada painel.
    """

    def __init__(self):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)

        self.title("Planify - Engenharia & Orçamentos")
        self.geometry("1024x720")

        # Drag & Drop
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self._on_drop)

        # Queue partilhada para thread-safety
        self._ui_queue = queue.Queue()

        # Controller (sem referência directa à View)
        self.controller = MainController(self._ui_queue, self.after)

        # Constrói UI
        self._setup_ui()

        # Liga logger à log box
        self.controller.logger.adicionar_callback(self._log_callback)

        # Inicia polling da queue
        self.controller.schedule_queue_poll()

        # Carrega estado
        self._carregar_perfil_padrao()
        self._atualizar_listas_visuais()
        self._atualizar_combo_modelos()
        self._carregar_ultima_sessao()

        # Atalho global
        self.bind('<Control-Return>', lambda e: self.executar())

    # ──────────────────────────────────────────────
    #  DRAG & DROP
    # ──────────────────────────────────────────────

    def _on_drop(self, event):
        path = event.data
        if path.startswith('{') and path.endswith('}'):
            path = path[1:-1]

        ext = os.path.splitext(path)[1].lower()
        if ext not in ['.xlsx', '.xls', '.xlsm']:
            messagebox.showwarning(
                "Arquivo Inválido", "Apenas arquivos Excel são aceitos!")
            return

        self.top_dashboard.set_file_label(os.path.basename(path), "lime")
        self.controller.logger.info(
            f"📂 Arquivo carregado via Drag & Drop: {os.path.basename(path)}")
        self._iniciar_leitura_segura(path)

    # ──────────────────────────────────────────────
    #  SETUP DA UI
    # ──────────────────────────────────────────────

    def _setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # TOPO: DASHBOARD (callbacks passados como kwargs)
        self.top_dashboard = TopDashboard(
            self,
            on_select_file=self.sel_sintetico,
            on_load_preview=self.carregar_preview,
            on_extract_text=self.extrair_dados_texto,
            on_toggle_theme=self._alternar_tema,
            on_kill_excel=self._limpar_excel_zumbis
        )
        self.top_dashboard.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        # TABVIEW
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=0)

        self.tab_main = self.tabview.add("🏗️ Painel de Orçamento")
        self.tab_config = self.tabview.add("⚙️ Configurações & Mapeamento")

        # Layout: Painel lateral + Tabela gigante
        self.tab_main.grid_columnconfigure(0, weight=0, minsize=320)
        self.tab_main.grid_columnconfigure(1, weight=1)
        self.tab_main.grid_rowconfigure(0, weight=1)

        # PAINEL ESQUERDO (callbacks passados como kwargs)
        self.side_panel = SidePanel(
            self.tab_main,
            on_limpar=self.limpar_dados_sessao,
            on_editor_db=self.abrir_editor_db
        )
        self.side_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)

        # PAINEL DIREITO (Tabela Visual)
        self.table_control = LevelSelector(self.tab_main)
        self.table_control.grid(row=0, column=1, sticky="nsew", pady=0)

        # ABA CONFIGURAÇÕES (callbacks passados como kwargs)
        self.config_panel = ConfigPanel(
            self.tab_config,
            on_model_change=self._ao_trocar_modelo,
            on_manage_models=self._abrir_gerenciador_modelos
        )
        self.config_panel.pack(fill="both", expand=True, padx=5, pady=5)

        # RODAPÉ FIXO
        bot = ctk.CTkFrame(self, height=100)
        bot.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        bot.grid_columnconfigure(0, weight=1)
        bot.grid_columnconfigure(1, weight=1)

        btn_frame = ctk.CTkFrame(bot, fg_color="transparent")
        btn_frame.grid(row=0, column=0, sticky="nsew", padx=10)

        ctk.CTkLabel(btn_frame, text="Dica: Pressione Ctrl + Enter para gerar",
                     font=("Arial", 10), text_color="gray").pack(side="top", pady=(5, 0))

        self.btn_run = ctk.CTkButton(
            btn_frame, text="🚀 GERAR ORÇAMENTO", command=self.executar,
            height=45, fg_color="green", font=("Arial", 16, "bold"))
        self.btn_run.pack(side="top", pady=5, fill="x")

        status_frame = ctk.CTkFrame(btn_frame, fg_color="transparent")
        status_frame.pack(side="top", fill="x")
        self.lbl_status = ctk.CTkLabel(
            status_frame, text="Pronto para uso.", font=("Arial", 11), text_color="gray")
        self.lbl_status.pack(side="left")

        self.progress = ctk.CTkProgressBar(
            status_frame, orientation="horizontal", mode="indeterminate")

        self.log_box = ctk.CTkTextbox(bot, height=90)
        self.log_box.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

        self.log_box.tag_config("ERROR", foreground="red")
        self.log_box.tag_config("SUCCESS", foreground="lime")
        self.log_box.tag_config("INFO", foreground="white")

    # ──────────────────────────────────────────────
    #  ORQUESTRADOR: Pede dados aos painéis
    # ──────────────────────────────────────────────

    def _carregar_perfil_padrao(self):
        prof = self.controller.dados_config.get("perfis", {}).get(
            "PADRAO", {}).get("input", {})
        self.config_panel.set_profile_mapping(prof)

    def _atualizar_listas_visuais(self):
        self.side_panel.atualizar_listas(self.controller.autocomplete)

    def _atualizar_combo_modelos(self):
        nomes = self.controller.template_manager.get_template_names()
        self.config_panel.update_model_list(nomes)

    def _ao_trocar_modelo(self, nome):
        path = self.controller.template_manager.get_template_path(nome)
        if path and os.path.exists(path):
            self.controller.modelo_path = path
            self.controller.logger.info(f"Modelo definido: {nome}")

    def _abrir_gerenciador_modelos(self):
        TemplateEditor(self, self.controller.template_manager,
                       self._atualizar_combo_modelos)

    def _alternar_tema(self):
        if self.top_dashboard.get_theme_state() == 1:
            ctk.set_appearance_mode("Dark")
        else:
            ctk.set_appearance_mode("Light")

    # ──────────────────────────────────────────────
    #  SESSÃO
    # ──────────────────────────────────────────────

    def _salvar_sessao_atual(self):
        data = self.side_panel.get_data()
        self.controller.salvar_sessao_atual(data)

    def _carregar_ultima_sessao(self):
        data = self.controller.carregar_ultima_sessao()
        if not data:
            return
        self.side_panel.set_data(data)
        self.controller.logger.info("Estado da última sessão restaurado.")

    # ──────────────────────────────────────────────
    #  IMPORTAÇÃO INTELIGENTE (WhatsApp)
    # ──────────────────────────────────────────────

    def extrair_dados_texto(self):
        texto = self.top_dashboard.get_import_text()
        if not texto:
            return messagebox.showwarning("Vazio", "Cole o texto do WhatsApp primeiro!")

        dados = self.controller.extrair_dados_texto(texto)
        count = self.side_panel.fill_from_extracted(dados)

        self.top_dashboard.clear_import_text()
        self.controller.logger.info(
            f"Importação Inteligente V2: {count} campos extraídos e normalizados!")
        messagebox.showinfo("Sucesso", f"{count} dados foram extraídos!")

    # ──────────────────────────────────────────────
    #  SELEÇÃO DE FICHEIRO
    # ──────────────────────────────────────────────

    def sel_sintetico(self):
        p = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
        if p:
            self.top_dashboard.set_file_label(Path(p).name, "lime")
            self.controller.logger.info(
                f"📂 Arquivo selecionado manualmente: {os.path.basename(p)}")
            self._iniciar_leitura_segura(p)

    # ──────────────────────────────────────────────
    #  LIMPEZA
    # ──────────────────────────────────────────────

    def limpar_dados_sessao(self):
        self.side_panel.limpar_campos()
        self.top_dashboard.set_file_label("Nenhum arquivo", "gray")
        self.top_dashboard.clear_import_text()
        self.table_control.clear()
        self.config_panel.reset_column_mapping()

        self.controller.limpar_dados_sessao()
        self.controller.logger.info("🧹 Sessão limpa para novo orçamento.")
        self.lbl_status.configure(text="Sessão limpa.", text_color="lime")

    def abrir_editor_db(self):
        DatabaseEditor(self, self.controller.autocomplete,
                       self._atualizar_listas_visuais)

    def _limpar_excel_zumbis(self):
        from utils.excel_killer import clean_zombie_excels
        count = clean_zombie_excels(force=True)
        if count > 0:
            self.lbl_status.configure(text=f"{count} processo(s) Excel eliminados.", text_color="orange")
            self.controller.logger.warning(f"🚨 BOTÃO DE PÂNICO: {count} processo(s) EXCEL.EXE eliminados.")
        else:
            self.lbl_status.configure(text="Nenhum processo Excel travado encontrado.", text_color="lime")
            self.controller.logger.info("Nenhum processo zumbi do Excel encontrado.")

    # ──────────────────────────────────────────────
    #  LEITURA SEGURA + PREVIEW
    # ──────────────────────────────────────────────

    def _iniciar_leitura_segura(self, path):
        self.limpar_dados_sessao()

        self.configure(cursor="watch")
        self.lbl_status.configure(
            text="Limpando corrupções do arquivo SIPAC/SEI...", text_color="orange")

        self.progress.pack(side="bottom", padx=10, pady=(0, 5), fill="x")
        self.progress.start()

        self.controller.iniciar_leitura_segura(
            path,
            on_success=self._on_limpar_planilha_sucesso,
            on_error=self._on_limpar_planilha_erro
        )

    def _on_limpar_planilha_sucesso(self, msg):
        self.progress.stop()
        self.progress.pack_forget()
        self.configure(cursor="")
        self.lbl_status.configure(
            text="Arquivo pronto e limpo!", text_color="gray")
        self._ler_colunas()
        self.carregar_preview()

    def _on_limpar_planilha_erro(self, msg):
        self.progress.stop()
        self.progress.pack_forget()
        self.configure(cursor="")
        messagebox.showerror(
            "Erro Crítico", f"Falha ao limpar o arquivo Excel:\n{msg.get('erro_msg', '')}")
        self.lbl_status.configure(
            text="Erro ao limpar arquivo.", text_color="red")

    def _ler_colunas(self):
        try:
            l = self.config_panel.get_start_line()
            cols = self.controller.ler_colunas(l)
            self.config_panel.update_column_options(cols)
        except Exception:
            pass

    def carregar_preview(self):
        self.configure(cursor="watch")
        self.lbl_status.configure(
            text="Carregando tabela...", text_color="orange")

        self.progress.pack(side="bottom", padx=10, pady=(0, 5), fill="x")
        self.progress.start()

        self._ler_colunas()
        l = self.config_panel.get_start_line()
        m = self.config_panel.get_column_mapping()

        self.table_control.clear()

        self.controller.carregar_preview(
            l, m["ITEM"], m["DESCRICAO"], m["CODIGO"], m["BANCO"], m["UNIT"],
            on_success=self._on_carregar_preview_sucesso,
            on_error=self._on_carregar_preview_erro
        )

    def _on_carregar_preview_sucesso(self, msg):
        self.progress.stop()
        self.progress.pack_forget()
        self.configure(cursor="")

        dados_linhas = msg.get('dados_linhas', [])
        for data in dados_linhas:
            self.table_control.add_row(
                data['index_excel'],
                data['item_val'],
                data['desc_val'],
                data['cod_val'],
                data['banco_val'],
                data['raw_row_data'],
                data['unit_val']
            )

        self.tabview.set("🏗️ Painel de Orçamento")
        self.lbl_status.configure(
            text="Tabela carregada com sucesso.", text_color="lime")

    def _on_carregar_preview_erro(self, msg):
        self.progress.stop()
        self.progress.pack_forget()
        self.configure(cursor="")
        messagebox.showerror("Erro de Leitura", msg.get('erro_msg', ''))
        self.lbl_status.configure(
            text="Erro ao ler ficheiro.", text_color="red")

    # ──────────────────────────────────────────────
    #  GERAÇÃO DE ORÇAMENTO
    # ──────────────────────────────────────────────

    @staticmethod
    def _obter_valor_seguro(valor, is_date=False, custom_placeholder=None):
        """Obtém valor de uma string, ou retorna placeholder."""
        valor = valor.strip().upper() if valor else ""
        if not valor:
            if custom_placeholder:
                return custom_placeholder
            return "xx/xx/xxxx" if is_date else "xxxxxxxxxx"
        return valor

    def executar(self):
        if not self.controller.modelo_path:
            return messagebox.showwarning(
                "Erro", "Nenhum modelo selecionado! Use o botão de engrenagem para adicionar um.")

        d = self.table_control.get_final_data()
        if not d:
            return messagebox.showwarning("Vazio", "Tabela vazia")

        # Recolhe dados dos painéis via API limpa
        m = self.config_panel.get_column_mapping()
        side_data = self.side_panel.get_data()
        config_data = self.config_panel.get_data()

        info = {
            "nome_arquivo": side_data.get("nome_arquivo", "") or "Orcamento",
            "descricao_header": side_data.get("descricao_header", "").upper(),
            "campus": side_data.get("campus", "").upper(),
            "setor": side_data.get("setor", "").upper(),
            "servidor": side_data.get("servidor", "").upper(),
            "elaborador": side_data.get("elaborador", "").upper(),
            "estagiario": side_data.get("estagiario", "").upper(),
            "fiscal": side_data.get("fiscal", "").upper(),
            "data": self._obter_valor_seguro(side_data.get("data", ""), is_date=True),
            "orcafascio": self._obter_valor_seguro(side_data.get("orcafascio", "")),
            "processo": self._obter_valor_seguro(side_data.get("processo", "")),
            "num_orcamento": self._obter_valor_seguro(side_data.get("num_orcamento", ""), custom_placeholder="XX"),
            "empenho": self._obter_valor_seguro(side_data.get("empenho", "")),
            "data_emissao": self._obter_valor_seguro(side_data.get("data_emissao", ""), is_date=True),
            "data_inicio": self._obter_valor_seguro(side_data.get("data_inicio", ""), is_date=True),
            "prazo": self._obter_valor_seguro(side_data.get("prazo", "")),
            "bdi": config_data["bdi"],
            "calc_mode": config_data["calc_mode"],
            "altura_linha": config_data["altura_linha"],
            "gerar_pdf": config_data["gerar_pdf"],
        }

        # Salva autocomplete para as chaves DB
        for key in self.side_panel.get_db_keys():
            self.controller.autocomplete.add_value(key, info.get(key, ""))

        self._atualizar_listas_visuais()
        self._salvar_sessao_atual()

        # UI feedback
        self.btn_run.configure(state="disabled", text="Processando...")
        self.lbl_status.configure(
            text="Iniciando motor de cálculo...", text_color="orange")

        self.progress.configure(mode="determinate")
        self.progress.set(0)
        self.progress.pack(side="bottom", padx=10, pady=(0, 5), fill="x")

        self.controller.gerar_orcamento(
            d, m, info, self.controller.modelo_path,
            on_progress=self._on_gerar_orcamento_progresso,
            on_success=self._on_gerar_orcamento_sucesso,
            on_error=self._on_gerar_orcamento_erro
        )

    def _on_gerar_orcamento_progresso(self, msg):
        pct = msg.get('percent', 0)
        self.progress.set(pct / 100.0)
        self.lbl_status.configure(text=f"Processando: {pct}%")

    def _on_gerar_orcamento_sucesso(self, msg):
        self.progress.pack_forget()
        self.progress.configure(mode="indeterminate")
        self.btn_run.configure(state="normal", text="🚀 GERAR ORÇAMENTO")

        self.lbl_status.configure(
            text="Concluído com sucesso!", text_color="lime")

        file_path = msg.get('msg', '')
        pdf_msg = msg.get('pdf_msg', '')
        messagebox.showinfo(
            "Sucesso", f"Salvo com sucesso em:\n{file_path}{pdf_msg}")
        try:
            os.startfile(file_path)
        except Exception:
            pass

    def _on_gerar_orcamento_erro(self, msg):
        self.progress.pack_forget()
        self.progress.configure(mode="indeterminate")
        self.btn_run.configure(state="normal", text="🚀 GERAR ORÇAMENTO")

        error_msg = msg.get('msg', '')
        self.lbl_status.configure(
            text="Erro no processamento.", text_color="red")
        if "Permission denied" in error_msg or "aberto" in error_msg:
            messagebox.showerror(
                "Arquivo Aberto",
                f"O arquivo parece estar aberto no Excel.\n\n"
                f"Por favor, feche o arquivo '{error_msg}' e tente novamente.")
        else:
            messagebox.showerror("Erro Fatal", error_msg)

    # ──────────────────────────────────────────────
    #  LOG CALLBACK
    # ──────────────────────────────────────────────

    def _log_callback(self, n, m):
        msg_upper = m.upper()
        tag = "INFO"
        if "ERRO" in msg_upper or "❌" in m:
            tag = "ERROR"
        elif "SUCESSO" in msg_upper or "✅" in m or "CONCLUÍDO" in msg_upper:
            tag = "SUCCESS"

        self.log_box.insert("end", m + "\n", tag)
        self.log_box.see("end")
