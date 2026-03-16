import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
from tkinterdnd2 import TkinterDnD, DND_FILES
import os
from pathlib import Path
from tkinter import ttk
import tkinter as tk

from controllers.main_controller import MainController
from ui.components.top_dashboard import TopDashboard
from ui.components.side_panel import SidePanel
from ui.components.config_panel import ConfigPanel

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


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


class LevelSelector(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.rows_data = []

        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Treeview",
                        background="#2B2B2B",
                        foreground="white",
                        rowheight=25,
                        fieldbackground="#2B2B2B",
                        borderwidth=0)
        style.configure("Treeview.Heading",
                        background="#1f538d",
                        foreground="white",
                        font=("Arial", 11, "bold"))
        style.map("Treeview", background=[("selected", "#3498DB")])

        self.tree = ttk.Treeview(self, columns=(
            "L", "Item", "Cod", "Banco", "Desc", "Nivel"), show="headings")

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(column=0, row=0, sticky="nsew")
        vsb.grid(column=1, row=0, sticky="ns")
        hsb.grid(column=0, row=1, sticky="ew")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.setup_headers()

        self.tree.bind("<Double-1>", self._mudar_nivel)
        self.tree.bind("1", lambda e: self._definir_nivel_teclado("N1"))
        self.tree.bind("2", lambda e: self._definir_nivel_teclado("N2"))
        self.tree.bind("3", lambda e: self._definir_nivel_teclado("N3"))
        self.tree.bind("i", lambda e: self._definir_nivel_teclado("ITEM"))
        self.tree.bind("I", lambda e: self._definir_nivel_teclado("ITEM"))
        self.tree.bind("g", lambda e: self._definir_nivel_teclado("IGNORAR"))
        self.tree.bind("G", lambda e: self._definir_nivel_teclado("IGNORAR"))
        self.tree.bind("<space>", self._mudar_nivel)

        ctk.CTkLabel(self, text="Atalhos: Selecione as linhas e aperte 1, 2, 3, I (Item), G (Ignorar) ou Espaço",
                     font=("Arial", 11, "bold"), text_color="gray").grid(column=0, row=2, pady=5)

    def setup_headers(self):
        self.tree.heading("L", text="L")
        self.tree.heading("Item", text="Item")
        self.tree.heading("Cod", text="Cód.")
        self.tree.heading("Banco", text="Banco")
        self.tree.heading("Desc", text="Descrição")
        self.tree.heading("Nivel", text="Nível")

        self.tree.column("L", width=40, anchor="center")
        self.tree.column("Item", width=80, anchor="w")
        self.tree.column("Cod", width=80, anchor="w")
        self.tree.column("Banco", width=80, anchor="w")
        self.tree.column("Desc", width=400, anchor="w")
        self.tree.column("Nivel", width=100, anchor="center")

    def clear(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.rows_data = []

    def add_row(self, index_excel, item_val, desc_val, cod_val, banco_val, raw_row_data, unit_val=None):
        suggestion = "ITEM"
        item_str = str(item_val).strip()

        try:
            val_unit = float(str(unit_val).replace(
                'R$', '').replace('.', '').replace(',', '.'))
            has_value = val_unit > 0
        except:
            has_value = False

        if not item_str or item_str == "nan" or item_str == "None":
            suggestion = "IGNORAR"
        elif not has_value:
            pontos = item_str.count('.')
            if pontos == 0:
                suggestion = "N1"
            elif pontos == 1:
                suggestion = "N2"
            elif pontos >= 2:
                suggestion = "N3"

        row_id = len(self.rows_data)
        self.rows_data.append({"raw_data": raw_row_data, "nivel": suggestion})

        self.tree.insert("", "end", iid=str(row_id), values=(
            index_excel,
            str(item_val)[:15],
            str(cod_val)[:10],
            str(banco_val)[:10],
            str(desc_val),
            suggestion
        ))

    def _mudar_nivel(self, event):
        selecionado = self.tree.selection()
        if not selecionado:
            return

        item_id = selecionado[0]
        valores_atuais = self.tree.item(item_id, "values")
        nivel_atual = valores_atuais[5]

        niveis_possiveis = ["N1", "N2", "N3", "ITEM", "IGNORAR"]

        try:
            proximo_indice = niveis_possiveis.index(nivel_atual) + 1
            if proximo_indice >= len(niveis_possiveis):
                proximo_indice = 0
            novo_nivel = niveis_possiveis[proximo_indice]
        except ValueError:
            novo_nivel = "ITEM"

        novos_valores = list(valores_atuais)
        novos_valores[5] = novo_nivel
        self.tree.item(item_id, values=novos_valores)
        self.rows_data[int(item_id)]["nivel"] = novo_nivel

    def _definir_nivel_teclado(self, novo_nivel):
        selecionados = self.tree.selection()
        if not selecionados:
            return

        for item_id in selecionados:
            valores_atuais = list(self.tree.item(item_id, "values"))
            valores_atuais[5] = novo_nivel
            self.tree.item(item_id, values=valores_atuais)
            self.rows_data[int(item_id)]["nivel"] = novo_nivel

    def get_final_data(self):
        final_list = []
        for r_data in self.rows_data:
            nivel = r_data["nivel"]
            if nivel != "IGNORAR":
                entry = r_data["raw_data"].copy()
                entry["_NIVEL_FORCADO"] = nivel
                final_list.append(entry)
        return final_list


class SisorcApp(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)

        self.title("🏗️ SISORC ULTIMATE - Pro Edition")
        self.geometry("1024x720")

        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self._on_drop)

        self.combos_db_refs = {}
        self.inputs_refs = {}

        self.controller = MainController(self)

        self._setup_ui()
        self.controller.logger.adicionar_callback(self._log_callback)

        self.carregar_ui_perfil("PADRAO")
        self.atualizar_listas_visuais()

        self._carregar_ultima_sessao()
        self.bind('<Control-Return>', lambda e: self.executar())

    def _on_drop(self, event):
        path = event.data
        if path.startswith('{') and path.endswith('}'):
            path = path[1:-1]

        ext = os.path.splitext(path)[1].lower()
        if ext not in ['.xlsx', '.xls', '.xlsm']:
            messagebox.showwarning(
                "Arquivo Inválido", "Apenas arquivos Excel são aceitos!")
            return

        self.lbl_sint.configure(text=os.path.basename(path), text_color="lime")
        self.controller.logger.info(
            f"📂 Arquivo carregado via Drag & Drop: {os.path.basename(path)}")
        self._iniciar_leitura_segura(path)

    def _setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # TOPO: DASHBOARD
        self.top_dashboard = TopDashboard(self, self)
        self.top_dashboard.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=0)

        self.tab_main = self.tabview.add("🏗️ Painel de Orçamento")
        self.tab_config = self.tabview.add("⚙️ Configurações & Mapeamento")

        # NOVO LAYOUT REVOLUCIONÁRIO: PAINEL LATERAL + TABELA GIGANTE
        self.tab_main.grid_columnconfigure(0, weight=0, minsize=320)
        self.tab_main.grid_columnconfigure(1, weight=1)
        self.tab_main.grid_rowconfigure(0, weight=1)

        # --- PAINEL ESQUERDO (DADOS DA OBRA) ---
        self.side_panel = SidePanel(self.tab_main, self)
        self.side_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)

        # --- PAINEL DIREITO (TABELA VISUAL GIGANTE) ---
        self.table_control = LevelSelector(self.tab_main)
        self.table_control.grid(row=0, column=1, sticky="nsew", pady=0)
        self.table_control.setup_headers()

        # ABA 2: CONFIGURAÇÕES E MAPEAMENTO
        self.config_panel = ConfigPanel(self.tab_config, self)
        self.config_panel.pack(fill="both", expand=True, padx=5, pady=5)

        # ==========================================
        # RODAPÉ FIXO (Sempre visível)
        # ==========================================
        bot = ctk.CTkFrame(self, height=100)
        bot.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        bot.grid_columnconfigure(0, weight=1)
        bot.grid_columnconfigure(1, weight=1)

        btn_frame = ctk.CTkFrame(bot, fg_color="transparent")
        btn_frame.grid(row=0, column=0, sticky="nsew", padx=10)

        ctk.CTkLabel(btn_frame, text="Dica: Pressione Ctrl + Enter para gerar",
                     font=("Arial", 10), text_color="gray").pack(side="top", pady=(5, 0))

        self.btn_run = ctk.CTkButton(btn_frame, text="🚀 GERAR ORÇAMENTO", command=self.executar,
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

        self._atualizar_combo_modelos()

    # --- Helpers do Novo Layout Lateral ---
    def _add_side_field(self, parent, label, db_key):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", pady=2, padx=5)
        ctk.CTkLabel(f, text=label, font=("Arial", 11, "bold"),
                     width=85, anchor="w").pack(side="left")
        w = ctk.CTkComboBox(f, width=190, values=[])
        w.set("")
        w.pack(side="right")
        self.combos_db_refs[db_key] = w
        self.inputs_refs[db_key] = w
        return w

    def _add_side_input(self, parent, label, ref_key):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", pady=2, padx=5)
        ctk.CTkLabel(f, text=label, font=("Arial", 11, "bold"),
                     width=85, anchor="w").pack(side="left")
        w = ctk.CTkEntry(f, width=190)
        w.pack(side="right")
        self.inputs_refs[ref_key] = w
        return w

    def _abrir_gerenciador_modelos(self):
        TemplateEditor(self, self.controller.template_manager,
                       self._atualizar_combo_modelos)

    def _atualizar_combo_modelos(self):
        nomes = self.controller.template_manager.get_template_names()
        if nomes:
            self.combo_modelos.configure(values=nomes)
            atual = self.combo_modelos.get()
            if atual not in nomes:
                self.combo_modelos.set(nomes[0])
                self._ao_trocar_modelo(nomes[0])
        else:
            self.combo_modelos.configure(values=["(Nenhum modelo)"])
            self.combo_modelos.set("(Nenhum modelo)")

    def _ao_trocar_modelo(self, nome):
        path = self.controller.template_manager.get_template_path(nome)
        if path and os.path.exists(path):
            self.controller.modelo_path = path
            self.controller.logger.info(f"Modelo definido: {nome}")

    def _alternar_tema(self):
        if self.switch_tema.get() == 1:
            ctk.set_appearance_mode("Dark")
        else:
            ctk.set_appearance_mode("Light")

    def _salvar_sessao_atual(self):
        data = {}
        for key, widget in self.inputs_refs.items():
            try:
                data[key] = widget.get()
            except:
                pass
        self.controller.salvar_sessao_atual(data)

    def _carregar_ultima_sessao(self):
        data = self.controller.carregar_ultima_sessao()
        if not data:
            return
        for key, valor in data.items():
            if key in self.inputs_refs and valor:
                widget = self.inputs_refs[key]
                if isinstance(widget, ctk.CTkEntry):
                    widget.delete(0, 'end')
                    widget.insert(0, valor)
                elif isinstance(widget, ctk.CTkComboBox):
                    widget.set(valor)
        self.controller.logger.info("Estado da última sessão restaurado.")

    def extrair_dados_texto(self):
        texto = self.txt_import.get("0.0", "end").strip()
        if not texto:
            return messagebox.showwarning("Vazio", "Cole o texto do WhatsApp primeiro!")

        dados = self.controller.extrair_dados_texto(texto)

        mapa = {
            "campus": self.cbo_campus, "setor": self.cbo_setor,
            "descricao_header": self.ent_desc_cabecalho, "servidor": self.cbo_servidor,
            "fiscal": self.cbo_fiscal, "elaborador": self.cbo_elab,
            "estagiario": self.cbo_estag, "processo": self.ent_processo,
            "orcafascio": self.ent_orcafascio, "empenho": self.ent_empenho,
            "num_orcamento": self.ent_num_orc
        }

        count = 0
        for key, valor in dados.items():
            if key in mapa and valor:
                widget = mapa[key]
                if isinstance(widget, ctk.CTkEntry):
                    widget.delete(0, 'end')
                    widget.insert(0, valor)
                elif isinstance(widget, ctk.CTkComboBox):
                    widget.set(valor)
                count += 1

        if "descricao_header" in dados and not self.ent_nome_arquivo.get():
            nome_seguro = "".join(
                x for x in dados["descricao_header"][:40] if x.isalnum() or x in " -_")
            self.ent_nome_arquivo.insert(0, nome_seguro)

        self.txt_import.delete("0.0", "end")
        self.controller.logger.info(
            f"Importação Inteligente V2: {count} campos extraídos e normalizados!")
        messagebox.showinfo("Sucesso", f"{count} dados foram extraídos!")

    def _calcular_prazo_auto(self, event=None):
        try:
            val_txt = self.ent_valor_sim.get().replace(
                'R$', '').replace('.', '').replace(',', '.').strip()
            if not val_txt:
                return
            valor = float(val_txt)
            prazo = ""
            if valor <= 50000:
                prazo = "30 DIAS"
            elif valor <= 100000:
                prazo = "60 DIAS"
            elif valor <= 150000:
                prazo = "90 DIAS"
            else:
                prazo = "A DEFINIR (ACORDO)"
            self.ent_prazo.delete(0, 'end')
            self.ent_prazo.insert(0, prazo)
            self.controller.logger.info(
                f"Prazo calculado automaticamente para R$ {valor:,.2f}: {prazo}")
        except:
            pass

    def atualizar_listas_visuais(self):
        for db_key, widget in self.combos_db_refs.items():
            lista = self.controller.autocomplete.get_list(db_key)
            if lista:
                widget.configure(values=lista)
            else:
                widget.configure(values=["(Digite um novo...)"])

    def abrir_editor_db(self):
        DatabaseEditor(self, self.controller.autocomplete,
                       self.atualizar_listas_visuais)

    def sel_sintetico(self):
        p = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
        if p:
            self.lbl_sint.configure(text=Path(p).name, text_color="lime")
            self.controller.logger.info(
                f"📂 Arquivo selecionado manualmente: {os.path.basename(p)}")
            self._iniciar_leitura_segura(p)

    def limpar_dados_sessao(self):
        for key, widget in self.inputs_refs.items():
            if isinstance(widget, ctk.CTkEntry):
                widget.delete(0, 'end')
            elif isinstance(widget, ctk.CTkComboBox):
                widget.set("")

        self.lbl_sint.configure(text="Nenhum arquivo", text_color="gray")
        self.txt_import.delete("0.0", "end")
        self.table_control.clear()

        for k, cb in self.combos_map.items():
            cb.set("...")

        self.controller.limpar_dados_sessao()
        self.controller.logger.info("🧹 Sessão limpa para novo orçamento.")
        self.lbl_status.configure(text="Sessão limpa.", text_color="lime")

    def _iniciar_leitura_segura(self, path):
        self.limpar_dados_sessao()

        self.configure(cursor="watch")
        self.lbl_status.configure(
            text="Limpando corrupções do arquivo SIPAC/SEI...", text_color="orange")

        self.progress.pack(side="bottom", padx=10, pady=(0, 5), fill="x")
        self.progress.start()

        self.controller.iniciar_leitura_segura(path)

    def _on_limpar_planilha_sucesso(self, path_limpo):
        self.progress.stop()
        self.progress.pack_forget()
        self.configure(cursor="")
        self.lbl_status.configure(
            text="Arquivo pronto e limpo!", text_color="gray")
        self.ler_colunas()
        self.carregar_preview()

    def _on_limpar_planilha_erro(self, msg):
        self.progress.stop()
        self.progress.pack_forget()
        self.configure(cursor="")
        messagebox.showerror(
            "Erro Crítico", f"Falha ao limpar o arquivo Excel:\n{msg}")
        self.lbl_status.configure(
            text="Erro ao limpar arquivo.", text_color="red")

    def ler_colunas(self):
        try:
            l = int(self.ent_line.get()) - 1
            cols = self.controller.ler_colunas(l)
            for k, cb in self.combos_map.items():
                cb.configure(values=cols)
                for c in cols:
                    if k in c.upper():
                        cb.set(c)
                    if k == "CODIGO" and "COD" in c.upper():
                        cb.set(c)
                    if k == "BANCO" and ("FONTE" in c.upper() or "REF" in c.upper()):
                        cb.set(c)
                    if k == "DESCRICAO" and "DESC" in c.upper():
                        cb.set(c)
                    if k == "UNIT" and "VALOR" in c.upper():
                        cb.set(c)
        except:
            pass

    def carregar_preview(self):
        self.configure(cursor="watch")
        self.lbl_status.configure(
            text="Carregando tabela...", text_color="orange")

        self.progress.pack(side="bottom", padx=10, pady=(0, 5), fill="x")
        self.progress.start()

        self.ler_colunas()
        l = int(self.ent_line.get()) - 1

        c_i = self.combos_map["ITEM"].get()
        c_d = self.combos_map["DESCRICAO"].get()
        c_c = self.combos_map["CODIGO"].get()
        c_b = self.combos_map["BANCO"].get()
        c_u = self.combos_map["UNIT"].get()

        self.table_control.clear()

        self.controller.carregar_preview(l, c_i, c_d, c_c, c_b, c_u)

    def _on_carregar_preview_sucesso(self, dados_linhas):
        self.progress.stop()
        self.progress.pack_forget()
        self.configure(cursor="")

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
        messagebox.showerror("Erro de Leitura", msg)
        self.lbl_status.configure(
            text="Erro ao ler ficheiro.", text_color="red")

    def carregar_ui_perfil(self, nome):
        prof = self.controller.dados_config["perfis"].get(
            "PADRAO", {}).get("input", {})
        for k, v in prof.items():
            if k in self.combos_map:
                self.combos_map[k].set(v)

    def _obter_valor_seguro(self, entry_widget, is_date=False, custom_placeholder=None):
        valor = entry_widget.get().strip().upper()
        if not valor:
            if custom_placeholder:
                return custom_placeholder
            return "xx/xx/xxxx" if is_date else "xxxxxxxxxx"
        return valor

    def executar(self):
        if not self.controller.modelo_path:
            return messagebox.showwarning("Erro", "Nenhum modelo selecionado! Use o botão de engrenagem para adicionar um.")
        d = self.table_control.get_final_data()
        if not d:
            return messagebox.showwarning("Vazio", "Tabela vazia")
        m = {k: cb.get() for k, cb in self.combos_map.items()}
        bdi_str = self.combo_bdi.get().split('%')[0].replace(',', '.')
        try:
            bdi = float(bdi_str) / 100
        except:
            bdi = 0.0
        metodo = self.combo_metodo_calc.get()
        if "Cortar" in metodo:
            calc_mode = "TRUNC"
        elif "Arredondar" in metodo:
            calc_mode = "ROUND"
        else:
            calc_mode = "EXACT"
        try:
            altura = float(self.ent_altura.get().replace(',', '.'))
        except:
            altura = 24.75

        info = {
            "nome_arquivo": self.ent_nome_arquivo.get() or "Orcamento",
            "descricao_header": self.ent_desc_cabecalho.get().upper(),
            "campus": self.cbo_campus.get().upper(),
            "setor": self.cbo_setor.get().upper(),
            "servidor": self.cbo_servidor.get().upper(),
            "elaborador": self.cbo_elab.get().upper(),
            "estagiario": self.cbo_estag.get().upper(),
            "fiscal": self.cbo_fiscal.get().upper(),
            "data": self._obter_valor_seguro(self.ent_data, is_date=True),
            "orcafascio": self._obter_valor_seguro(self.ent_orcafascio),
            "processo": self._obter_valor_seguro(self.ent_processo),
            "num_orcamento": self._obter_valor_seguro(self.ent_num_orc, custom_placeholder="XX"),
            "empenho": self._obter_valor_seguro(self.ent_empenho),
            "data_emissao": self._obter_valor_seguro(self.ent_data_emissao, is_date=True),
            "data_inicio": self._obter_valor_seguro(self.ent_data_inicio, is_date=True),
            "prazo": self._obter_valor_seguro(self.ent_prazo),
            "bdi": bdi,
            "calc_mode": calc_mode,
            "altura_linha": altura,
            "gerar_pdf": self.chk_pdf.get()
        }

        for key in ["campus", "setor", "servidor", "elaborador", "estagiario", "fiscal"]:
            self.controller.autocomplete.add_value(key, info[key])

        self.atualizar_listas_visuais()
        self._salvar_sessao_atual()

        self.btn_run.configure(state="disabled", text="Processando...")
        self.lbl_status.configure(
            text="Iniciando motor de cálculo...", text_color="orange")

        self.progress.configure(mode="determinate")
        self.progress.set(0)
        self.progress.pack(side="bottom", padx=10, pady=(0, 5), fill="x")

        self.controller.gerar_orcamento(
            d, m, info, self.controller.modelo_path)

    def _on_gerar_orcamento_progresso(self, pct):
        self.progress.set(pct / 100.0)
        self.lbl_status.configure(text=f"Processando: {pct}%")

    def _on_gerar_orcamento_sucesso(self, msg, duration, info_data, raw_data, pdf_msg):
        self.progress.pack_forget()
        self.progress.configure(mode="indeterminate")
        self.btn_run.configure(state="normal", text="🚀 GERAR ORÇAMENTO")

        self.lbl_status.configure(
            text="Concluído com sucesso!", text_color="lime")
        messagebox.showinfo(
            "Sucesso", f"Salvo com sucesso em:\n{msg}{pdf_msg}")
        try:
            os.startfile(msg)
        except:
            pass

    def _on_gerar_orcamento_erro(self, msg):
        self.progress.pack_forget()
        self.progress.configure(mode="indeterminate")
        self.btn_run.configure(state="normal", text="🚀 GERAR ORÇAMENTO")

        self.lbl_status.configure(
            text="Erro no processamento.", text_color="red")
        if "Permission denied" in msg or "aberto" in msg:
            messagebox.showerror(
                "Arquivo Aberto", f"O arquivo parece estar aberto no Excel.\n\nPor favor, feche o arquivo '{msg}' e tente novamente.")
        else:
            messagebox.showerror("Erro Fatal", msg)

    def _log_callback(self, n, m):
        msg_upper = m.upper()
        tag = "INFO"
        if "ERRO" in msg_upper or "❌" in m:
            tag = "ERROR"
        elif "SUCESSO" in msg_upper or "✅" in m or "CONCLUÍDO" in msg_upper:
            tag = "SUCCESS"

        self.log_box.insert("end", m + "\n", tag)
        self.log_box.see("end")
