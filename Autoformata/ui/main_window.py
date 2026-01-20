import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import threading
import os
from pathlib import Path
from core.excel_handler import OrcamentoEngine
from utils.logger import Logger
from utils.config_manager import ConfigManager
from utils.autocomplete_manager import AutocompleteManager

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class DatabaseEditor(ctk.CTkToplevel):
    """Janela Flutuante para Gerenciar Listas"""
    def __init__(self, parent, manager, callback_refresh):
        super().__init__(parent)
        self.title("📝 Editor de Listas")
        self.geometry("500x600")
        self.manager = manager
        self.callback_refresh = callback_refresh
        self.resizable(False, False)
        
        # Faz a janela ser modal (ficar na frente)
        self.transient(parent)
        self.grab_set()
        
        self._setup_ui()

    def _setup_ui(self):
        # 1. Seleção da Categoria
        top_frame = ctk.CTkFrame(self)
        top_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(top_frame, text="Selecione a Lista para Editar:", font=("Arial", 14, "bold")).pack(pady=5)
        
        # Mapeamento Nome Amigável -> Chave do Banco
        self.cats = {
            "Campi (Campus)": "campus",
            "Setores": "setor",
            "Servidores": "servidor",
            "Elaboradores": "elaborador",
            "Estagiários": "estagiario",
            "Fiscais": "fiscal"
        }
        
        self.combo_cat = ctk.CTkComboBox(top_frame, values=list(self.cats.keys()), command=self._carregar_lista, width=300)
        self.combo_cat.pack(pady=5)
        
        # 2. Área de Rolagem para os Itens
        self.scroll = ctk.CTkScrollableFrame(self, label_text="Itens Salvos")
        self.scroll.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 3. Botão Fechar
        ctk.CTkButton(self, text="Concluir", command=self.destroy, fg_color="gray").pack(pady=10)
        
        # Carrega a primeira lista
        self._carregar_lista(list(self.cats.keys())[0])

    def _carregar_lista(self, cat_friendly):
        # Limpa lista atual
        for widget in self.scroll.winfo_children():
            widget.destroy()
            
        key = self.cats[cat_friendly]
        items = self.manager.get_list(key)
        
        if not items:
            ctk.CTkLabel(self.scroll, text="(Lista Vazia)", text_color="gray").pack(pady=20)
            return

        for item in items:
            row = ctk.CTkFrame(self.scroll, fg_color="transparent")
            row.pack(fill="x", pady=2)
            
            # Botão de Excluir
            btn_del = ctk.CTkButton(
                row, text="🗑️", width=40, fg_color="#C0392B", hover_color="#E74C3C",
                command=lambda k=key, i=item: self._deletar_item(k, i)
            )
            btn_del.pack(side="right", padx=5)
            
            # Texto
            ctk.CTkLabel(row, text=item, anchor="w").pack(side="left", padx=5, fill="x", expand=True)

    def _deletar_item(self, key, item):
        if self.manager.remove_value(key, item):
            self._carregar_lista(self.combo_cat.get()) # Recarrega a tela
            self.callback_refresh() # Atualiza a janela principal em tempo real

class LevelSelector(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.rows_data = [] 
        self.headers_created = False

    def setup_headers(self):
        if self.headers_created: return
        headers = [("Linha", 40), ("Item", 70), ("Cód.", 60), ("Banco", 60), ("Descrição", 350), ("Nível", 200)]
        for i, (txt, w) in enumerate(headers):
            l = ctk.CTkLabel(self, text=txt, font=("Arial", 11, "bold"), anchor="w")
            l.grid(row=0, column=i, padx=5, sticky="w")
        self.headers_created = True

    def clear(self):
        for widget in self.winfo_children():
            if int(widget.grid_info()["row"]) > 0: widget.destroy()
        self.rows_data = []

    def add_row(self, index_excel, item_val, desc_val, cod_val, banco_val, raw_row_data):
        row_idx = len(self.rows_data) + 1
        ctk.CTkLabel(self, text=str(index_excel), width=40, text_color="orange").grid(row=row_idx, column=0)
        ctk.CTkLabel(self, text=str(item_val)[:15], width=70, anchor="w").grid(row=row_idx, column=1)
        ctk.CTkLabel(self, text=str(cod_val)[:10], width=60, anchor="w", text_color="cyan").grid(row=row_idx, column=2)
        ctk.CTkLabel(self, text=str(banco_val)[:10], width=60, anchor="w", text_color="yellow").grid(row=row_idx, column=3)
        
        desc_short = str(desc_val)[:50] + "..." if len(str(desc_val))>50 else str(desc_val)
        ctk.CTkLabel(self, text=desc_short, width=350, anchor="w").grid(row=row_idx, column=4, padx=5)
        
        suggestion = "ITEM"
        item_str = str(item_val).strip()
        if not item_str or item_str == "nan": suggestion = "IGNORAR"
        elif item_str.count('.') == 0 and item_str.isdigit(): suggestion = "N1"
        elif item_str.count('.') == 1: suggestion = "N2"
        elif item_str.count('.') == 2: suggestion = "N3"
        
        seg = ctk.CTkSegmentedButton(self, values=["N1", "N2", "N3", "ITEM", "IGNORAR"], width=200)
        seg.set(suggestion)
        seg.grid(row=row_idx, column=5, padx=5, pady=2)
        
        self.rows_data.append({"raw_data": raw_row_data, "level_widget": seg})

    def get_final_data(self):
        final_list = []
        for row in self.rows_data:
            nivel = row["level_widget"].get()
            if nivel != "IGNORAR":
                entry = row["raw_data"].copy()
                entry["_NIVEL_FORCADO"] = nivel
                final_list.append(entry)
        return final_list

class SisorcApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("🏗️ SISORC v32 - Versão Completa")
        self.geometry("1300x950")
        
        self.config_manager = ConfigManager()
        self.dados_config = self.config_manager.load_profiles()
        self.autocomplete = AutocompleteManager()
        self.sintetico_path = ""
        self.modelo_path = ""
        self.combos_db_refs = {} 
        
        self._setup_ui()
        self.logger = Logger("SISORC")
        self.logger.adicionar_callback(self._log_callback)
        self.carregar_ui_perfil("PADRAO")
        self.atualizar_listas_visuais()

    def _setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # TOPO
        top = ctk.CTkFrame(self)
        top.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        
        ctk.CTkButton(top, text="1. Sintético", width=100, command=self.sel_sintetico).pack(side="left", padx=5)
        self.lbl_sint = ctk.CTkLabel(top, text="...", text_color="gray")
        self.lbl_sint.pack(side="left", padx=5)
        
        ctk.CTkButton(top, text="2. Modelo", width=100, command=self.sel_modelo, fg_color="#555").pack(side="left", padx=5)
        
        ctk.CTkLabel(top, text="| Início:").pack(side="left", padx=5)
        self.ent_line = ctk.CTkEntry(top, width=40)
        self.ent_line.insert(0, "4")
        self.ent_line.pack(side="left")

        ctk.CTkLabel(top, text="Qtd:").pack(side="left", padx=2)
        self.ent_qtd = ctk.CTkEntry(top, width=40)
        self.ent_qtd.insert(0, "500")
        self.ent_qtd.pack(side="left")
        
        ctk.CTkButton(top, text="🔄 3. CARREGAR", command=self.carregar_preview, fg_color="#E67E22").pack(side="left", padx=15)

        # ABAS
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        
        self.tab_visual = self.tabview.add("🔍 Controle Visual")
        self.tab_map = self.tabview.add("⚙️ Mapeamento")
        self.tab_proj = self.tabview.add("📋 Dados da Obra")

        # --- ABA VISUAL ---
        self.table_control = LevelSelector(self.tab_visual)
        self.table_control.pack(fill="both", expand=True)
        self.table_control.setup_headers()

        # --- ABA MAPEAMENTO ---
        self._setup_mapeamento()

        # --- ABA DADOS OBRA ---
        f_scroll = ctk.CTkScrollableFrame(self.tab_proj)
        f_scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Botão de Gerenciar Listas (NOVO)
        f_top_dados = ctk.CTkFrame(f_scroll, fg_color="transparent")
        f_top_dados.pack(fill="x", padx=10, pady=5)
        ctk.CTkButton(f_top_dados, text="📝 Gerenciar Listas Salvas", command=self.abrir_editor_db, width=200, fg_color="#555").pack(side="right")

        ctk.CTkLabel(f_scroll, text="Nome do Arquivo (Sem extensão):", font=("Arial", 12, "bold")).pack(anchor="w", padx=10)
        self.ent_nome_arquivo = ctk.CTkEntry(f_scroll, width=500)
        self.ent_nome_arquivo.pack(anchor="w", padx=10, pady=(0, 10))

        ctk.CTkLabel(f_scroll, text="Descrição/Título (C15):", font=("Arial", 12, "bold")).pack(anchor="w", padx=10)
        self.ent_desc_cabecalho = ctk.CTkEntry(f_scroll, width=500)
        self.ent_desc_cabecalho.pack(anchor="w", padx=10, pady=(0, 10))

        grid_f = ctk.CTkFrame(f_scroll, fg_color="transparent")
        grid_f.pack(fill="x", padx=10)

        self.cbo_campus = self._add_field(grid_f, "Campus (A8):", "campus", 0, 0)
        self.cbo_setor = self._add_field(grid_f, "Setor (A9):", "setor", 1, 0)
        self.cbo_servidor = self._add_field(grid_f, "Servidor (A10):", "servidor", 2, 0)
        self.cbo_elab = self._add_field(grid_f, "Elaborado Por (A13):", "elaborador", 3, 0)
        self.cbo_estag = self._add_field(grid_f, "Estagiário (A14):", "estagiario", 4, 0)

        self.cbo_fiscal = self._add_field(grid_f, "Fiscal (D22):", "fiscal", 0, 1)
        self.ent_data = self._add_input(grid_f, "Data (A18):", 1, 1, "xx/xx/xxxx")
        self.ent_orcafascio = self._add_input(grid_f, "Cód. Orçafascio (E18):", 2, 1)
        self.ent_processo = self._add_input(grid_f, "Num. Processo (E21):", 3, 1)
        
        # Painel Financeiro
        f_fin = ctk.CTkFrame(f_scroll)
        f_fin.pack(fill="x", padx=10, pady=15)
        ctk.CTkLabel(f_fin, text="Configuração Financeira & Layout", font=("Arial", 12, "bold", "underline")).pack(pady=5)
        
        fin_grid = ctk.CTkFrame(f_fin, fg_color="transparent")
        fin_grid.pack(fill="x", padx=5, pady=5)

        ctk.CTkLabel(fin_grid, text="BDI:").grid(row=0, column=0, padx=5, sticky="w")
        self.combo_bdi = ctk.CTkComboBox(fin_grid, width=220, values=[
            "28,82% (SUP - Desc 0,19)",
            "35,18% (PRUMO - Desc 0,0601)",
            "0,00%"
        ])
        self.combo_bdi.grid(row=0, column=1, padx=5, sticky="w")

        ctk.CTkLabel(fin_grid, text="Método de Cálculo:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.combo_metodo_calc = ctk.CTkComboBox(fin_grid, width=250, values=[
            "Cortar Casas (Padrão - Ignora resto)",
            "Arredondar (2 Casas - Matemático)",
            "Exato (Sem tratamento - Excel)"
        ])
        self.combo_metodo_calc.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.combo_metodo_calc.set("Cortar Casas (Padrão - Ignora resto)")

        ctk.CTkLabel(fin_grid, text="Altura Linha (Px):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.ent_altura = ctk.CTkEntry(fin_grid, width=100)
        self.ent_altura.insert(0, "24.75")
        self.ent_altura.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # RODAPÉ
        bot = ctk.CTkFrame(self, height=100)
        bot.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        
        btn_frame = ctk.CTkFrame(bot, fg_color="transparent")
        btn_frame.pack(side="left", fill="both", expand=True)

        self.btn_run = ctk.CTkButton(btn_frame, text="🚀 GERAR ORÇAMENTO", command=self.executar, height=40, fg_color="green", font=("Arial", 14, "bold"))
        self.btn_run.pack(side="top", padx=10, pady=(10, 5), fill="x")

        self.progress = ctk.CTkProgressBar(btn_frame, orientation="horizontal", mode="indeterminate")
        self.progress.pack(side="bottom", padx=10, pady=(0, 10), fill="x")
        self.progress.pack_forget()

        self.log_box = ctk.CTkTextbox(bot, height=80, width=400)
        self.log_box.pack(side="right", padx=10, pady=10)

    def _add_field(self, parent, label, db_key, row, col):
        ctk.CTkLabel(parent, text=label, font=("Arial", 11, "bold")).grid(row=row*2, column=col, sticky="w", padx=5, pady=(5,0))
        cbo = ctk.CTkComboBox(parent, width=250, values=[])
        cbo.set("")
        cbo.grid(row=row*2+1, column=col, sticky="w", padx=5, pady=(0,5))
        self.combos_db_refs[db_key] = cbo
        return cbo

    def _add_input(self, parent, label, row, col, default=""):
        ctk.CTkLabel(parent, text=label, font=("Arial", 11, "bold")).grid(row=row*2, column=col, sticky="w", padx=5, pady=(5,0))
        ent = ctk.CTkEntry(parent, width=250)
        if default: ent.insert(0, default)
        ent.grid(row=row*2+1, column=col, sticky="w", padx=5, pady=(0,5))
        return ent
    
    def atualizar_listas_visuais(self):
        for db_key, widget in self.combos_db_refs.items():
            lista = self.autocomplete.get_list(db_key)
            if lista:
                widget.configure(values=lista)
            else:
                widget.configure(values=["(Digite um novo...)"])

    def abrir_editor_db(self):
        """Abre a janela de edição de listas"""
        DatabaseEditor(self, self.autocomplete, self.atualizar_listas_visuais)

    def _setup_mapeamento(self):
        f = ctk.CTkFrame(self.tab_map)
        f.pack(fill="both", expand=True, padx=20, pady=20)
        self.combos_map = {}
        campos = ["ITEM", "CODIGO", "BANCO", "DESCRICAO", "UNID", "QUANT", "UNIT"]
        for camp in campos:
            r = ctk.CTkFrame(f, fg_color="transparent")
            r.pack(fill="x", pady=2)
            ctk.CTkLabel(r, text=f"{camp}:", width=100, anchor="w").pack(side="left")
            cb = ctk.CTkComboBox(r, values=["..."], width=300)
            cb.pack(side="left")
            self.combos_map[camp] = cb

    def sel_sintetico(self):
        p = filedialog.askopenfilename()
        if p:
            self.sintetico_path = p
            self.lbl_sint.configure(text=Path(p).name, text_color="lime")
            self.ler_colunas()

    def sel_modelo(self):
        p = filedialog.askopenfilename()
        if p: self.modelo_path = p

    def ler_colunas(self):
        try:
            l = int(self.ent_line.get()) - 1
            df = pd.read_excel(self.sintetico_path, header=l, nrows=5)
            cols = [str(c).strip() for c in df.columns if "Unnamed" not in str(c)]
            for k, cb in self.combos_map.items():
                cb.configure(values=cols)
                for c in cols:
                    if k in c.upper(): cb.set(c)
                    if k=="CODIGO" and "COD" in c.upper(): cb.set(c)
                    if k=="BANCO" and ("FONTE" in c.upper() or "REF" in c.upper()): cb.set(c)
                    if k=="DESCRICAO" and "DESC" in c.upper(): cb.set(c)
                    if k=="UNIT" and "VALOR" in c.upper(): cb.set(c)
        except: pass

    def carregar_preview(self):
        if not self.sintetico_path: return messagebox.showwarning("Erro", "Selecione o sintético")
        self.configure(cursor="watch")
        self.update()
        try:
            l = int(self.ent_line.get()) - 1
            df = pd.read_excel(self.sintetico_path, header=l, nrows=int(self.ent_qtd.get()))
            df.columns = [str(c).strip() for c in df.columns]
            c_i, c_d = self.combos_map["ITEM"].get(), self.combos_map["DESCRICAO"].get()
            c_c, c_b = self.combos_map["CODIGO"].get(), self.combos_map["BANCO"].get()
            self.table_control.clear()
            for idx, row in df.iterrows():
                if str(row.get(c_d, 'nan')) == 'nan': continue
                self.table_control.add_row(l+idx+2, row.get(c_i,''), row.get(c_d,''), row.get(c_c,''), row.get(c_b,''), row)
            self.tabview.set("🔍 Controle Visual")
        except Exception as e: messagebox.showerror("Erro", str(e))
        finally: self.configure(cursor="")

    def carregar_ui_perfil(self, nome):
        prof = self.dados_config["perfis"].get("PADRAO", {}).get("input", {})
        for k, v in prof.items():
            if k in self.combos_map: self.combos_map[k].set(v)

    def executar(self):
        if not self.modelo_path: return messagebox.showwarning("Erro", "Falta Modelo")
        d = self.table_control.get_final_data()
        if not d: return messagebox.showwarning("Vazio", "Tabela vazia")
        m = {k: cb.get() for k, cb in self.combos_map.items()}
        bdi_str = self.combo_bdi.get().split('%')[0].replace(',', '.')
        try: bdi = float(bdi_str) / 100
        except: bdi = 0.0
        metodo = self.combo_metodo_calc.get()
        if "Cortar" in metodo: calc_mode = "TRUNC"
        elif "Arredondar" in metodo: calc_mode = "ROUND"
        else: calc_mode = "EXACT"
        try: altura = float(self.ent_altura.get().replace(',', '.'))
        except: altura = 24.75
        info = {
            "nome_arquivo": self.ent_nome_arquivo.get() or "Orcamento",
            "descricao_header": self.ent_desc_cabecalho.get(),
            "campus": self.cbo_campus.get(),
            "setor": self.cbo_setor.get(),
            "servidor": self.cbo_servidor.get(),
            "elaborador": self.cbo_elab.get(),
            "estagiario": self.cbo_estag.get(),
            "fiscal": self.cbo_fiscal.get(),
            "data": self.ent_data.get(),
            "orcafascio": self.ent_orcafascio.get(),
            "processo": self.ent_processo.get(),
            "bdi": bdi,
            "calc_mode": calc_mode,
            "altura_linha": altura
        }
        for key in ["campus", "setor", "servidor", "elaborador", "estagiario", "fiscal"]:
             self.autocomplete.add_value(key, info[key])
        self.atualizar_listas_visuais()
        self.btn_run.configure(state="disabled", text="Processando...")
        self.progress.pack(side="bottom", padx=10, pady=(0, 10), fill="x")
        self.progress.start()
        threading.Thread(target=self._run, args=(d, m, info)).start()

    def _run(self, d, m, p):
        eng = OrcamentoEngine({})
        ok, msg, _ = eng.gerar_excel_final(d, self.modelo_path, m, p)
        self.after(0, lambda: self._finish_run(ok, msg))

    def _finish_run(self, ok, msg):
        self.progress.stop()
        self.progress.pack_forget()
        self.btn_run.configure(state="normal", text="🚀 GERAR ORÇAMENTO")
        if ok:
            messagebox.showinfo("Sucesso", f"Salvo:\n{msg}")
            try: os.startfile(msg)
            except: pass
        else: messagebox.showerror("Erro", msg)

    def _log_callback(self, n, m):
        self.log_box.insert("end", m + "\n")
        self.log_box.see("end")