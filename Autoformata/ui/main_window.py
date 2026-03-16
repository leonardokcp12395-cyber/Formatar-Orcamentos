import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
# Importações para Drag & Drop e Caminhos
from tkinterdnd2 import TkinterDnD, DND_FILES 
import pandas as pd
import threading
import os
import shutil  # CORREÇÃO: Importação crucial para copiar o arquivo com segurança!
import json
import time
from datetime import datetime
from pathlib import Path
from tkinter import ttk
import tkinter as tk

import win32com.client
import pythoncom

# Importações do Core
from core.excel_handler import OrcamentoEngine
from core.database import DatabaseManager
from core.paths import get_app_dir

# Importações de Utils
from utils.logger import Logger
from utils.config_manager import ConfigManager
from utils.autocomplete_manager import AutocompleteManager
from utils.smart_parser import SmartParser
from utils.pdf_exporter import PDFExporter
from utils.template_manager import TemplateManager

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
        ctk.CTkLabel(f_new, text="Adicionar Novo Modelo", font=("Arial", 12, "bold")).pack(pady=5)
        
        f_btn = ctk.CTkFrame(f_new, fg_color="transparent")
        f_btn.pack(fill="x")
        
        self.btn_file = ctk.CTkButton(f_btn, text="📂 Selecionar Excel (.xlsx)", command=self._pick_file)
        self.btn_file.pack(side="left", padx=5, fill="x", expand=True)
        
        self.lbl_file = ctk.CTkLabel(f_new, text="Nenhum arquivo selecionado", text_color="gray", font=("Arial", 10))
        self.lbl_file.pack(pady=2)

        f_inputs = ctk.CTkFrame(f_new, fg_color="transparent")
        f_inputs.pack(fill="x", pady=5)
        
        self.ent_name = ctk.CTkEntry(f_inputs, placeholder_text="Nome (ex: PADRAO 2026)")
        self.ent_name.pack(side="left", padx=5, fill="x", expand=True)
        
        self.ent_start = ctk.CTkEntry(f_inputs, width=80, placeholder_text="Linha Início")
        self.ent_start.insert(0, "25")
        self.ent_start.pack(side="right", padx=5)

        ctk.CTkButton(f_new, text="💾 Salvar Modelo", command=self._save, fg_color="green").pack(fill="x", padx=5, pady=5)

        ctk.CTkLabel(self, text="Modelos Salvos:", font=("Arial", 12, "bold")).pack(pady=(10,5))
        self.scroll = ctk.CTkScrollableFrame(self)
        self.scroll.pack(fill="both", expand=True, padx=10, pady=10)
        self._load_list()

    def _pick_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p:
            self.path_temp = p
            self.lbl_file.configure(text=os.path.basename(p), text_color="lime")

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
            self.lbl_file.configure(text="Nenhum arquivo selecionado", text_color="gray")
            self._load_list()
            self.callback_refresh()
        else:
            messagebox.showerror("Erro", msg)

    def _load_list(self):
        for w in self.scroll.winfo_children(): w.destroy()
        names = self.manager.get_template_names()
        for n in names:
            f = ctk.CTkFrame(self.scroll, fg_color="transparent")
            f.pack(fill="x", pady=2)
            ctk.CTkLabel(f, text=n).pack(side="left", padx=5)
            ctk.CTkButton(f, text="🗑️", width=30, fg_color="red", command=lambda x=n: self._del(x)).pack(side="right")

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
        ctk.CTkLabel(top_frame, text="Selecione a Lista para Editar:", font=("Arial", 14, "bold")).pack(pady=5)
        
        self.cats = {
            "Campi (Campus)": "campus", "Setores": "setor", "Servidores": "servidor",
            "Elaboradores": "elaborador", "Estagiários": "estagiario", "Fiscais": "fiscal"
        }
        self.combo_cat = ctk.CTkComboBox(top_frame, values=list(self.cats.keys()), command=self._carregar_lista, width=300)
        self.combo_cat.pack(pady=5)
        
        self.scroll = ctk.CTkScrollableFrame(self, label_text="Itens Salvos")
        self.scroll.pack(fill="both", expand=True, padx=10, pady=5)
        ctk.CTkButton(self, text="Concluir", command=self.destroy, fg_color="gray").pack(pady=10)
        self._carregar_lista(list(self.cats.keys())[0])

    def _carregar_lista(self, cat_friendly):
        for widget in self.scroll.winfo_children(): widget.destroy()
        key = self.cats[cat_friendly]
        items = self.manager.get_list(key)
        if not items:
            ctk.CTkLabel(self.scroll, text="(Lista Vazia)", text_color="gray").pack(pady=20)
            return
        for item in items:
            row = ctk.CTkFrame(self.scroll, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkButton(row, text="🗑️", width=40, fg_color="#C0392B", hover_color="#E74C3C",
                command=lambda k=key, i=item: self._deletar_item(k, i)).pack(side="right", padx=5)
            ctk.CTkLabel(row, text=item, anchor="w").pack(side="left", padx=5, fill="x", expand=True)

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

        self.tree = ttk.Treeview(self, columns=("L", "Item", "Cod", "Banco", "Desc", "Nivel"), show="headings")
        
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
            val_unit = float(str(unit_val).replace('R$', '').replace('.', '').replace(',', '.'))
            has_value = val_unit > 0
        except:
            has_value = False

        if not item_str or item_str == "nan" or item_str == "None": suggestion = "IGNORAR"
        elif not has_value:
            pontos = item_str.count('.')
            if pontos == 0: suggestion = "N1"
            elif pontos == 1: suggestion = "N2"
            elif pontos >= 2: suggestion = "N3"
        
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
        if not selecionado: return
        
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
        if not selecionados: return
        
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

        self.logger = Logger("SISORC")
        self.config_manager = ConfigManager()
        self.dados_config = self.config_manager.load_profiles()
        self.autocomplete = AutocompleteManager()
        self.template_manager = TemplateManager()
        
        db_path = get_app_dir() / 'sisorc_history.db'
        db_config = {'database': {'nome_arquivo': str(db_path)}}
        self.db_manager = DatabaseManager(db_config)
        
        # Variáveis seguras para a máquina de lavar de planilhas
        self.sintetico_original_path = ""
        self.sintetico_limpo_path = ""
        
        self.modelo_path = ""
        self.combos_db_refs = {} 
        self.inputs_refs = {} 
        
        self._setup_ui()
        self.logger.adicionar_callback(self._log_callback)
        
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
            messagebox.showwarning("Arquivo Inválido", "Apenas arquivos Excel são aceitos!")
            return

        self.sintetico_original_path = path
        self.lbl_sint.configure(text=os.path.basename(path), text_color="lime")
        self.logger.info(f"📂 Arquivo carregado via Drag & Drop: {os.path.basename(path)}")
        self._iniciar_leitura_segura()

    def _setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # TOPO: DASHBOARD
        top_dashboard = ctk.CTkFrame(self, fg_color="#1E1E1E", border_width=1, border_color="#3498DB")
        top_dashboard.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        top_dashboard.grid_columnconfigure(0, weight=1)
        top_dashboard.grid_columnconfigure(1, weight=1)

        f_sint = ctk.CTkFrame(top_dashboard, fg_color="transparent")
        f_sint.grid(row=0, column=0, padx=15, pady=10, sticky="nsew")
        ctk.CTkLabel(f_sint, text="1. Arraste ou Selecione o Sintético", font=("Arial", 13, "bold"), text_color="#3498DB").pack(anchor="w")
        
        f_sint_inner = ctk.CTkFrame(f_sint, fg_color="transparent")
        f_sint_inner.pack(fill="x", pady=5)
        ctk.CTkButton(f_sint_inner, text="📂 Selecionar", width=100, command=self.sel_sintetico).pack(side="left", padx=(0,10))
        self.lbl_sint = ctk.CTkLabel(f_sint_inner, text="Nenhum arquivo", text_color="gray")
        self.lbl_sint.pack(side="left")
        
        ctk.CTkButton(f_sint, text="🔄 Carregar Tabela Visual", command=self.carregar_preview, fg_color="#E67E22").pack(anchor="w", pady=5)

        f_wpp = ctk.CTkFrame(top_dashboard, fg_color="transparent")
        f_wpp.grid(row=0, column=1, padx=15, pady=10, sticky="nsew")
        ctk.CTkLabel(f_wpp, text="2. Importação Inteligente (WhatsApp)", font=("Arial", 13, "bold"), text_color="#8E44AD").pack(anchor="w")
        
        f_wpp_inner = ctk.CTkFrame(f_wpp, fg_color="transparent")
        f_wpp_inner.pack(fill="x", pady=5)
        self.txt_import = ctk.CTkTextbox(f_wpp_inner, height=55)
        self.txt_import.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(f_wpp_inner, text="🪄 Extrair", command=self.extrair_dados_texto, fg_color="#8E44AD", width=80, height=55).pack(side="right")

        self.switch_tema = ctk.CTkSwitch(top_dashboard, text="Modo Escuro", command=self._alternar_tema)
        self.switch_tema.select()
        self.switch_tema.place(relx=0.98, rely=0.1, anchor="ne")

        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=0)
        
        self.tab_main = self.tabview.add("🏗️ Painel de Orçamento")
        self.tab_config = self.tabview.add("⚙️ Configurações & Mapeamento")

        # -----------------------------------------------------------------
        # NOVO LAYOUT REVOLUCIONÁRIO: PAINEL LATERAL + TABELA GIGANTE
        # -----------------------------------------------------------------
        self.tab_main.grid_columnconfigure(0, weight=0, minsize=320) # Lado Esquerdo Fixo
        self.tab_main.grid_columnconfigure(1, weight=1) # Lado Direito Ocupa Todo o Resto
        self.tab_main.grid_rowconfigure(0, weight=1) # Ocupa a altura inteira

        # --- PAINEL ESQUERDO (DADOS DA OBRA) ---
        f_dados = ctk.CTkScrollableFrame(self.tab_main, fg_color="transparent", width=310)
        f_dados.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)
        
        f_top_dados = ctk.CTkFrame(f_dados, fg_color="transparent")
        f_top_dados.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(f_top_dados, text="Dados da Obra", font=("Arial", 14, "bold")).pack(side="left")
        ctk.CTkButton(f_top_dados, text="📝 Listas", command=self.abrir_editor_db, width=60, height=24, fg_color="#444").pack(side="right")

        ctk.CTkLabel(f_dados, text="Nome (Arquivo):", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_nome_arquivo = ctk.CTkEntry(f_dados, width=290)
        self.ent_nome_arquivo.pack(anchor="w", pady=(0,5))
        self.inputs_refs["nome_arquivo"] = self.ent_nome_arquivo

        ctk.CTkLabel(f_dados, text="Descrição (C15):", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_desc_cabecalho = ctk.CTkEntry(f_dados, width=290)
        self.ent_desc_cabecalho.pack(anchor="w", pady=(0,10))
        self.inputs_refs["descricao_header"] = self.ent_desc_cabecalho

        grid_f = ctk.CTkFrame(f_dados, fg_color="#212121")
        grid_f.pack(fill="x", pady=5)
        
        self.cbo_campus = self._add_side_field(grid_f, "Campus:", "campus")
        self.cbo_setor = self._add_side_field(grid_f, "Setor:", "setor")
        self.cbo_fiscal = self._add_side_field(grid_f, "Fiscal:", "fiscal")
        self.cbo_servidor = self._add_side_field(grid_f, "Servidor:", "servidor")
        self.cbo_elab = self._add_side_field(grid_f, "Elaborador:", "elaborador")
        self.cbo_estag = self._add_side_field(grid_f, "Estagiário:", "estagiario")

        self.ent_data = self._add_side_input(grid_f, "Data Elab.:", "data")
        self.ent_orcafascio = self._add_side_input(grid_f, "Orçafascio:", "orcafascio")
        self.ent_processo = self._add_side_input(grid_f, "Processo:", "processo")
        self.ent_num_orc = self._add_side_input(grid_f, "Nº Orc:", "num_orcamento")
        self.ent_empenho = self._add_side_input(grid_f, "Empenho:", "empenho")
        self.ent_data_emissao = self._add_side_input(grid_f, "Emissão:", "data_emissao") 
        self.ent_data_inicio = self._add_side_input(grid_f, "Início:", "data_inicio")
        
        f_p_inner = ctk.CTkFrame(f_dados, fg_color="transparent")
        f_p_inner.pack(fill="x", pady=10)
        ctk.CTkLabel(f_p_inner, text="Valor Sim. (R$):", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_valor_sim = ctk.CTkEntry(f_p_inner, width=290, placeholder_text="0,00")
        self.ent_valor_sim.pack(anchor="w", pady=(0,5))
        self.ent_valor_sim.bind("<FocusOut>", self._calcular_prazo_auto)
        self.ent_valor_sim.bind("<Return>", self._calcular_prazo_auto)
        
        ctk.CTkLabel(f_p_inner, text="Prazo Final:", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_prazo = ctk.CTkEntry(f_p_inner, width=290)
        self.ent_prazo.pack(anchor="w")
        self.inputs_refs["prazo"] = self.ent_prazo

        # --- PAINEL DIREITO (TABELA VISUAL GIGANTE) ---
        self.table_control = LevelSelector(self.tab_main)
        self.table_control.grid(row=0, column=1, sticky="nsew", pady=0)
        self.table_control.setup_headers()

        # ------------------------------------------
        # ABA 2: CONFIGURAÇÕES E MAPEAMENTO
        # ------------------------------------------
        f_conf_scroll = ctk.CTkScrollableFrame(self.tab_config)
        f_conf_scroll.pack(fill="both", expand=True, padx=5, pady=5)

        f_leitura = ctk.CTkFrame(f_conf_scroll)
        f_leitura.pack(fill="x", pady=5, padx=10)
        ctk.CTkLabel(f_leitura, text="1. Modelo e Sintético", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)
        
        f_l_inner = ctk.CTkFrame(f_leitura, fg_color="transparent")
        f_l_inner.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(f_l_inner, text="Modelo Excel:").pack(side="left")
        self.combo_modelos = ctk.CTkComboBox(f_l_inner, width=250, command=self._ao_trocar_modelo)
        self.combo_modelos.pack(side="left", padx=5)
        ctk.CTkButton(f_l_inner, text="⚙️ Gerenciar Modelos", width=140, fg_color="#555", command=self._abrir_gerenciador_modelos).pack(side="left", padx=5)
        
        ctk.CTkLabel(f_l_inner, text="Ler Sintético a partir da Linha:").pack(side="left", padx=(20, 5))
        self.ent_line = ctk.CTkEntry(f_l_inner, width=50)
        self.ent_line.insert(0, "4")
        self.ent_line.pack(side="left")

        f_fin = ctk.CTkFrame(f_conf_scroll)
        f_fin.pack(fill="x", pady=10, padx=10)
        ctk.CTkLabel(f_fin, text="2. Financeiro & Saída", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)
        
        fin_grid = ctk.CTkFrame(f_fin, fg_color="transparent")
        fin_grid.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(fin_grid, text="BDI:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.combo_bdi = ctk.CTkComboBox(fin_grid, width=250, values=["28,82% (SUP - Desc 0,19)", "35,18% (PRUMO - Desc 0,0601)", "0,00%"])
        self.combo_bdi.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.inputs_refs["bdi_combo"] = self.combo_bdi

        ctk.CTkLabel(fin_grid, text="Método de Cálculo:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.combo_metodo_calc = ctk.CTkComboBox(fin_grid, width=250, values=["Cortar Casas (Padrão - Ignora resto)", "Arredondar (2 Casas - Matemático)", "Exato (Sem tratamento - Excel)"])
        self.combo_metodo_calc.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.combo_metodo_calc.set("Cortar Casas (Padrão - Ignora resto)")

        ctk.CTkLabel(fin_grid, text="Altura Linha (Px):").grid(row=0, column=2, padx=(20, 5), pady=5, sticky="w")
        self.ent_altura = ctk.CTkEntry(fin_grid, width=100)
        self.ent_altura.insert(0, "24.75")
        self.ent_altura.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        self.chk_pdf = ctk.CTkCheckBox(fin_grid, text="Gerar PDF automaticamente após Excel", text_color="lime")
        self.chk_pdf.grid(row=1, column=2, columnspan=2, padx=(20, 5), pady=5, sticky="w")

        f_map = ctk.CTkFrame(f_conf_scroll)
        f_map.pack(fill="x", pady=10, padx=10)
        ctk.CTkLabel(f_map, text="3. Mapeamento de Colunas", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)
        
        f_map_inner = ctk.CTkFrame(f_map, fg_color="transparent")
        f_map_inner.pack(fill="x", padx=10, pady=5)
        
        self.combos_map = {}
        campos = ["ITEM", "CODIGO", "BANCO", "DESCRICAO", "UNID", "QUANT", "UNIT"]
        for i, camp in enumerate(campos):
            r = i // 2
            c = (i % 2) * 2
            ctk.CTkLabel(f_map_inner, text=f"{camp}:", width=80, anchor="e").grid(row=r, column=c, padx=5, pady=5)
            cb = ctk.CTkComboBox(f_map_inner, values=["..."], width=200)
            cb.grid(row=r, column=c+1, padx=5, pady=5)
            self.combos_map[camp] = cb

        # ==========================================
        # 3. RODAPÉ FIXO (Sempre visível)
        # ==========================================
        bot = ctk.CTkFrame(self, height=100)
        bot.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        bot.grid_columnconfigure(0, weight=1)
        bot.grid_columnconfigure(1, weight=1)
        
        btn_frame = ctk.CTkFrame(bot, fg_color="transparent")
        btn_frame.grid(row=0, column=0, sticky="nsew", padx=10)

        ctk.CTkLabel(btn_frame, text="Dica: Pressione Ctrl + Enter para gerar", font=("Arial", 10), text_color="gray").pack(side="top", pady=(5,0))

        self.btn_run = ctk.CTkButton(btn_frame, text="🚀 GERAR ORÇAMENTO", command=self.executar, height=45, fg_color="green", font=("Arial", 16, "bold"))
        self.btn_run.pack(side="top", pady=5, fill="x")

        status_frame = ctk.CTkFrame(btn_frame, fg_color="transparent")
        status_frame.pack(side="top", fill="x")
        self.lbl_status = ctk.CTkLabel(status_frame, text="Pronto para uso.", font=("Arial", 11), text_color="gray")
        self.lbl_status.pack(side="left")

        self.progress = ctk.CTkProgressBar(status_frame, orientation="horizontal", mode="indeterminate")

        self.log_box = ctk.CTkTextbox(bot, height=90)
        self.log_box.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

        self._atualizar_combo_modelos()

    # --- Helpers do Novo Layout Lateral ---
    def _add_side_field(self, parent, label, db_key):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", pady=2, padx=5)
        ctk.CTkLabel(f, text=label, font=("Arial", 11, "bold"), width=85, anchor="w").pack(side="left")
        w = ctk.CTkComboBox(f, width=190, values=[])
        w.set("")
        w.pack(side="right")
        self.combos_db_refs[db_key] = w
        self.inputs_refs[db_key] = w
        return w

    def _add_side_input(self, parent, label, ref_key):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", pady=2, padx=5)
        ctk.CTkLabel(f, text=label, font=("Arial", 11, "bold"), width=85, anchor="w").pack(side="left")
        w = ctk.CTkEntry(f, width=190)
        w.pack(side="right")
        self.inputs_refs[ref_key] = w
        return w

    def _abrir_gerenciador_modelos(self):
        TemplateEditor(self, self.template_manager, self._atualizar_combo_modelos)

    def _atualizar_combo_modelos(self):
        nomes = self.template_manager.get_template_names()
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
        path = self.template_manager.get_template_path(nome)
        if path and os.path.exists(path):
            self.modelo_path = path
            self.logger.info(f"Modelo definido: {nome}")

    def _alternar_tema(self):
        if self.switch_tema.get() == 1: ctk.set_appearance_mode("Dark")
        else: ctk.set_appearance_mode("Light")

    def _get_session_path(self):
        return get_app_dir() / "config" / "last_session.json"

    def _salvar_sessao_atual(self):
        data = {}
        for key, widget in self.inputs_refs.items():
            try: data[key] = widget.get()
            except: pass
        try:
            sess_path = self._get_session_path()
            sess_path.parent.mkdir(parents=True, exist_ok=True)
            with open(sess_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False)
        except Exception as e: print(f"Erro ao salvar sessão: {e}")

    def _carregar_ultima_sessao(self):
        path = self._get_session_path()
        if not path.exists(): return
        try:
            with open(path, 'r', encoding='utf-8') as f: data = json.load(f)
            for key, valor in data.items():
                if key in self.inputs_refs and valor:
                    widget = self.inputs_refs[key]
                    if isinstance(widget, ctk.CTkEntry):
                        widget.delete(0, 'end')
                        widget.insert(0, valor)
                    elif isinstance(widget, ctk.CTkComboBox):
                        widget.set(valor)
            self.logger.info("Estado da última sessão restaurado.")
        except Exception as e: print(f"Erro ao carregar sessão: {e}")

    def extrair_dados_texto(self):
        texto = self.txt_import.get("0.0", "end").strip()
        if not texto: return messagebox.showwarning("Vazio", "Cole o texto do WhatsApp primeiro!")
        
        dados = SmartParser.parse_whatsapp_text(texto, self.autocomplete)
        
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
             nome_seguro = "".join(x for x in dados["descricao_header"][:40] if x.isalnum() or x in " -_")
             self.ent_nome_arquivo.insert(0, nome_seguro)

        self.txt_import.delete("0.0", "end")
        self.logger.info(f"Importação Inteligente V2: {count} campos extraídos e normalizados!")
        messagebox.showinfo("Sucesso", f"{count} dados foram extraídos!")

    def _calcular_prazo_auto(self, event=None):
        try:
            val_txt = self.ent_valor_sim.get().replace('R$','').replace('.','').replace(',','.').strip()
            if not val_txt: return
            valor = float(val_txt)
            prazo = ""
            if valor <= 50000: prazo = "30 DIAS"
            elif valor <= 100000: prazo = "60 DIAS"
            elif valor <= 150000: prazo = "90 DIAS"
            else: prazo = "A DEFINIR (ACORDO)"
            self.ent_prazo.delete(0, 'end')
            self.ent_prazo.insert(0, prazo)
            self.logger.info(f"Prazo calculado automaticamente para R$ {valor:,.2f}: {prazo}")
        except: pass 
    
    def atualizar_listas_visuais(self):
        for db_key, widget in self.combos_db_refs.items():
            lista = self.autocomplete.get_list(db_key)
            if lista:
                widget.configure(values=lista)
            else:
                widget.configure(values=["(Digite um novo...)"])

    def abrir_editor_db(self):
        DatabaseEditor(self, self.autocomplete, self.atualizar_listas_visuais)

    def sel_sintetico(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
        if p:
            self.sintetico_original_path = p
            self.lbl_sint.configure(text=Path(p).name, text_color="lime")
            self.logger.info(f"📂 Arquivo selecionado manualmente: {os.path.basename(p)}")
            self._iniciar_leitura_segura()

    # ==========================================================
    # MÁQUINA DE LAVAR DE PLANILHAS (Win32COM)
    # ==========================================================
    def _iniciar_leitura_segura(self):
        """Orquestra a limpeza do arquivo SIPAC/SEI antes de tentar ler com Pandas"""
        self.configure(cursor="watch")
        self.lbl_status.configure(text="Limpando corrupções do arquivo SIPAC/SEI...", text_color="orange")
        self.update()
        
        try:
            ok, path_limpo = self._limpar_planilha_sipac(self.sintetico_original_path)
            if ok:
                self.sintetico_limpo_path = path_limpo
                self.lbl_status.configure(text="Arquivo pronto e limpo!", text_color="gray")
                self.ler_colunas()
                self.carregar_preview()
            else:
                messagebox.showerror("Erro Crítico", f"Falha ao limpar o arquivo Excel:\n{path_limpo}")
                self.lbl_status.configure(text="Erro ao limpar arquivo.", text_color="red")
        finally:
            self.configure(cursor="")

    def _limpar_planilha_sipac(self, caminho_original):
        """Abre o ficheiro no próprio Excel invisível e guarda uma cópia limpa e sem erros"""
        
        temp_dir = get_app_dir() / "Output"
        temp_dir.mkdir(exist_ok=True)
        
        # 1. FAZ UMA CÓPIA BRUTA PRIMEIRO (Tira o bloqueio do Windows)
        caminho_copia = str(temp_dir / "temp_original_desbloqueado.xlsx")
        try:
            shutil.copy2(caminho_original, caminho_copia)
        except Exception as e:
            return False, f"Falha ao tirar bloqueio de segurança: {e}"

        caminho_limpo = str(temp_dir / "temp_sintetico_limpo.xlsx")
        if os.path.exists(caminho_limpo):
            try: os.remove(caminho_limpo)
            except: pass

        excel = None
        wb = None
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # 2. Caminho Absoluto rigoroso para o Windows não se perder nas barras
            caminho_abs = os.path.abspath(caminho_copia)
            
            # 3. Abre em modo somente leitura e sem atualizar links (evita pop-ups que travam o script)
            wb = excel.Workbooks.Open(caminho_abs, UpdateLinks=False, ReadOnly=True)
            
            # 4. Salva como XLSX nativo limpo (FileFormat=51)
            wb.SaveAs(os.path.abspath(caminho_limpo), FileFormat=51) 
            
            wb.Close(False)
            excel.Quit()
            return True, caminho_limpo
        except Exception as e:
            if wb:
                try: wb.Close(False)
                except: pass
            if excel:
                try: excel.Quit()
                except: pass
            return False, f"Erro COM do Windows: {str(e)}"
        finally:
            try: pythoncom.CoUninitialize()
            except: pass

    def ler_colunas(self):
        if not getattr(self, 'sintetico_limpo_path', None): return
        try:
            l = int(self.ent_line.get()) - 1
            # Como o ficheiro agora está limpo, o pandas lê sem dar erros bizarros
            df = pd.read_excel(self.sintetico_limpo_path, header=l, nrows=5)
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

    # ==========================================================
    # LEITOR E RADAR INTELIGENTE (100% Funcional e Rápido)
    # ==========================================================
    def carregar_preview(self):
        if not getattr(self, 'sintetico_limpo_path', None): 
            return messagebox.showwarning("Erro", "Selecione o sintético primeiro.")
            
        self.configure(cursor="watch")
        self.lbl_status.configure(text="Carregando tabela...", text_color="orange")
        self.update()
        
        try:
            # Força atualizar as colunas para o caso do utilizador ter mudado o número da linha
            self.ler_colunas() 
            
            l = int(self.ent_line.get()) - 1
            
            # O Pandas agora carrega o ficheiro sem se preocupar com os XMLs corrompidos!
            df = pd.read_excel(self.sintetico_limpo_path, header=l)
            df.columns = [str(c).strip() for c in df.columns]
            
            c_i = self.combos_map["ITEM"].get()
            c_d = self.combos_map["DESCRICAO"].get()
            c_c = self.combos_map["CODIGO"].get()
            c_b = self.combos_map["BANCO"].get()
            c_u = self.combos_map["UNIT"].get()

            self.table_control.clear()
            
            # RADAR DE PARADA
            palavras_parada = ["TOTAL SEM BDI", "TOTAL DO BDI", "TOTAL GERAL", "VALOR GLOBAL", "CUSTO TOTAL"]
            
            for idx, row in df.iterrows():
                desc_val = str(row.get(c_d, 'nan')).strip()
                desc_upper = desc_val.upper()
                
                # Se o radar apitar com as palavras do rodapé, encerra o loop de leitura!
                if any(p in desc_upper for p in palavras_parada):
                    self.logger.info(f"🛑 Fim do orçamento detetado pelo radar na linha {l+idx+2}.")
                    break
                    
                if desc_val == 'nan' or desc_val == '' or desc_val == 'None': continue
                
                unit_val = row.get(c_u, 0)
                self.table_control.add_row(l+idx+2, row.get(c_i,''), desc_val, row.get(c_c,''), row.get(c_b,''), row, unit_val)
                
            self.tabview.set("🏗️ Painel de Orçamento")
            self.lbl_status.configure(text="Tabela carregada com sucesso.", text_color="lime")
        except Exception as e: 
            messagebox.showerror("Erro de Leitura", f"Ocorreu um erro ao carregar a tabela:\n{str(e)}")
            self.lbl_status.configure(text="Erro ao ler ficheiro.", text_color="red")
        finally: 
            self.configure(cursor="")

    def carregar_ui_perfil(self, nome):
        prof = self.dados_config["perfis"].get("PADRAO", {}).get("input", {})
        for k, v in prof.items():
            if k in self.combos_map: self.combos_map[k].set(v)

    def _obter_valor_seguro(self, entry_widget, is_date=False, custom_placeholder=None):
        valor = entry_widget.get().strip().upper() 
        if not valor:
            if custom_placeholder: return custom_placeholder
            return "xx/xx/xxxx" if is_date else "xxxxxxxxxx"
        return valor

    def executar(self):
        if not self.modelo_path: return messagebox.showwarning("Erro", "Nenhum modelo selecionado! Use o botão de engrenagem para adicionar um.")
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
             self.autocomplete.add_value(key, info[key])
        
        self.atualizar_listas_visuais()
        self._salvar_sessao_atual()
        
        self.btn_run.configure(state="disabled", text="Processando...")
        self.lbl_status.configure(text="Iniciando motor de cálculo...", text_color="orange")
        self.progress.pack(side="bottom", padx=10, pady=(0, 5), fill="x")
        self.progress.start()
        
        threading.Thread(target=self._run, args=(d, m, info)).start()

    def _run(self, d, m, p):
        start_time = time.time()
        eng = OrcamentoEngine({})
        ok, msg, extra_info = eng.gerar_excel_final(d, self.modelo_path, m, p)
        
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
        self.after(0, lambda: self._finish_run(ok, msg, duration, p, d, pdf_msg))

    def _finish_run(self, ok, msg, duration, info_data, raw_data, pdf_msg):
        self.progress.stop()
        self.progress.pack_forget()
        self.btn_run.configure(state="normal", text="🚀 GERAR ORÇAMENTO")
        
        if ok:
            try:
                dados_historico = {
                    'data_geracao': info_data.get('data'),
                    'nome_obra': info_data.get('nome_arquivo'),
                    'local': f"{info_data.get('campus')} - {info_data.get('setor')}",
                    'bdi': info_data.get('bdi'),
                    'valor_total': 0.0,
                    'arquivo_saida': msg,
                    'num_itens': len(raw_data),
                    'num_titulos': sum(1 for x in raw_data if x.get('_NIVEL_FORCADO') != 'ITEM'),
                    'duracao_processamento': round(duration, 2)
                }
                self.db_manager.inserir_orcamento(dados_historico)
                self.lbl_status.configure(text="Concluído com sucesso!", text_color="lime")
                self.logger.info("✅ Histórico salvo no banco de dados.")
            except Exception as e:
                self.logger.error(f"Erro ao salvar histórico: {e}")

            messagebox.showinfo("Sucesso", f"Salvo com sucesso em:\n{msg}{pdf_msg}")
            try: os.startfile(msg)
            except: pass
        else:
            self.lbl_status.configure(text="Erro no processamento.", text_color="red")
            if "Permission denied" in msg or "aberto" in msg:
                messagebox.showerror("Arquivo Aberto", f"O arquivo parece estar aberto no Excel.\n\nPor favor, feche o arquivo '{msg}' e tente novamente.")
            else:
                messagebox.showerror("Erro Fatal", msg)

    def _log_callback(self, n, m):
        self.log_box.insert("end", m + "\n")
        self.log_box.see("end")