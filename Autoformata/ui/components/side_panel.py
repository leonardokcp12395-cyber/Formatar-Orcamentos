import customtkinter as ctk

class SidePanel(ctk.CTkScrollableFrame):
    def __init__(self, master, main_window, **kwargs):
        super().__init__(master, fg_color="transparent", width=310, **kwargs)
        self.main_window = main_window

        f_top_dados = ctk.CTkFrame(self, fg_color="transparent")
        f_top_dados.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(f_top_dados, text="Dados da Obra", font=("Arial", 14, "bold")).pack(side="left")

        f_top_btns = ctk.CTkFrame(f_top_dados, fg_color="transparent")
        f_top_btns.pack(side="right")

        ctk.CTkButton(f_top_btns, text="🧹 Novo", command=self.main_window.limpar_dados_sessao,
                      width=60, height=24, fg_color="#E74C3C", hover_color="#C0392B").pack(side="left", padx=(0, 5))
        ctk.CTkButton(f_top_btns, text="📝 Listas", command=self.main_window.abrir_editor_db,
                      width=60, height=24, fg_color="#444").pack(side="left")

        self.inputs_refs = {}
        self.combos_db_refs = {}

        ctk.CTkLabel(self, text="Nome (Arquivo):", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_nome_arquivo = ctk.CTkEntry(self, width=290)
        self.ent_nome_arquivo.pack(anchor="w", pady=(0, 5))
        self.inputs_refs["nome_arquivo"] = self.ent_nome_arquivo

        ctk.CTkLabel(self, text="Descrição (C15):", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_desc_cabecalho = ctk.CTkEntry(self, width=290)
        self.ent_desc_cabecalho.pack(anchor="w", pady=(0, 10))
        self.inputs_refs["descricao_header"] = self.ent_desc_cabecalho

        grid_f = ctk.CTkFrame(self, fg_color="#212121")
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

        f_p_inner = ctk.CTkFrame(self, fg_color="transparent")
        f_p_inner.pack(fill="x", pady=10)
        ctk.CTkLabel(f_p_inner, text="Valor Sim. (R$):", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_valor_sim = ctk.CTkEntry(f_p_inner, width=290, placeholder_text="0,00")
        self.ent_valor_sim.pack(anchor="w", pady=(0, 5))
        self.ent_valor_sim.bind("<FocusOut>", self.main_window._calcular_prazo_auto)
        self.ent_valor_sim.bind("<Return>", self.main_window._calcular_prazo_auto)

        ctk.CTkLabel(f_p_inner, text="Prazo Final:", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_prazo = ctk.CTkEntry(f_p_inner, width=290)
        self.ent_prazo.pack(anchor="w")
        self.inputs_refs["prazo"] = self.ent_prazo

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

    def limpar_campos(self):
        for key, widget in self.inputs_refs.items():
            if isinstance(widget, ctk.CTkEntry):
                widget.delete(0, 'end')
            elif isinstance(widget, ctk.CTkComboBox):
                widget.set("")
        if hasattr(self, 'ent_valor_sim'):
            self.ent_valor_sim.delete(0, 'end')