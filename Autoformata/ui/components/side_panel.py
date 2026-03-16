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

        ctk.CTkLabel(self, text="Nome (Arquivo):", font=("Arial", 11, "bold")).pack(anchor="w")
        self.main_window.ent_nome_arquivo = ctk.CTkEntry(self, width=290)
        self.main_window.ent_nome_arquivo.pack(anchor="w", pady=(0, 5))
        self.main_window.inputs_refs["nome_arquivo"] = self.main_window.ent_nome_arquivo

        ctk.CTkLabel(self, text="Descrição (C15):", font=("Arial", 11, "bold")).pack(anchor="w")
        self.main_window.ent_desc_cabecalho = ctk.CTkEntry(self, width=290)
        self.main_window.ent_desc_cabecalho.pack(anchor="w", pady=(0, 10))
        self.main_window.inputs_refs["descricao_header"] = self.main_window.ent_desc_cabecalho

        grid_f = ctk.CTkFrame(self, fg_color="#212121")
        grid_f.pack(fill="x", pady=5)

        self.main_window.cbo_campus = self.main_window._add_side_field(grid_f, "Campus:", "campus")
        self.main_window.cbo_setor = self.main_window._add_side_field(grid_f, "Setor:", "setor")
        self.main_window.cbo_fiscal = self.main_window._add_side_field(grid_f, "Fiscal:", "fiscal")
        self.main_window.cbo_servidor = self.main_window._add_side_field(grid_f, "Servidor:", "servidor")
        self.main_window.cbo_elab = self.main_window._add_side_field(grid_f, "Elaborador:", "elaborador")
        self.main_window.cbo_estag = self.main_window._add_side_field(grid_f, "Estagiário:", "estagiario")

        self.main_window.ent_data = self.main_window._add_side_input(grid_f, "Data Elab.:", "data")
        self.main_window.ent_orcafascio = self.main_window._add_side_input(grid_f, "Orçafascio:", "orcafascio")
        self.main_window.ent_processo = self.main_window._add_side_input(grid_f, "Processo:", "processo")
        self.main_window.ent_num_orc = self.main_window._add_side_input(grid_f, "Nº Orc:", "num_orcamento")
        self.main_window.ent_empenho = self.main_window._add_side_input(grid_f, "Empenho:", "empenho")
        self.main_window.ent_data_emissao = self.main_window._add_side_input(grid_f, "Emissão:", "data_emissao")
        self.main_window.ent_data_inicio = self.main_window._add_side_input(grid_f, "Início:", "data_inicio")

        f_p_inner = ctk.CTkFrame(self, fg_color="transparent")
        f_p_inner.pack(fill="x", pady=10)
        ctk.CTkLabel(f_p_inner, text="Valor Sim. (R$):", font=("Arial", 11, "bold")).pack(anchor="w")
        self.main_window.ent_valor_sim = ctk.CTkEntry(f_p_inner, width=290, placeholder_text="0,00")
        self.main_window.ent_valor_sim.pack(anchor="w", pady=(0, 5))
        self.main_window.ent_valor_sim.bind("<FocusOut>", self.main_window._calcular_prazo_auto)
        self.main_window.ent_valor_sim.bind("<Return>", self.main_window._calcular_prazo_auto)

        ctk.CTkLabel(f_p_inner, text="Prazo Final:", font=("Arial", 11, "bold")).pack(anchor="w")
        self.main_window.ent_prazo = ctk.CTkEntry(f_p_inner, width=290)
        self.main_window.ent_prazo.pack(anchor="w")
        self.main_window.inputs_refs["prazo"] = self.main_window.ent_prazo