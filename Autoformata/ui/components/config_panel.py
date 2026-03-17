import customtkinter as ctk

class ConfigPanel(ctk.CTkScrollableFrame):
    def __init__(self, master, main_window, **kwargs):
        super().__init__(master, **kwargs)
        self.main_window = main_window

        f_leitura = ctk.CTkFrame(self)
        f_leitura.pack(fill="x", pady=5, padx=10)
        ctk.CTkLabel(f_leitura, text="1. Modelo e Sintético", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)

        f_l_inner = ctk.CTkFrame(f_leitura, fg_color="transparent")
        f_l_inner.pack(fill="x", padx=10, pady=5)

        self.inputs_refs = {}

        ctk.CTkLabel(f_l_inner, text="Modelo Excel:").pack(side="left")
        self.combo_modelos = ctk.CTkComboBox(f_l_inner, width=250, command=self.main_window._ao_trocar_modelo)
        self.combo_modelos.pack(side="left", padx=5)
        ctk.CTkButton(f_l_inner, text="⚙️ Gerenciar Modelos", width=140, fg_color="#555",
                      command=self.main_window._abrir_gerenciador_modelos).pack(side="left", padx=5)

        ctk.CTkLabel(f_l_inner, text="Ler Sintético a partir da Linha:").pack(side="left", padx=(20, 5))
        self.ent_line = ctk.CTkEntry(f_l_inner, width=50)
        self.ent_line.insert(0, "4")
        self.ent_line.pack(side="left")

        f_fin = ctk.CTkFrame(self)
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

        f_map = ctk.CTkFrame(self)
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

        self.main_window._atualizar_combo_modelos()

    def limpar_campos(self):
        for k, cb in self.combos_map.items():
            cb.set("...")