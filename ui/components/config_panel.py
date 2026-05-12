import customtkinter as ctk


class ConfigPanel(ctk.CTkScrollableFrame):
    """
    Aba de Configurações & Mapeamento.
    Todos os widgets são self. internos. Comunica com o exterior via
    get_data(), get_column_mapping(), update_column_options() e callbacks.
    """

    def __init__(self, master, *, on_model_change, on_manage_models, **kwargs):
        super().__init__(master, **kwargs)
        self._on_model_change = on_model_change
        self._on_manage_models = on_manage_models

        self.combos_map = {}
        self._build_ui()

    # ──────────────────────────────────────────────
    #  CONSTRUÇÃO DA UI
    # ──────────────────────────────────────────────

    def _build_ui(self):
        # === SEÇÃO 1: Modelo e Sintético ===
        f_leitura = ctk.CTkFrame(self)
        f_leitura.pack(fill="x", pady=5, padx=10)
        ctk.CTkLabel(f_leitura, text="1. Modelo e Sintético",
                     font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)

        f_l_inner = ctk.CTkFrame(f_leitura, fg_color="transparent")
        f_l_inner.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(f_l_inner, text="Modelo Excel:").pack(side="left")
        self.combo_modelos = ctk.CTkComboBox(
            f_l_inner, width=250, command=self._on_model_selected)
        self.combo_modelos.pack(side="left", padx=5)
        ctk.CTkButton(f_l_inner, text="⚙️ Gerenciar Modelos", width=140, fg_color="#555",
                      command=self._on_manage_models).pack(side="left", padx=5)

        ctk.CTkLabel(f_l_inner, text="Ler Sintético a partir da Linha:").pack(side="left", padx=(20, 5))
        self.ent_line = ctk.CTkEntry(f_l_inner, width=50)
        self.ent_line.insert(0, "4")
        self.ent_line.pack(side="left")

        # === SEÇÃO 2: Financeiro & Saída ===
        f_fin = ctk.CTkFrame(self)
        f_fin.pack(fill="x", pady=10, padx=10)
        ctk.CTkLabel(f_fin, text="2. Financeiro & Saída",
                     font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)

        fin_grid = ctk.CTkFrame(f_fin, fg_color="transparent")
        fin_grid.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(fin_grid, text="BDI:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.combo_bdi = ctk.CTkComboBox(
            fin_grid, width=250,
            values=["28,82% (SUP - Desc 0,19)", "35,18% (PRUMO - Desc 0,0601)", "0,00%"])
        self.combo_bdi.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ctk.CTkLabel(fin_grid, text="Método de Cálculo:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.combo_metodo_calc = ctk.CTkComboBox(
            fin_grid, width=250,
            values=["Cortar Casas (Padrão - Ignora resto)",
                    "Arredondar (2 Casas - Matemático)",
                    "Exato (Sem tratamento - Excel)"])
        self.combo_metodo_calc.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.combo_metodo_calc.set("Cortar Casas (Padrão - Ignora resto)")

        ctk.CTkLabel(fin_grid, text="Altura Linha (Px):").grid(row=0, column=2, padx=(20, 5), pady=5, sticky="w")
        self.ent_altura = ctk.CTkEntry(fin_grid, width=100)
        self.ent_altura.insert(0, "24.75")
        self.ent_altura.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        self.chk_pdf = ctk.CTkCheckBox(
            fin_grid, text="Gerar PDF automaticamente após Excel", text_color="lime")
        self.chk_pdf.grid(row=1, column=2, columnspan=2, padx=(20, 5), pady=5, sticky="w")

        # === SEÇÃO 3: Mapeamento de Colunas ===
        f_map = ctk.CTkFrame(self)
        f_map.pack(fill="x", pady=10, padx=10)
        ctk.CTkLabel(f_map, text="3. Mapeamento de Colunas",
                     font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)

        f_map_inner = ctk.CTkFrame(f_map, fg_color="transparent")
        f_map_inner.pack(fill="x", padx=10, pady=5)

        campos = ["ITEM", "CODIGO", "BANCO", "DESCRICAO", "UNID", "QUANT", "UNIT"]
        for i, camp in enumerate(campos):
            r = i // 2
            c = (i % 2) * 2
            ctk.CTkLabel(f_map_inner, text=f"{camp}:", width=80, anchor="e").grid(
                row=r, column=c, padx=5, pady=5)
            cb = ctk.CTkComboBox(f_map_inner, values=["..."], width=200)
            cb.grid(row=r, column=c + 1, padx=5, pady=5)
            self.combos_map[camp] = cb

    def _on_model_selected(self, nome):
        """Chamado internamente quando o combo de modelos muda."""
        if self._on_model_change:
            self._on_model_change(nome)

    # ──────────────────────────────────────────────
    #  API PÚBLICA
    # ──────────────────────────────────────────────

    def get_data(self) -> dict:
        """Retorna configurações financeiras e de saída."""
        bdi_str = self.combo_bdi.get().split('%')[0].replace(',', '.')
        try:
            bdi = float(bdi_str) / 100
        except (ValueError, TypeError):
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
        except (ValueError, TypeError):
            altura = 24.75

        return {
            "bdi": bdi,
            "calc_mode": calc_mode,
            "altura_linha": altura,
            "gerar_pdf": self.chk_pdf.get(),
            "start_line": self.ent_line.get(),
        }

    def get_column_mapping(self) -> dict:
        """Retorna o mapeamento chave → nome da coluna selecionada."""
        return {k: cb.get() for k, cb in self.combos_map.items()}

    def get_start_line(self) -> int:
        """Retorna a linha inicial de leitura (0-indexed para Pandas)."""
        try:
            return int(self.ent_line.get()) - 1
        except (ValueError, TypeError):
            return 3  # fallback: linha 4 do Excel → 3 no Pandas

    def update_column_options(self, cols: list):
        """Atualiza as opções de todas as ComboBoxes de mapeamento."""
        for k, cb in self.combos_map.items():
            cb.configure(values=cols)
            for c in cols:
                cu = c.upper()
                if k in cu:
                    cb.set(c)
                if k == "CODIGO" and "COD" in cu:
                    cb.set(c)
                if k == "BANCO" and ("FONTE" in cu or "REF" in cu):
                    cb.set(c)
                if k == "DESCRICAO" and "DESC" in cu:
                    cb.set(c)
                if k == "UNIT" and "VALOR" in cu:
                    cb.set(c)

    def reset_column_mapping(self):
        """Reseta todos os combos de mapeamento."""
        for k, cb in self.combos_map.items():
            cb.set("...")

    def update_model_list(self, names: list):
        """Atualiza lista de modelos disponíveis."""
        if names:
            self.combo_modelos.configure(values=names)
            atual = self.combo_modelos.get()
            if atual not in names:
                self.combo_modelos.set(names[0])
                self._on_model_selected(names[0])
        else:
            self.combo_modelos.configure(values=["(Nenhum modelo)"])
            self.combo_modelos.set("(Nenhum modelo)")

    def set_profile_mapping(self, mapping: dict):
        """Aplica mapeamento de colunas de um perfil."""
        for k, v in mapping.items():
            if k in self.combos_map:
                self.combos_map[k].set(v)