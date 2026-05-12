import customtkinter as ctk


class SidePanel(ctk.CTkScrollableFrame):
    """
    Painel lateral esquerdo — Dados da Obra.
    Todos os widgets são self. internos. Comunica com o exterior via
    get_data(), set_data(), limpar_campos() e callbacks.
    """

    def __init__(self, master, *, on_limpar, on_editor_db, **kwargs):
        super().__init__(master, fg_color="transparent", width=310, **kwargs)
        self._on_limpar = on_limpar
        self._on_editor_db = on_editor_db

        # Referências internas dos widgets
        self._combos_db = {}   # chave DB → widget ComboBox (para atualizar listas)
        self._all_inputs = {}  # chave → widget (Entry ou ComboBox, para sessão)

        self._build_ui()

    # ──────────────────────────────────────────────
    #  CONSTRUÇÃO DA UI
    # ──────────────────────────────────────────────

    def _build_ui(self):
        # Topo: título + botões
        f_top = ctk.CTkFrame(self, fg_color="transparent")
        f_top.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(f_top, text="Dados da Obra", font=("Arial", 14, "bold")).pack(side="left")

        f_btns = ctk.CTkFrame(f_top, fg_color="transparent")
        f_btns.pack(side="right")
        ctk.CTkButton(f_btns, text="🧹 Novo", command=self._on_limpar,
                      width=60, height=24, fg_color="#E74C3C", hover_color="#C0392B").pack(side="left", padx=(0, 5))
        ctk.CTkButton(f_btns, text="📝 Listas", command=self._on_editor_db,
                      width=60, height=24, fg_color="#444").pack(side="left")

        # Nome do arquivo
        ctk.CTkLabel(self, text="Nome (Arquivo):", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_nome_arquivo = ctk.CTkEntry(self, width=290)
        self.ent_nome_arquivo.pack(anchor="w", pady=(0, 5))
        self._all_inputs["nome_arquivo"] = self.ent_nome_arquivo

        # Descrição cabeçalho
        ctk.CTkLabel(self, text="Descrição (C15):", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_desc_cabecalho = ctk.CTkEntry(self, width=290)
        self.ent_desc_cabecalho.pack(anchor="w", pady=(0, 10))
        self._all_inputs["descricao_header"] = self.ent_desc_cabecalho

        # Grid de combos + inputs
        grid_f = ctk.CTkFrame(self, fg_color="#212121")
        grid_f.pack(fill="x", pady=5)

        self.cbo_campus = self._add_field(grid_f, "Campus:", "campus")
        self.cbo_setor = self._add_field(grid_f, "Setor:", "setor")
        self.cbo_fiscal = self._add_field(grid_f, "Fiscal:", "fiscal")
        self.cbo_servidor = self._add_field(grid_f, "Servidor:", "servidor")
        self.cbo_elab = self._add_field(grid_f, "Elaborador:", "elaborador")
        self.cbo_estag = self._add_field(grid_f, "Estagiário:", "estagiario")

        self.ent_data = self._add_input(grid_f, "Data Elab.:", "data")
        self.ent_orcafascio = self._add_input(grid_f, "Orçafascio:", "orcafascio")
        self.ent_processo = self._add_input(grid_f, "Processo:", "processo")
        self.ent_num_orc = self._add_input(grid_f, "Nº Orc:", "num_orcamento")
        self.ent_empenho = self._add_input(grid_f, "Empenho:", "empenho")
        self.ent_data_emissao = self._add_input(grid_f, "Emissão:", "data_emissao")
        self.ent_data_inicio = self._add_input(grid_f, "Início:", "data_inicio")

        # Valor simulado + Prazo
        f_p = ctk.CTkFrame(self, fg_color="transparent")
        f_p.pack(fill="x", pady=10)

        ctk.CTkLabel(f_p, text="Valor Sim. (R$):", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_valor_sim = ctk.CTkEntry(f_p, width=290, placeholder_text="0,00")
        self.ent_valor_sim.pack(anchor="w", pady=(0, 5))
        self.ent_valor_sim.bind("<FocusOut>", self._calcular_prazo_auto)
        self.ent_valor_sim.bind("<Return>", self._calcular_prazo_auto)

        ctk.CTkLabel(f_p, text="Prazo Final:", font=("Arial", 11, "bold")).pack(anchor="w")
        self.ent_prazo = ctk.CTkEntry(f_p, width=290)
        self.ent_prazo.pack(anchor="w")
        self._all_inputs["prazo"] = self.ent_prazo

    # ──────────────────────────────────────────────
    #  HELPERS INTERNOS
    # ──────────────────────────────────────────────

    def _add_field(self, parent, label, db_key):
        """Cria um campo com ComboBox ligado a uma chave do autocomplete."""
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", pady=2, padx=5)
        ctk.CTkLabel(f, text=label, font=("Arial", 11, "bold"),
                     width=85, anchor="w").pack(side="left")
        w = ctk.CTkComboBox(f, width=190, values=[])
        w.set("")
        w.pack(side="right")
        self._combos_db[db_key] = w
        self._all_inputs[db_key] = w
        return w

    def _add_input(self, parent, label, ref_key):
        """Cria um campo de texto simples."""
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", pady=2, padx=5)
        ctk.CTkLabel(f, text=label, font=("Arial", 11, "bold"),
                     width=85, anchor="w").pack(side="left")
        w = ctk.CTkEntry(f, width=190)
        w.pack(side="right")
        self._all_inputs[ref_key] = w
        return w

    def _calcular_prazo_auto(self, event=None):
        """Calcula prazo automaticamente com base no valor simulado."""
        try:
            val_txt = self.ent_valor_sim.get().replace(
                'R$', '').replace('.', '').replace(',', '.').strip()
            if not val_txt:
                return
            valor = float(val_txt)
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
        except Exception:
            pass

    # ──────────────────────────────────────────────
    #  API PÚBLICA
    # ──────────────────────────────────────────────

    def get_data(self) -> dict:
        """Recolhe todos os valores dos inputs num dicionário limpo."""
        data = {}
        for key, widget in self._all_inputs.items():
            try:
                data[key] = widget.get()
            except Exception:
                data[key] = ""
        # Valor simulado não está em _all_inputs por design (nunca vai para a sessão)
        return data

    def set_data(self, data: dict):
        """Preenche os widgets a partir de um dicionário."""
        for key, valor in data.items():
            if key in self._all_inputs and valor:
                widget = self._all_inputs[key]
                if isinstance(widget, ctk.CTkEntry):
                    widget.delete(0, 'end')
                    widget.insert(0, valor)
                elif isinstance(widget, ctk.CTkComboBox):
                    widget.set(valor)

    def limpar_campos(self):
        """Reseta todos os formulários."""
        for key, widget in self._all_inputs.items():
            if isinstance(widget, ctk.CTkEntry):
                widget.delete(0, 'end')
            elif isinstance(widget, ctk.CTkComboBox):
                widget.set("")
        self.ent_valor_sim.delete(0, 'end')

    def atualizar_listas(self, autocomplete_mgr):
        """Atualiza os valores dos ComboBoxes com dados do autocomplete manager."""
        for db_key, widget in self._combos_db.items():
            lista = autocomplete_mgr.get_list(db_key)
            if lista:
                widget.configure(values=lista)
            else:
                widget.configure(values=["(Digite um novo...)"])

    def get_db_keys(self) -> list:
        """Retorna as chaves DB que este painel gere (para salvar autocomplete)."""
        return list(self._combos_db.keys())

    def fill_from_extracted(self, dados: dict):
        """Preenche campos a partir de dados extraídos pelo SmartParser."""
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

        # Preenche nome do arquivo a partir da descrição, se vazio
        if "descricao_header" in dados and not self.ent_nome_arquivo.get():
            nome_seguro = "".join(
                x for x in dados["descricao_header"][:40] if x.isalnum() or x in " -_")
            self.ent_nome_arquivo.insert(0, nome_seguro)

        return count