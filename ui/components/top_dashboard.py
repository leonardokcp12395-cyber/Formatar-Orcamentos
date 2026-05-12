import customtkinter as ctk


class TopDashboard(ctk.CTkFrame):
    """
    Barra superior — Seleção de Sintético + Importação Inteligente (WhatsApp).
    Todos os widgets são self. internos. Comunica via callbacks.
    """

    def __init__(self, master, *, on_select_file, on_load_preview,
                 on_extract_text, on_toggle_theme, on_kill_excel=None, **kwargs):
        super().__init__(master, fg_color="#1E1E1E", border_width=1,
                         border_color="#3498DB", **kwargs)

        self._on_select_file = on_select_file
        self._on_load_preview = on_load_preview
        self._on_extract_text = on_extract_text
        self._on_toggle_theme = on_toggle_theme
        self._on_kill_excel = on_kill_excel

        self._build_ui()

    # ──────────────────────────────────────────────
    #  CONSTRUÇÃO DA UI
    # ──────────────────────────────────────────────

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Coluna esquerda: Sintético
        f_sint = ctk.CTkFrame(self, fg_color="transparent")
        f_sint.grid(row=0, column=0, padx=15, pady=10, sticky="nsew")
        ctk.CTkLabel(f_sint, text="1. Arraste ou Selecione o Sintético",
                     font=("Arial", 13, "bold"), text_color="#3498DB").pack(anchor="w")

        f_sint_inner = ctk.CTkFrame(f_sint, fg_color="transparent")
        f_sint_inner.pack(fill="x", pady=5)
        ctk.CTkButton(f_sint_inner, text="📂 Selecionar", width=100,
                      command=self._on_select_file).pack(side="left", padx=(0, 10))
        self.lbl_sint = ctk.CTkLabel(f_sint_inner, text="Nenhum arquivo", text_color="gray")
        self.lbl_sint.pack(side="left")

        ctk.CTkButton(f_sint, text="🔄 Carregar Tabela Visual",
                      command=self._on_load_preview, fg_color="#E67E22").pack(anchor="w", pady=5)

        # Coluna direita: WhatsApp
        f_wpp = ctk.CTkFrame(self, fg_color="transparent")
        f_wpp.grid(row=0, column=1, padx=15, pady=10, sticky="nsew")
        ctk.CTkLabel(f_wpp, text="2. Importação Inteligente (WhatsApp)",
                     font=("Arial", 13, "bold"), text_color="#8E44AD").pack(anchor="w")

        f_wpp_inner = ctk.CTkFrame(f_wpp, fg_color="transparent")
        f_wpp_inner.pack(fill="x", pady=5)
        self.txt_import = ctk.CTkTextbox(f_wpp_inner, height=55)
        self.txt_import.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(f_wpp_inner, text="🪄 Extrair",
                      command=self._on_extract_text,
                      fg_color="#8E44AD", width=80, height=55).pack(side="right")

        # Switch Tema
        self.switch_tema = ctk.CTkSwitch(
            self, text="Modo Escuro", command=self._on_toggle_theme)
        self.switch_tema.select()
        self.switch_tema.place(relx=0.98, rely=0.1, anchor="ne")

        # Botão Pânico Excel (Zombie Killer)
        if hasattr(self, '_on_kill_excel') and self._on_kill_excel:
            btn_panic = ctk.CTkButton(self, text="🚨 Desbugar Excel", 
                                      command=self._on_kill_excel, 
                                      fg_color="#C0392B", hover_color="#A93226", width=120)
            btn_panic.place(relx=0.98, rely=0.5, anchor="ne")

    # ──────────────────────────────────────────────
    #  API PÚBLICA
    # ──────────────────────────────────────────────

    def get_import_text(self) -> str:
        """Retorna o texto colado no campo de importação."""
        return self.txt_import.get("0.0", "end").strip()

    def clear_import_text(self):
        """Limpa o campo de importação."""
        self.txt_import.delete("0.0", "end")

    def set_file_label(self, name: str, color: str = "lime"):
        """Atualiza a label do ficheiro selecionado."""
        self.lbl_sint.configure(text=name, text_color=color)

    def get_theme_state(self) -> int:
        """Retorna o estado do switch de tema (1 = escuro, 0 = claro)."""
        return self.switch_tema.get()