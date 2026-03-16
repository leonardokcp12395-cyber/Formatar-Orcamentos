import customtkinter as ctk

class TopDashboard(ctk.CTkFrame):
    def __init__(self, master, main_window, **kwargs):
        super().__init__(master, fg_color="#1E1E1E", border_width=1, border_color="#3498DB", **kwargs)
        self.main_window = main_window

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        f_sint = ctk.CTkFrame(self, fg_color="transparent")
        f_sint.grid(row=0, column=0, padx=15, pady=10, sticky="nsew")
        ctk.CTkLabel(f_sint, text="1. Arraste ou Selecione o Sintético", font=("Arial", 13, "bold"), text_color="#3498DB").pack(anchor="w")

        f_sint_inner = ctk.CTkFrame(f_sint, fg_color="transparent")
        f_sint_inner.pack(fill="x", pady=5)
        ctk.CTkButton(f_sint_inner, text="📂 Selecionar", width=100, command=self.main_window.sel_sintetico).pack(side="left", padx=(0, 10))
        self.main_window.lbl_sint = ctk.CTkLabel(f_sint_inner, text="Nenhum arquivo", text_color="gray")
        self.main_window.lbl_sint.pack(side="left")

        ctk.CTkButton(f_sint, text="🔄 Carregar Tabela Visual", command=self.main_window.carregar_preview, fg_color="#E67E22").pack(anchor="w", pady=5)

        f_wpp = ctk.CTkFrame(self, fg_color="transparent")
        f_wpp.grid(row=0, column=1, padx=15, pady=10, sticky="nsew")
        ctk.CTkLabel(f_wpp, text="2. Importação Inteligente (WhatsApp)", font=("Arial", 13, "bold"), text_color="#8E44AD").pack(anchor="w")

        f_wpp_inner = ctk.CTkFrame(f_wpp, fg_color="transparent")
        f_wpp_inner.pack(fill="x", pady=5)
        self.main_window.txt_import = ctk.CTkTextbox(f_wpp_inner, height=55)
        self.main_window.txt_import.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(f_wpp_inner, text="🪄 Extrair", command=self.main_window.extrair_dados_texto, fg_color="#8E44AD", width=80, height=55).pack(side="right")

        self.main_window.switch_tema = ctk.CTkSwitch(self, text="Modo Escuro", command=self.main_window._alternar_tema)
        self.main_window.switch_tema.select()
        self.main_window.switch_tema.place(relx=0.98, rely=0.1, anchor="ne")