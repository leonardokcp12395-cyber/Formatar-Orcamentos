import customtkinter as ctk
from tkinter import ttk


class LevelSelector(ctk.CTkFrame):
    """
    Tabela visual para classificar linhas do orçamento (N1, N2, N3, ITEM, IGNORAR).
    Extraido de main_window.py para ficheiro próprio.
    
    Correções aplicadas:
    - ttk.Style: Nome "Excel.Treeview" (contém .Treeview para herdar layout base)
    - Âncoras: "e" em vez de "right" (Tkinter não suporta "right")
    - Conversor _parse_numeric seguro para valores Pandas nativos
    """

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.rows_data = []

        # ── ESTILO (FIX: nome tem obrigatoriamente de conter .Treeview) ──
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Excel.Treeview",
                        background="#2B2B2B",
                        foreground="white",
                        rowheight=25,
                        fieldbackground="#2B2B2B",
                        borderwidth=0)
        style.configure("Excel.Treeview.Heading",
                        background="#1f538d",
                        foreground="white",
                        font=("Arial", 11, "bold"))
        style.map("Excel.Treeview", background=[("selected", "#3498DB")])

        # ── TREEVIEW ──
        self.tree = ttk.Treeview(
            self,
            style="Excel.Treeview",
            columns=("L", "Item", "Cod", "Banco", "Desc", "Nivel"),
            show="headings"
        )

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(column=0, row=0, sticky="nsew")
        vsb.grid(column=1, row=0, sticky="ns")
        hsb.grid(column=0, row=1, sticky="ew")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.setup_headers()

        # ── BINDINGS ──
        self.tree.bind("<Double-1>", self._mudar_nivel)
        self.tree.bind("1", lambda e: self._definir_nivel_teclado("N1"))
        self.tree.bind("2", lambda e: self._definir_nivel_teclado("N2"))
        self.tree.bind("3", lambda e: self._definir_nivel_teclado("N3"))
        self.tree.bind("i", lambda e: self._definir_nivel_teclado("ITEM"))
        self.tree.bind("I", lambda e: self._definir_nivel_teclado("ITEM"))
        self.tree.bind("g", lambda e: self._definir_nivel_teclado("IGNORAR"))
        self.tree.bind("G", lambda e: self._definir_nivel_teclado("IGNORAR"))
        self.tree.bind("<space>", self._mudar_nivel)

        ctk.CTkLabel(
            self,
            text="Atalhos: Selecione as linhas e aperte 1, 2, 3, I (Item), G (Ignorar) ou Espaço",
            font=("Arial", 11, "bold"), text_color="gray"
        ).grid(column=0, row=2, pady=5)

    def setup_headers(self):
        self.tree.heading("L", text="L")
        self.tree.heading("Item", text="Item")
        self.tree.heading("Cod", text="Cód.")
        self.tree.heading("Banco", text="Banco")
        self.tree.heading("Desc", text="Descrição")
        self.tree.heading("Nivel", text="Nível")

        # FIX: anchor="e" (East) para colunas monetárias — Tkinter não suporta "right"
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

        has_value = False
        parsed = self._parse_numeric(unit_val)
        if parsed is not None and parsed > 0:
            has_value = True

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

    @staticmethod
    def _parse_numeric(val):
        """
        Conversor seguro: deteta se o valor do Pandas já é float/int nativo
        e preserva-o. Se for string, faz .replace de forma segura.
        """
        if val is None:
            return None

        # Se já for numérico nativo do Pandas/Python, preserva
        if isinstance(val, (int, float)):
            return float(val)

        # Se for string, tenta converter
        try:
            s = str(val).replace('R$', '').strip()
            if not s or s.lower() == 'nan' or s.lower() == 'none':
                return None
            # Formato PT-BR: 1.500,50 → 1500.50
            if ',' in s and '.' in s:
                s = s.replace('.', '').replace(',', '.')
            elif ',' in s:
                s = s.replace(',', '.')
            return float(s)
        except (ValueError, TypeError):
            return None

    @staticmethod
    def format_ptbr(val):
        """Formata valor numérico para padrão PT-BR (ex: 1.500,50)."""
        if val is None:
            return ""
        try:
            return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return str(val)

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
