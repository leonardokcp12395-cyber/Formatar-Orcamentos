import customtkinter as ctk
from tkinter import ttk

class ExcelPreview(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        style = ttk.Style(self)
        style.theme_use("default")

        # CORREÇÃO: O nome do estilo deve conter ".Treeview" para o Tkinter herdar o layout base
        style.configure("Excel.Treeview",
                        background="#FFFFFF",
                        foreground="#000000",
                        rowheight=25,
                        fieldbackground="#FFFFFF",
                        borderwidth=1,
                        relief="solid")

        style.configure("Excel.Treeview.Heading",
                        background="#F3F3F3",
                        foreground="#000000",
                        font=("Arial", 10, "bold"),
                        borderwidth=1,
                        relief="solid")

        # CORREÇÃO APLICADA AQUI NO PARÂMETRO STYLE
        self.tree = ttk.Treeview(self, style="Excel.Treeview", show="headings")

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(column=0, row=0, sticky="nsew")
        vsb.grid(column=1, row=0, sticky="ns")
        hsb.grid(column=0, row=1, sticky="ew")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.setup_headers()
        self.setup_tags()

    def setup_headers(self):
        cols = ("A", "B", "C", "D", "E", "F", "G", "H")
        self.tree["columns"] = cols

        self.tree.heading("A", text="ITEM")
        self.tree.heading("B", text="CÓDIGO")
        self.tree.heading("C", text="BANCO")
        self.tree.heading("D", text="DESCRIÇÃO")
        self.tree.heading("E", text="UNID")
        self.tree.heading("F", text="QUANT")
        self.tree.heading("G", text="UNIT")
        self.tree.heading("H", text="TOTAL")

        self.tree.column("A", width=50, anchor="center")
        self.tree.column("B", width=80, anchor="center")
        self.tree.column("C", width=80, anchor="center")
        self.tree.column("D", width=400, anchor="w")
        self.tree.column("E", width=50, anchor="center")
        self.tree.column("F", width=60, anchor="right")
        self.tree.column("G", width=80, anchor="right")
        self.tree.column("H", width=80, anchor="right")

    def setup_tags(self):
        # Configure tags for Excel simulation
        self.tree.tag_configure("N1", background="#9BC2E6", font=("Arial", 11, "bold"))
        self.tree.tag_configure("N2", background="#BDD7EE", font=("Arial", 11, "bold"))
        self.tree.tag_configure("N3", background="#DDEBF7", font=("Arial", 11, "bold"))
        self.tree.tag_configure("ITEM", background="#FFFFFF", font=("Arial", 10))

    def clear(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    @staticmethod
    def _parse_numeric(val):
        if val is None or val == "" or val == "nan":
            return None
        if isinstance(val, (int, float)):
            return float(val)

        s = str(val).strip().replace('R$', '').strip()
        if ',' in s and '.' in s:
            s = s.replace('.', '').replace(',', '.')
        elif ',' in s:
            s = s.replace(',', '.')

        try:
            return float(s)
        except ValueError:
            return None

    @staticmethod
    def _format_pt_br(val, is_currency=False):
        if val is None:
            return ""
        s = f"{val:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
        if is_currency:
            return f"R$ {s}"
        return s

    def popular_dados(self, dados_linhas):
        self.clear()

        for data in dados_linhas:
            nivel = data["nivel"]
            if nivel == "IGNORAR":
                continue

            raw = data["raw_data"]

            item = raw.get("ITEM_SIM", "")
            cod = raw.get("COD_SIM", "")
            banco = raw.get("BANCO_SIM", "")
            desc = raw.get("DESC_SIM", "")
            unid = raw.get("UNID_SIM", "") if nivel == "ITEM" else ""

            quant_raw = raw.get("QUANT_SIM", "") if nivel == "ITEM" else ""
            unit_raw = raw.get("UNIT_SIM", "") if nivel == "ITEM" else ""

            quant_str = ""
            unit_str = ""
            total_str = ""

            if nivel == "ITEM":
                q = self._parse_numeric(quant_raw)
                u = self._parse_numeric(unit_raw)

                if q is not None:
                    quant_str = self._format_pt_br(q, is_currency=False)
                if u is not None:
                    unit_str = self._format_pt_br(u, is_currency=True)
                if q is not None and u is not None:
                    total_str = self._format_pt_br(q * u, is_currency=True)

            self.tree.insert("", "end", values=(item, cod, banco, desc, unid, quant_str, unit_str, total_str), tags=(nivel,))
