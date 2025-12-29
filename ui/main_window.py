"""
SISORC ULTIMATE v13.0 - Interface Principal CORRIGIDA
Melhorias: Validações, Feedback visual, Tratamento de erros
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import threading
from pathlib import Path
from typing import Optional
from core.sanitizer import ExcelSanitizer
from core.excel_handler import OrcamentoEngine
from core.database import DatabaseManager
from utils.helpers import ConfigLoader, FileValidator, DataValidator
from utils.logger import Logger, LogLevel
from datetime import datetime

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class LogViewer(ctk.CTkTextbox):
    """Visualizador de logs em tempo real"""
    
    CORES = {
        LogLevel.INFO: "#10b981",
        LogLevel.WARNING: "#f59e0b",
        LogLevel.ERROR: "#ef4444",
        LogLevel.DEBUG: "#6b7280"
    }
    
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.configure(state="disabled", wrap="word")
    
    def adicionar_log(self, nivel: int, mensagem: str):
        """Adiciona mensagem de log com cor"""
        self.configure(state="normal")
        
        # Determina cor
        cor = self.CORES.get(nivel, "white")
        
        # Adiciona texto
        self.insert("end", mensagem + "\n")
        
        # Scroll automático
        self.see("end")
        
        self.configure(state="disabled")
    
    def limpar(self):
        """Limpa o log"""
        self.configure(state="normal")
        self.delete("1.0", "end")
        self.configure(state="disabled")


class LevelSelector(ctk.CTkScrollableFrame):
    """Seletor de níveis com visualização melhorada"""
    
    CORES_NIVEL = {
        "N1": {"texto": "#9BC2E6", "bg": "#1a4d7a"},
        "N2": {"texto": "#BDD7EE", "bg": "#2d5f8f"},
        "N3": {"texto": "#DDEBF7", "bg": "#3d6f9f"},
        "ITEM": {"texto": "white", "bg": "transparent"}
    }
    
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.decisoes = {}
        self.row_widgets = {}
        self.df_cache = None
        self._criar_cabecalho()
    
    def _criar_cabecalho(self):
        """Cria cabeçalho da tabela"""
        headers = [
            ("#", 50),
            ("Item", 100),
            ("Descrição", 450),
            ("Nível", 300)
        ]
        
        for i, (texto, largura) in enumerate(headers):
            lbl = ctk.CTkLabel(
                self,
                text=texto,
                font=("Arial", 13, "bold"),
                width=largura,
                anchor="w"
            )
            lbl.grid(row=0, column=i, padx=5, pady=(5, 10), sticky="w")
    
    def carregar_dados(self, df: pd.DataFrame):
        """Carrega dados do DataFrame"""
        self._limpar_tabela()
        self.df_cache = df.copy()
        
        # Detecta colunas
        col_item = self._detectar_coluna(df, ["ITEM", "IT", "NUM"])
        col_desc = self._detectar_coluna(df, ["DESCRIÇÃO", "DESCRIÇAO", "DISCRIMINAÇÃO"])
        
        if not col_item or not col_desc:
            Logger.warning("Usando primeira e segunda colunas como fallback")
            col_item = df.columns[0]
            col_desc = df.columns[1] if len(df.columns) > 1 else df.columns[0]
        
        Logger.info(f"Colunas: Item='{col_item}', Desc='{col_desc}'")
        
        # Renderiza linhas
        for idx, row in df.iterrows():
            self._criar_linha(idx, row, col_item, col_desc)
        
        Logger.info(f"✓ {len(df)} linhas carregadas")
    
    def _detectar_coluna(self, df: pd.DataFrame, keywords: list) -> Optional[str]:
        """Detecta coluna por palavras-chave"""
        for col in df.columns:
            col_upper = str(col).upper()
            if any(kw.upper() in col_upper for kw in keywords):
                return col
        return None
    
    def _criar_linha(self, idx: int, row: pd.Series, col_item: str, col_desc: str):
        """Cria linha visual"""
        linha_visual = idx + 1
        
        # Valores
        val_item = str(row[col_item]) if pd.notna(row[col_item]) else ""
        val_desc = str(row[col_desc]) if pd.notna(row[col_desc]) else ""
        val_desc_trunc = (val_desc[:65] + "...") if len(val_desc) > 65 else val_desc
        
        # Labels
        lbl_idx = ctk.CTkLabel(self, text=str(linha_visual), width=50, anchor="center")
        lbl_idx.grid(row=linha_visual, column=0, padx=3, pady=2)
        
        lbl_item = ctk.CTkLabel(self, text=val_item, width=100, anchor="w")
        lbl_item.grid(row=linha_visual, column=1, padx=3, pady=2)
        
        lbl_desc = ctk.CTkLabel(
            self, 
            text=val_desc_trunc, 
            width=450, 
            anchor="w",
            wraplength=440
        )
        lbl_desc.grid(row=linha_visual, column=2, padx=3, pady=2)
        
        self.row_widgets[idx] = [lbl_idx, lbl_item, lbl_desc]
        
        # Sugestão inteligente
        sugestao = self._sugerir_nivel(val_item)
        self.decisoes[idx] = sugestao
        
        # Segmented Button
        seg = ctk.CTkSegmentedButton(
            self,
            values=["N1", "N2", "N3", "ITEM"],
            command=lambda v, i=idx: self._atualizar_nivel(i, v),
            width=290
        )
        seg.set(sugestao)
        seg.grid(row=linha_visual, column=3, padx=5, pady=2)
        
        self._aplicar_estilo(idx, sugestao)
    
    def _sugerir_nivel(self, val_item: str) -> str:
        """Sugere nível baseado no padrão do item"""
        clean = val_item.strip().replace('.0', '')
        clean = ''.join(c for c in clean if c.isdigit() or c == '.')
        
        if not clean:
            return "ITEM"
        
        num_pontos = clean.count(".")
        
        if num_pontos == 0 and clean.isdigit():
            return "N1"
        elif num_pontos == 1:
            return "N2"
        elif num_pontos == 2:
            return "N3"
        else:
            return "ITEM"
    
    def _atualizar_nivel(self, index: int, valor: str):
        """Atualiza nível"""
        self.decisoes[index] = valor
        self._aplicar_estilo(index, valor)
    
    def _aplicar_estilo(self, index: int, nivel: str):
        """Aplica estilo visual"""
        if index not in self.row_widgets:
            return
        
        config = self.CORES_NIVEL.get(nivel, self.CORES_NIVEL["ITEM"])
        
        for widget in self.row_widgets[index]:
            widget.configure(text_color=config["texto"])
    
    def _limpar_tabela(self):
        """Limpa tabela"""
        for widget in self.winfo_children():
            info = widget.grid_info()
            if info and info.get('row', 0) > 0:
                widget.destroy()
        
        self.decisoes.clear()
        self.row_widgets.clear()
    
    def get_mapa_niveis(self) -> dict:
        """Retorna mapa de níveis"""
        return self.decisoes.copy()


class SisorcApp(ctk.CTk):
    """Aplicação principal CORRIGIDA"""
    
    def __init__(self):
        super().__init__()
        self.title("🏗️ SISORC ULTIMATE v13.0 - Professional Edition")
        self.geometry("1300x950")
        
        # Estado
        self.sintetico_path = ""
        self.modelo_path = ""
        self.processando = False
        
        # Config
        self._carregar_config()
        
        # Logger
        self.logger = Logger("SISORC")
        
        # UI
        self._criar_interface()
        self._carregar_historico()
        
        Logger.titulo("SISORC ULTIMATE v13.0 INICIADO")
    
    def _carregar_config(self):
        """Carrega configurações"""
        try:
            self.config = ConfigLoader.carregar('config/settings.json')
        except:
            self.config = {
                'sanitizacao': {'timeout': 30, 'limpar_temp': True},
                'historico': {},
                'excel': {},
                'database': {'nome_arquivo': 'sisorc.db'}
            }
    
    def _criar_interface(self):
        """Cria interface"""
        # Container principal com 2 colunas
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Coluna esquerda (controles)
        left_frame = ctk.CTkScrollableFrame(main_frame, width=900)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        # Coluna direita (log)
        right_frame = ctk.CTkFrame(main_frame, width=350)
        right_frame.pack(side="right", fill="both", padx=(5, 0))
        
        # Cria seções
        self._criar_secao_arquivos(left_frame)
        self._criar_separador(left_frame)
        self._criar_secao_parametros(left_frame)
        self._criar_separador(left_frame)
        self._criar_secao_tabela(left_frame)
        self._criar_separador(left_frame)
        self._criar_secao_projeto(left_frame)
        self._criar_botao_executar(left_frame)
        
        # Log
        self._criar_secao_log(right_frame)
    
    def _criar_separador(self, parent):
        """Cria linha separadora"""
        ctk.CTkFrame(parent, height=2, fg_color="gray30").pack(fill="x", pady=12)
    
    def _criar_secao_arquivos(self, parent):
        """Seção de arquivos"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(
            frame,
            text="📁 1. Seleção de Arquivos",
            font=("Arial", 16, "bold")
        ).pack(anchor="w", padx=5, pady=(5, 10))
        
        # Sintético
        f_sint = ctk.CTkFrame(frame, fg_color="transparent")
        f_sint.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkButton(
            f_sint,
            text="📊 Sintético (Dados)",
            command=self.selecionar_sintetico,
            width=180,
            height=38
        ).pack(side="left", padx=3)
        
        self.lbl_sintetico = ctk.CTkLabel(
            f_sint,
            text="Nenhum arquivo selecionado",
            text_color="gray50",
            anchor="w"
        )
        self.lbl_sintetico.pack(side="left", fill="x", expand=True, padx=10)
        
        # Modelo
        f_mod = ctk.CTkFrame(frame, fg_color="transparent")
        f_mod.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkButton(
            f_mod,
            text="📄 Modelo (Template)",
            command=self.selecionar_modelo,
            width=180,
            height=38,
            fg_color="#555555",
            hover_color="#666666"
        ).pack(side="left", padx=3)
        
        self.lbl_modelo = ctk.CTkLabel(
            f_mod,
            text="Nenhum arquivo selecionado",
            text_color="gray50",
            anchor="w"
        )
        self.lbl_modelo.pack(side="left", fill="x", expand=True, padx=10)
    
    def _criar_secao_parametros(self, parent):
        """Seção de parâmetros"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(
            frame,
            text="⚙️ 2. Parâmetros de Leitura",
            font=("Arial", 16, "bold")
        ).pack(anchor="w", padx=5, pady=(5, 10))
        
        params = ctk.CTkFrame(frame, fg_color="transparent")
        params.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkLabel(params, text="Linha Inicial:").pack(side="left", padx=5)
        self.ent_linha_inicial = ctk.CTkEntry(params, width=80)
        self.ent_linha_inicial.insert(0, "5")
        self.ent_linha_inicial.pack(side="left", padx=5)
        
        ctk.CTkLabel(params, text="Qtd. Linhas:").pack(side="left", padx=(15, 5))
        self.ent_qtd_linhas = ctk.CTkEntry(params, width=80)
        self.ent_qtd_linhas.insert(0, "100")
        self.ent_qtd_linhas.pack(side="left", padx=5)
        
        ctk.CTkButton(
            params,
            text="🔄 Carregar Tabela",
            command=self.carregar_preview,
            width=160,
            height=35,
            fg_color="#2563eb",
            hover_color="#1e40af"
        ).pack(side="left", padx=(25, 5))
    
    def _criar_secao_tabela(self, parent):
        """Seção da tabela"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="both", expand=True, pady=5)
        
        ctk.CTkLabel(
            frame,
            text="📋 3. Definição de Níveis",
            font=("Arial", 16, "bold")
        ).pack(anchor="w", padx=5, pady=(5, 10))
        
        self.tabela = LevelSelector(frame, height=280)
        self.tabela.pack(fill="both", expand=True, padx=5, pady=5)
    
    def _criar_secao_projeto(self, parent):
        """Seção de dados do projeto"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=5)
        
        ctk.CTkLabel(
            frame,
            text="🏗️ 4. Dados do Projeto",
            font=("Arial", 16, "bold")
        ).pack(anchor="w", padx=5, pady=(5, 10))
        
        grid = ctk.CTkFrame(frame, fg_color="transparent")
        grid.pack(padx=5, pady=5)
        
        # Obra
        ctk.CTkLabel(grid, text="Nome da Obra:", width=110, anchor="e").grid(
            row=0, column=0, padx=5, pady=5, sticky="e"
        )
        self.ent_obra = ctk.CTkEntry(grid, width=330, height=35)
        self.ent_obra.grid(row=0, column=1, padx=5, pady=5)
        
        # Local
        ctk.CTkLabel(grid, text="Local:", width=110, anchor="e").grid(
            row=0, column=2, padx=5, pady=5, sticky="e"
        )
        self.ent_local = ctk.CTkEntry(grid, width=230, height=35)
        self.ent_local.grid(row=0, column=3, padx=5, pady=5)
        
        # BDI
        ctk.CTkLabel(grid, text="BDI (%):", width=110, anchor="e").grid(
            row=1, column=0, padx=5, pady=5, sticky="e"
        )
        self.ent_bdi = ctk.CTkEntry(grid, width=100, height=35)
        self.ent_bdi.insert(0, "25.00")
        self.ent_bdi.grid(row=1, column=1, padx=5, pady=5, sticky="w")
    
    def _criar_botao_executar(self, parent):
        """Botão de execução"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=15)
        
        self.btn_executar = ctk.CTkButton(
            frame,
            text="🚀 GERAR ORÇAMENTO",
            command=self.executar,
            height=55,
            font=("Arial", 17, "bold"),
            fg_color="#16a34a",
            hover_color="#15803d"
        )
        self.btn_executar.pack(fill="x", padx=5)
    
    def _criar_secao_log(self, parent):
        """Seção de log"""
        ctk.CTkLabel(
            parent,
            text="📝 Log de Processamento",
            font=("Arial", 15, "bold")
        ).pack(anchor="w", padx=10, pady=(10, 5))
        
        # Log viewer
        self.log_viewer = LogViewer(parent, height=800, font=("Consolas", 11))
        self.log_viewer.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Conecta logger
        self.logger.adicionar_callback(self._callback_log)
        
        # Botão limpar
        ctk.CTkButton(
            parent,
            text="🗑️ Limpar Log",
            command=self.log_viewer.limpar,
            width=120,
            height=30,
            fg_color="gray40"
        ).pack(pady=(0, 10))
    
    def _callback_log(self, nivel, mensagem):
        """Callback para receber logs"""
        self.log_viewer.adicionar_log(nivel, mensagem)
    
    def selecionar_sintetico(self):
        """Seleciona sintético"""
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo SINTÉTICO",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]
        )
        
        if caminho:
            valido, msg = FileValidator.validar_excel(caminho)
            if not valido:
                messagebox.showerror("Arquivo Inválido", f"❌ {msg}")
                return
            
            self.sintetico_path = caminho
            nome = Path(caminho).name
            self.lbl_sintetico.configure(text=f"✓ {nome}", text_color="#10b981")
            Logger.info(f"Sintético: {nome}")
            self._salvar_historico()
    
    def selecionar_modelo(self):
        """Seleciona modelo"""
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo MODELO",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]
        )
        
        if caminho:
            valido, msg = FileValidator.validar_excel(caminho)
            if not valido:
                messagebox.showerror("Arquivo Inválido", f"❌ {msg}")
                return
            
            self.modelo_path = caminho
            nome = Path(caminho).name
            self.lbl_modelo.configure(text=f"✓ {nome}", text_color="#10b981")
            Logger.info(f"Modelo: {nome}")
            self._salvar_historico()
    
    def carregar_preview(self):
        """Carrega preview"""
        if not self.sintetico_path:
            messagebox.showwarning("Aviso", "⚠️ Selecione o sintético primeiro!")
            return
        
        try:
            linha_ini = int(self.ent_linha_inicial.get())
            qtd = int(self.ent_qtd_linhas.get())
            
            if linha_ini < 1 or qtd < 1:
                raise ValueError("Valores devem ser maiores que zero")
            
        except ValueError as e:
            messagebox.showerror("Erro", f"❌ Valores inválidos: {e}")
            return
        
        Logger.titulo("CARREGANDO PREVIEW")
        
        try:
            # Sanitiza
            sanitizer = ExcelSanitizer(self.config.get('sanitizacao', {}))
            sucesso, resultado, linha_header = sanitizer.sanitizar_arquivo(self.sintetico_path)
            
            if not sucesso:
                messagebox.showerror("Erro", f"❌ {resultado}")
                return
            
            # Lê dados
            skip = max(0, linha_ini - 2)
            df = pd.read_excel(resultado, header=skip, nrows=qtd)
            df.columns = df.columns.str.strip().str.upper()
            df = df.dropna(how='all')
            
            if len(df) == 0:
                messagebox.showwarning("Aviso", "⚠️ Nenhum dado encontrado!")
                return
            
            # Carrega tabela
            self.tabela.carregar_dados(df)
            
            # Limpa temp
            sanitizer.limpar_arquivos_temp()
            
            messagebox.showinfo("Sucesso", f"✅ {len(df)} linhas carregadas!")
            
        except Exception as e:
            Logger.error(f"Erro: {e}")
            messagebox.showerror("Erro", f"❌ {str(e)}")
    
    def executar(self):
        """Executa processamento"""
        if self.processando:
            messagebox.showwarning("Aviso", "⚠️ Já está processando!")
            return
        
        # Validações
        if not self.sintetico_path or not self.modelo_path:
            messagebox.showwarning("Aviso", "⚠️ Selecione ambos os arquivos!")
            return
        
        if not self.tabela.decisoes:
            messagebox.showwarning("Aviso", "⚠️ Carregue a tabela primeiro!")
            return
        
        # Valida dados
        nome_obra = self.ent_obra.get().strip()
        valido, msg = DataValidator.validar_nome_obra(nome_obra)
        if not valido:
            messagebox.showerror("Dados Inválidos", f"❌ {msg}")
            return
        
        bdi_str = self.ent_bdi.get().strip()
        valido, bdi_num, msg = DataValidator.validar_bdi(bdi_str)
        if not valido:
            messagebox.showerror("BDI Inválido", f"❌ {msg}")
            return
        
        # Prepara dados
        dados_projeto = {
            'obra': nome_obra,
            'local': self.ent_local.get().strip(),
            'bdi': bdi_num
        }
        
        mapa_niveis = self.tabela.get_mapa_niveis()
        
        try:
            linha_ini = int(self.ent_linha_inicial.get())
            qtd = int(self.ent_qtd_linhas.get())
            intervalo = (linha_ini, linha_ini + qtd - 1)
        except:
            messagebox.showerror("Erro", "❌ Valores de linha inválidos")
            return
        
        # Desabilita botão
        self.processando = True
        self.btn_executar.configure(
            state="disabled",
            text="⏳ Processando...",
            fg_color="gray50"
        )
        
        Logger.titulo("INICIANDO PROCESSAMENTO")
        
        # Thread
        thread = threading.Thread(
            target=self._processar_backend,
            args=(dados_projeto, mapa_niveis, intervalo),
            daemon=True
        )
        thread.start()
    
    def _processar_backend(self, dados, mapa, intervalo):
        """Processa em background"""
        try:
            # Sanitiza
            sanitizer = ExcelSanitizer(self.config.get('sanitizacao', {}))
            sucesso, arquivo_limpo, _ = sanitizer.sanitizar_arquivo(self.sintetico_path)
            
            if not sucesso:
                self.after(0, lambda: self._finalizar(False, arquivo_limpo))
                return
            
            # Processa
            engine = OrcamentoEngine(self.config)
            sucesso, resultado, info = engine.processar_orcamento(
                arquivo_limpo,
                self.modelo_path,
                dados,
                mapa,
                intervalo
            )
            
            # Limpa
            sanitizer.limpar_arquivos_temp()
            
            # Salva no banco
            if sucesso:
                try:
                    db = DatabaseManager(self.config)
                    db.inserir_orcamento({
                        'data_geracao': datetime.now().isoformat(),
                        'nome_obra': dados['obra'],
                        'local': dados['local'],
                        'bdi': dados['bdi'],
                        'valor_total': 0,
                        'arquivo_saida': resultado,
                        'num_itens': info.get('itens', 0),
                        'num_titulos': 0,
                        'duracao_processamento': 0
                    })
                except:
                    pass
            
            # Finaliza
            self.after(0, lambda: self._finalizar(sucesso, resultado))
            
        except Exception as e:
            Logger.error(f"Erro crítico: {e}")
            self.after(0, lambda: self._finalizar(False, str(e)))
    
    def _finalizar(self, sucesso: bool, mensagem: str):
        """Finaliza processamento"""
        self.processando = False
        self.btn_executar.configure(
            state="normal",
            text="🚀 GERAR ORÇAMENTO",
            fg_color="#16a34a"
        )
        
        if sucesso:
            Logger.titulo("✅ PROCESSAMENTO CONCLUÍDO")
            
            resposta = messagebox.askyesno(
                "✅ Sucesso!",
                f"Orçamento gerado!\n\n📁 {mensagem}\n\nAbrir arquivo?"
            )
            
            if resposta:
                import os
                try:
                    os.startfile(mensagem)
                except:
                    pass
        else:
            Logger.error(f"Falha: {mensagem}")
            messagebox.showerror("❌ Erro", f"Falha:\n\n{mensagem}")
    
    def _carregar_historico(self):
        """Carrega histórico"""
        hist = self.config.get('historico', {})
        
        sint = hist.get('ultimo_sintetico', '')
        if sint and Path(sint).exists():
            self.sintetico_path = sint
            self.lbl_sintetico.configure(
                text=f"✓ {Path(sint).name}",
                text_color="#10b981"
            )
        
        mod = hist.get('ultimo_modelo', '')
        if mod and Path(mod).exists():
            self.modelo_path = mod
            self.lbl_modelo.configure(
                text=f"✓ {Path(mod).name}",
                text_color="#10b981"
            )
    
    def _salvar_historico(self):
        """Salva histórico"""
        try:
            if 'historico' not in self.config:
                self.config['historico'] = {}
            
            self.config['historico']['ultimo_sintetico'] = self.sintetico_path
            self.config['historico']['ultimo_modelo'] = self.modelo_path
            
            ConfigLoader.salvar('config/settings.json', self.config)
        except Exception as e:
            Logger.warning(f"Erro ao salvar histórico: {e}")


if __name__ == "__main__":
    app = SisorcApp()
    app.mainloop()