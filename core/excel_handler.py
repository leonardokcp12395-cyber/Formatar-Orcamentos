import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from datetime import datetime
import os
import shutil
import unicodedata
import re
from utils.logger import Logger

class OrcamentoEngine:
    def __init__(self, config):
        self.config = config
        self.output_dir = self.config.get('output', {}).get('dir', 'orcamentos_gerados')
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        # Definição de estilos
        self.cores = {
            'N1': PatternFill(start_color="9BC2E6", fill_type="solid"),
            'N2': PatternFill(start_color="BDD7EE", fill_type="solid"),
            'N3': PatternFill(start_color="DDEBF7", fill_type="solid"),
        }

        self.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        self.align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
        self.font_bold = Font(bold=True)

    def processar_orcamento(self, arquivo_limpo, modelo_path, dados_projeto, mapa_niveis, intervalo_linhas):
        """
        Processa o orçamento com detecção robusta de colunas.
        """
        Logger.titulo("INICIANDO ENGINE DE ORÇAMENTO (ENHANCED)")

        try:
            linha_ini, linha_fim = intervalo_linhas
            Logger.info(f"Intervalo de dados: Linhas {linha_ini} a {linha_fim}")

            # 1. Detectar cabeçalho e Mapear Colunas
            # Lê primeiras 100 linhas para achar o cabeçalho, independente de onde o usuário disse que começam os DADOS
            df_search = pd.read_excel(arquivo_limpo, header=None, nrows=100)

            col_map = {
                'item': -1, 'desc': -1, 'unid': -1,
                'quant': -1, 'unit': -1, 'cod': -1
            }

            header_row_idx = -1

            # Busca melhor linha de cabeçalho
            for r_idx, row in df_search.iterrows():
                row_vals = [self._normalize_str(v) for v in row.values]
                if any("ITEM" in v for v in row_vals) and \
                   (any("DESCRI" in v for v in row_vals) or any("DISCRIMINA" in v for v in row_vals)):
                    header_row_idx = r_idx
                    # Mapeia colunas nesta linha
                    for c_idx, val in enumerate(row_vals):
                        if self._match_col(val, ['ITEM', 'IT', 'NUM']): col_map['item'] = c_idx
                        elif self._match_col(val, ['DESCRI', 'DISCRIMINA', 'ESPECIFICA']): col_map['desc'] = c_idx
                        elif self._match_col(val, ['UNID', 'UND', 'UNIDADE']): col_map['unid'] = c_idx
                        elif self._match_col(val, ['QUANT', 'QTD', 'QTDE', 'METRAGEM']): col_map['quant'] = c_idx
                        elif self._match_col(val, ['UNIT', 'PRECO', 'VALOR UNIT']): col_map['unit'] = c_idx
                        elif self._match_col(val, ['COD', 'SINAPI', 'REFERENCIA']): col_map['cod'] = c_idx
                    break

            Logger.info(f"Cabeçalho encontrado na linha {header_row_idx}")
            Logger.info(f"Mapa de colunas: {col_map}")

            # Fallbacks
            if col_map['item'] == -1: col_map['item'] = 0
            if col_map['desc'] == -1: col_map['desc'] = 1

            # 2. Ler Dados
            # Lê apenas as linhas de dados especificadas pelo usuário
            # Ajuste de índice: skiprows é 0-based. Se dados começam na linha 5 (usuário), skiprows=4.
            skip = linha_ini - 1
            qtd = linha_fim - linha_ini + 1

            df_dados = pd.read_excel(arquivo_limpo, header=None, skiprows=skip, nrows=qtd)
            df_dados = df_dados.dropna(how='all')

            if df_dados.empty:
                return False, "Nenhum dado encontrado no intervalo.", {}

            # 3. Preparar Saída
            filename = f"ORC_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_path = os.path.join(self.output_dir, filename)
            shutil.copy(modelo_path, save_path)

            wb = openpyxl.load_workbook(save_path)
            ws = wb.active

            # 4. Preencher Cabeçalho (Fixo)
            self._preencher_cabecalho_modelo(ws, dados_projeto)

            # 5. Inserir Itens
            linha_base = self.config.get('excel', {}).get('linha_inicial_modelo', 25)
            current_row = linha_base
            itens_processados = 0

            for idx, row in df_dados.iterrows():
                # Tenta pegar nível do mapa. Se não tiver, tenta inferir.
                nivel = mapa_niveis.get(idx, None)

                # Extração de valores segura
                val_item = self._get_val(row, col_map['item'])
                val_desc = self._get_val(row, col_map['desc'])

                # Se nível não veio do mapa (ex: processamento batch), infere
                if not nivel:
                    nivel = self._inferir_nivel(val_item)

                ws.insert_rows(current_row)

                ws.cell(row=current_row, column=1, value=val_item)
                ws.cell(row=current_row, column=2, value=self._get_val(row, col_map['cod']))
                ws.cell(row=current_row, column=4, value=val_desc)

                if nivel == 'ITEM':
                    ws.cell(row=current_row, column=5, value=self._get_val(row, col_map['unid']))

                    val_quant = self._get_float(row, col_map['quant'])
                    if val_quant is not None:
                        ws.cell(row=current_row, column=6, value=val_quant)
                        ws.cell(row=current_row, column=6).number_format = '#,##0.00'

                    val_unit = self._get_float(row, col_map['unit'])
                    if val_unit is not None:
                        ws.cell(row=current_row, column=7, value=val_unit)
                        ws.cell(row=current_row, column=7).number_format = '#,##0.00'

                    ws.cell(row=current_row, column=8, value=f"=F{current_row}*G{current_row}")
                    ws.cell(row=current_row, column=8).number_format = '#,##0.00'

                self._aplicar_estilo(ws, current_row, nivel)
                current_row += 1
                itens_processados += 1

            # 6. Rodapé
            self._ajustar_rodape(ws, current_row, linha_base, dados_projeto)

            wb.save(save_path)
            return True, save_path, {'itens': itens_processados}

        except Exception as e:
            Logger.error(f"Erro processamento: {e}")
            return False, str(e), {}

    def _preencher_cabecalho_modelo(self, ws, dados):
        try:
            if dados.get('obra'): ws['C8'] = dados['obra']
            if dados.get('local'): ws['C9'] = dados['local']
            if dados.get('email'): ws['H10'] = dados['email']
            ws['C17'] = datetime.now().strftime("%d/%m/%Y")
        except: pass

    def _ajustar_rodape(self, ws, current_row, linha_base, dados):
        # Busca onde foi parar o rodapé
        row_total = None
        for r in range(current_row, current_row + 50): # Busca expandida
            val = ws.cell(row=r, column=1).value
            if val and isinstance(val, str) and "Total sem BDI" in val:
                row_total = r
                break

        if row_total:
            last = current_row - 1
            if last >= linha_base:
                ws.cell(row=row_total, column=8).value = f"=SUM(H{linha_base}:H{last})"

                # BDI Seguro
                try:
                    raw_bdi = str(dados.get('bdi', 0)).replace(',', '.')
                    bdi = float(raw_bdi) / 100
                except:
                    bdi = 0.0

                ws.cell(row=row_total+1, column=8).value = f"=H{row_total}*{bdi}"
                ws.cell(row=row_total+2, column=8).value = f"=H{row_total}+H{row_total+1}"

    def _match_col(self, val, keywords):
        val = str(val).upper()
        return any(k in val for k in keywords)

    def _normalize_str(self, text):
        if pd.isna(text): return ""
        text = str(text).upper().strip()
        return ''.join(c for c in unicodedata.normalize('NFD', text)
                      if unicodedata.category(c) != 'Mn')

    def _get_val(self, row, col_idx):
        if col_idx == -1 or col_idx >= len(row): return ""
        val = row.iloc[col_idx]
        if pd.isna(val): return ""
        return str(val).strip()

    def _get_float(self, row, col_idx):
        if col_idx == -1 or col_idx >= len(row): return None
        val = row.iloc[col_idx]
        if pd.isna(val): return None
        if isinstance(val, (int, float)): return val
        try:
            # Limpeza agressiva: remove R$, espaços, e converte vírgula
            clean = str(val).replace('R$', '').replace(' ', '')
            clean = clean.replace('.', '').replace(',', '.') # Assume milhar=ponto, decimal=virgula
            return float(clean)
        except:
            return None

    def _inferir_nivel(self, val_item):
        """Inferência simples baseada em pontos 1.1.1"""
        if not val_item: return 'ITEM'
        dots = val_item.count('.')
        if dots == 0: return 'N1'
        if dots == 1: return 'N2'
        if dots == 2: return 'N3'
        return 'ITEM'

    def _aplicar_estilo(self, ws, row, nivel):
        for col in range(1, 9):
            cell = ws.cell(row=row, column=col)
            cell.border = self.border
            cell.alignment = self.align_left if col == 4 else self.align_center

            if nivel in self.cores:
                cell.fill = self.cores[nivel]
                cell.font = self.font_bold
            if nivel in ['N1', 'N2', 'N3']:
                cell.font = self.font_bold
