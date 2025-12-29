import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from datetime import datetime
import os
import shutil
from utils.logger import Logger

class OrcamentoEngine:
    def __init__(self, config):
        self.config = config
        # Tenta pegar diretório de output da config, senão usa padrão
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
        Processa o orçamento e gera o arquivo final.
        """
        Logger.titulo("INICIANDO ENGINE DE ORÇAMENTO")

        try:
            linha_ini, linha_fim = intervalo_linhas
            Logger.info(f"Processando intervalo: Linhas {linha_ini} a {linha_fim}")

            # 1. Detectar cabeçalho e colunas (lendo primeiras 50 linhas)
            # Precisamos saber onde estão as colunas ITEM, DESCRIÇÃO, etc.
            df_head = pd.read_excel(arquivo_limpo, header=None, nrows=50)

            idx_item, idx_desc, idx_unid, idx_quant, idx_unit, idx_cod = None, None, None, None, None, None
            header_row_idx = -1

            # Procurar linha de cabeçalho
            for r in range(len(df_head)):
                row_vals = [str(v).upper() for v in df_head.iloc[r].values]
                if any("ITEM" in v for v in row_vals) and any("DESCRI" in v for v in row_vals):
                    header_row_idx = r
                    # Mapear índices das colunas nessa linha
                    for c, val in enumerate(row_vals):
                        if "ITEM" in val or "IT" in val: idx_item = c
                        elif "DESCRI" in val or "DISCRIMINA" in val: idx_desc = c
                        elif "UNID" in val or "UND" in val: idx_unid = c
                        elif "QUANT" in val or "QTD" in val: idx_quant = c
                        elif "UNIT" in val or "PREÇO" in val: idx_unit = c
                        elif "COD" in val: idx_cod = c
                    break

            if header_row_idx == -1:
                 # Fallback: Tenta usar a linha anterior ao inicio dos dados como cabeçalho
                 fallback_header = max(0, linha_ini - 2)
                 Logger.warning(f"Cabeçalho não detectado automaticamente. Tentando linha {fallback_header + 1}")
                 # Se o intervalo começa muito longe, precisamos ler especificamente essa região
                 if fallback_header > 50:
                     df_fallback = pd.read_excel(arquivo_limpo, header=None, skiprows=fallback_header, nrows=1)
                     # Lógica de mapeamento similar ou assumir posicional
                     pass

            # Se ainda não achou colunas, usa posicional padrão
            if idx_item is None: idx_item = 0
            if idx_desc is None: idx_desc = 1
            # Outros podem ser None

            Logger.info(f"Colunas mapeadas (índices): Item={idx_item}, Desc={idx_desc}")

            # 2. Ler os dados do intervalo
            # skiprows = linha_ini - 1 (porque Excel é 1-based e skiprows é 0-based quantidade de linhas para pular antes)
            # Mas cuidado: skiprows=0 lê a linha 1. skiprows=4 lê a linha 5.
            # Então skiprows = linha_ini - 1 está correto para começar a ler na linha_ini.

            skip = linha_ini - 1
            qtd = linha_fim - linha_ini + 1

            df_dados = pd.read_excel(arquivo_limpo, header=None, skiprows=skip, nrows=qtd)
            df_dados = df_dados.dropna(how='all')

            if df_dados.empty:
                 return False, "Nenhum dado encontrado no intervalo.", {}

            # 3. Preparar arquivo de saída
            filename = f"ORC_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_path = os.path.join(self.output_dir, filename)
            shutil.copy(modelo_path, save_path)

            wb = openpyxl.load_workbook(save_path)
            ws = wb.active

            # 4. Preencher cabeçalho
            cells_map = {
                'C8': dados_projeto.get('obra', ''),
                'C9': dados_projeto.get('local', ''),
                'H10': dados_projeto.get('email', ''),
                'C17': datetime.now().strftime("%d/%m/%Y")
            }

            for cell_coord, valor in cells_map.items():
                try: ws[cell_coord] = valor
                except: pass

            # 5. Inserir dados
            linha_base = self.config.get('excel', {}).get('linha_inicial_modelo', 25)
            current_row = linha_base

            # O mapa_niveis usa índices.
            # Se a leitura dos dados foi feita exatamente no intervalo, o índice 0 do df_dados corresponde à linha_ini.
            # Precisamos ajustar o acesso ao mapa_niveis se ele usar índices absolutos ou relativos.
            # A UI (ui/main_window.py) gera o mapa iterando sobre o df que ela leu com `header=skip`.
            # Se a UI lê com header, o índice 0 é a primeira linha de DADOS.
            # Aqui estamos lendo sem header, então o índice 0 também é a primeira linha de DADOS.
            # Os índices devem bater.

            itens_processados = 0

            for idx, row in df_dados.iterrows():
                # idx reinicia em 0? Sim, read_excel sem index_col gera RangeIndex.
                # O mapa_niveis da UI foi gerado sobre um df que também provavelmente tinha RangeIndex resetado (se leu apenas o pedaço).
                # Assumindo que sim.

                nivel = mapa_niveis.get(idx, 'ITEM')

                ws.insert_rows(current_row)

                # Extrair valores usando índices de coluna mapeados
                val_item = str(row.iloc[idx_item]) if idx_item is not None and idx_item < len(row) and pd.notna(row.iloc[idx_item]) else ""
                val_desc = str(row.iloc[idx_desc]) if idx_desc is not None and idx_desc < len(row) and pd.notna(row.iloc[idx_desc]) else ""
                val_cod = str(row.iloc[idx_cod]) if idx_cod is not None and idx_cod < len(row) and pd.notna(row.iloc[idx_cod]) else ""
                val_unid = str(row.iloc[idx_unid]) if idx_unid is not None and idx_unid < len(row) and pd.notna(row.iloc[idx_unid]) else ""

                val_quant = None
                if idx_quant is not None and idx_quant < len(row):
                    val_quant = self._clean_float(row.iloc[idx_quant])

                val_unit = None
                if idx_unit is not None and idx_unit < len(row):
                    val_unit = self._clean_float(row.iloc[idx_unit])

                # Escrever
                ws.cell(row=current_row, column=1, value=val_item)
                ws.cell(row=current_row, column=2, value=val_cod)
                ws.cell(row=current_row, column=4, value=val_desc)

                if nivel == 'ITEM':
                    ws.cell(row=current_row, column=5, value=val_unid)

                    if val_quant is not None:
                        ws.cell(row=current_row, column=6, value=val_quant)
                        ws.cell(row=current_row, column=6).number_format = '#,##0.00'

                    if val_unit is not None:
                        ws.cell(row=current_row, column=7, value=val_unit)
                        ws.cell(row=current_row, column=7).number_format = '#,##0.00'

                    # Fórmula
                    ws.cell(row=current_row, column=8, value=f"=F{current_row}*G{current_row}")
                    ws.cell(row=current_row, column=8).number_format = '#,##0.00'

                self._aplicar_estilo(ws, current_row, nivel)
                current_row += 1
                itens_processados += 1

            # 6. Rodapé
            row_total = None
            for r in range(current_row, current_row + 20):
                val = ws.cell(row=r, column=1).value
                if val and isinstance(val, str) and "Total sem BDI" in val:
                    row_total = r
                    break

            if row_total:
                last = current_row - 1
                if last >= linha_base:
                    ws.cell(row=row_total, column=8).value = f"=SUM(H{linha_base}:H{last})"

                    # Tratar BDI com segurança (pode vir como string com vírgula)
                    raw_bdi = dados_projeto.get('bdi', 0)
                    try:
                        if isinstance(raw_bdi, str):
                            raw_bdi = raw_bdi.replace(',', '.')
                        bdi = float(raw_bdi) / 100
                    except:
                        bdi = 0.0

                    ws.cell(row=row_total+1, column=8).value = f"=H{row_total}*{bdi}"
                    ws.cell(row=row_total+2, column=8).value = f"=H{row_total}+H{row_total+1}"

            wb.save(save_path)
            return True, save_path, {'itens': itens_processados}

        except Exception as e:
            Logger.error(f"Erro processamento: {e}")
            return False, str(e), {}

    def _clean_float(self, val):
        if pd.isna(val): return None
        if isinstance(val, (int, float)): return val
        try:
            return float(str(val).replace('R$', '').replace('.', '').replace(',', '.').strip())
        except:
            return None

    def _aplicar_estilo(self, ws, row, nivel):
        for col in range(1, 9):
            cell = ws.cell(row=row, column=col)
            cell.border = self.border
            if col == 4: cell.alignment = self.align_left
            else: cell.alignment = self.align_center

            if nivel in self.cores:
                cell.fill = self.cores[nivel]
                cell.font = self.font_bold
            if nivel in ['N1', 'N2', 'N3']:
                cell.font = self.font_bold
