import pandas as pd
import openpyxl
from openpyxl.styles import Border, Side, Alignment
from datetime import datetime
import os
import shutil
from logger import setup_logger

logger = setup_logger()

class ExcelHandler:
    def __init__(self, output_dir):
        self.output_dir = output_dir
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

    def _aplicar_estilo_linha(self, ws, row, col_inicial, col_final):
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))
        align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

        for col in range(col_inicial, col_final + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            if col == 4: # Descrição
                cell.alignment = align_left
            else:
                cell.alignment = align_center

    def processar_modelo_insert(self, modelo_path, df_dados, info_cabecalho=None):
        logger.info(">>> INICIANDO GERAÇÃO DO ORÇAMENTO <<<")
        try:
            filename = f"ORC_Final_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            save_path = os.path.join(self.output_dir, filename)
            shutil.copy(modelo_path, save_path)
            
            wb = openpyxl.load_workbook(save_path)
            ws = wb.active

            # Preencher Cabeçalho
            if info_cabecalho:
                # Mapeamento de Células (Ajuste conforme seu modelo Excel)
                mapa = {'setor': 'C8', 'telefone': 'H9', 'servidor': 'C9', 'email': 'H10'}
                ws['C17'] = datetime.now().strftime("%d/%m/%Y")
                for key, cell in mapa.items():
                    if info_cabecalho.get(key):
                        ws[cell] = info_cabecalho[key]

            # Inserir Itens
            LINHA_INICIAL = 25
            current_row = LINHA_INICIAL
            
            col_indices = {'ITEM': 1, 'CÓDIGO': 2, 'BANCO': 3, 'DESCRIÇÃO': 4, 'UND': 5, 'QUANT.': 6, 'VALOR UNIT': 7}

            for index, row in df_dados.iterrows():
                ws.insert_rows(current_row)
                for col_name, col_idx in col_indices.items():
                    val = row.get(col_name, '')
                    # Conversão numérica segura
                    if col_name in ['QUANT.', 'VALOR UNIT']:
                        try:
                            if isinstance(val, str): val = float(val.replace('.', '').replace(',', '.'))
                            else: val = float(val)
                        except: pass
                    ws.cell(row=current_row, column=col_idx).value = val

                # Fórmula
                if row.get('QUANT.') and row.get('VALOR UNIT'):
                    ws.cell(row=current_row, column=8).value = f"=F{current_row}*G{current_row}"
                    ws.cell(row=current_row, column=8).number_format = '#,##0.00'
                
                self._aplicar_estilo_linha(ws, current_row, 1, 8)
                current_row += 1

            # Arrumar Rodapé (Procura onde foi parar)
            row_total = None
            for r in range(current_row, ws.max_row + 1):
                if ws.cell(row=r, column=1).value and "Total sem BDI" in str(ws.cell(row=r, column=1).value):
                    row_total = r
                    break
            
            if row_total:
                last_item = current_row - 1
                ws.cell(row=row_total, column=8).value = f"=SUM(H{LINHA_INICIAL}:H{last_item})"
                ws.cell(row=row_total+1, column=8).value = f"=H{row_total}*0.2882" # BDI
                ws.cell(row=row_total+2, column=8).value = f"=H{row_total}+H{row_total+1}" # Geral

            wb.save(save_path)
            return str(save_path)

        except Exception as e:
            logger.error(f"Erro no Handler: {e}")
            raise e