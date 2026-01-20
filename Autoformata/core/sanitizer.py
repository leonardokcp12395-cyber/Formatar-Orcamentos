import pandas as pd
import os
from utils.logger import Logger

class ExcelSanitizer:
    def __init__(self, config=None):
        self.config = config or {}

    def sanitizar_arquivo(self, input_path):
        """
        Localiza automaticamente onde os dados começam.
        Retorna: (Sucesso, CaminhoArquivo, LinhaDoCabecalho)
        """
        Logger.info(f"Analisando estrutura do arquivo: {input_path}")
        
        try:
            if not os.path.exists(input_path):
                return False, "Arquivo não encontrado", 0

            # 1. Varredura Inteligente
            # Lê as primeiras 50 linhas sem cabeçalho
            df_temp = pd.read_excel(input_path, header=None, nrows=50)
            
            linha_cabecalho = -1
            
            # Procura por linha que tenha palavras-chave de orçamentos
            for idx, row in df_temp.iterrows():
                linha_str = row.astype(str).str.upper().values
                
                # Critério: Tem que ter "ITEM" e ("DESCRIÇÃO" ou "DISCRIMINAÇÃO")
                tem_item = any('ITEM' in str(x) for x in linha_str)
                tem_desc = any('DESCRI' in str(x) or 'DISCRIM' in str(x) or 'SERVIÇO' in str(x) for x in linha_str)
                
                if tem_item and tem_desc:
                    linha_cabecalho = idx # Pandas usa base-0
                    Logger.info(f"Cabeçalho detectado na linha Excel: {idx + 1}")
                    break
            
            # Se não achou, assume 0 (primeira linha) como fallback
            if linha_cabecalho == -1:
                Logger.warning("Cabeçalho não detectado explicitamente. Usando linha 1.")
                linha_cabecalho = 0

            # Retorna o caminho original e a linha descoberta
            return True, input_path, linha_cabecalho

        except Exception as e:
            Logger.error(f"Erro na sanitização: {str(e)}")
            return False, str(e), 0

    def limpar_arquivos_temp(self):
        pass