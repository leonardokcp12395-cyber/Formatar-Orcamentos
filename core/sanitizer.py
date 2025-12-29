import pandas as pd
import re
import os
from utils.logger import Logger
import unicodedata

class ExcelSanitizer:
    def __init__(self, config):
        self.config = config

    def sanitizar_arquivo(self, input_path):
        """
        Sanitiza o arquivo Excel para leitura segura.
        Procura dinamicamente pelo cabeçalho da tabela.
        Retorna (sucesso, caminho_arquivo, linha_cabecalho).
        """
        Logger.info(f"Iniciando sanitização avançada de: {input_path}")

        try:
            if not os.path.exists(input_path):
                return False, "Arquivo não encontrado", 0

            # Ler um pedaço maior do arquivo para encontrar o cabeçalho
            # Arquivos do governo/sistemas podem ter cabeçalhos grandes
            df_search = pd.read_excel(input_path, header=None, nrows=100)

            header_row_idx = -1
            best_score = 0

            # Palavras-chave obrigatórias e opcionais para identificar a linha de títulos
            keywords = {
                'ITEM': 2,       # Peso 2
                'DESCRI': 2,     # Peso 2 (Descrição, Discriminação)
                'QUANT': 1,      # Peso 1
                'UNIT': 1,       # Peso 1 (Unitário)
                'TOTAL': 1,      # Peso 1
                'COD': 1         # Peso 1 (Código)
            }

            for idx, row in df_search.iterrows():
                row_text = " ".join([str(val).upper() for val in row.values])
                row_text = self._normalize_text(row_text)

                score = 0
                for kw, weight in keywords.items():
                    if kw in row_text:
                        score += weight

                # Exige pelo menos ITEM e DESCRIÇÃO para considerar válido
                if 'ITEM' in row_text and ('DESCRI' in row_text or 'DISCRIMINA' in row_text):
                     if score > best_score:
                        best_score = score
                        header_row_idx = idx

            if header_row_idx != -1:
                Logger.info(f"Cabeçalho detectado na linha {header_row_idx + 1} (Score: {best_score})")
                return True, input_path, header_row_idx
            else:
                Logger.warning("Cabeçalho não detectado automaticamente. Usando linha 0 como fallback.")
                return True, input_path, 0

        except Exception as e:
            Logger.error(f"Erro na sanitização: {str(e)}")
            return False, str(e), 0

    def _normalize_text(self, text):
        """Remove acentos e caracteres especiais para comparação"""
        return ''.join(c for c in unicodedata.normalize('NFD', text)
                      if unicodedata.category(c) != 'Mn')

    def limpar_arquivos_temp(self):
        """Remove arquivos temporários gerados"""
        pass
