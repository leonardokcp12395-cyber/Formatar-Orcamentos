import pandas as pd
import re
import os
from utils.logger import Logger

class ExcelSanitizer:
    def __init__(self, config):
        self.config = config

    def sanitizar_arquivo(self, input_path):
        """
        Sanitiza o arquivo Excel para leitura segura.
        Retorna (sucesso, caminho_arquivo_limpo, linha_cabecalho).
        """
        Logger.info(f"Iniciando sanitização de: {input_path}")

        try:
            # Neste exemplo simples, apenas verificamos se o arquivo existe e retornamos ele mesmo
            # Em uma implementação real, trataríamos células mescladas, etc.
            if not os.path.exists(input_path):
                return False, "Arquivo não encontrado", 0

            # Tentar ler para ver se não está corrompido
            try:
                pd.read_excel(input_path, nrows=5)
            except Exception as e:
                return False, f"Erro ao ler arquivo: {e}", 0

            # Detectar linha de cabeçalho
            # Vamos ler as primeiras 50 linhas
            df_head = pd.read_excel(input_path, header=None, nrows=50)
            linha_header = 0

            for idx, row in df_head.iterrows():
                row_str = str(row.values).upper()
                if 'ITEM' in row_str and 'DESCRI' in row_str:
                    linha_header = idx
                    break

            # Se precisarmos gerar um arquivo limpo temporário (ex: unmerge cells)
            # aqui seria o lugar. Como estamos simplificando:
            return True, input_path, linha_header

        except Exception as e:
            Logger.error(f"Erro na sanitização: {str(e)}")
            return False, str(e), 0

    def limpar_arquivos_temp(self):
        """Remove arquivos temporários gerados"""
        pass
