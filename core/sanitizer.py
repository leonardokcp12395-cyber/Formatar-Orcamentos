import pandas as pd
import re
from logger import setup_logger

logger = setup_logger()

class DataSanitizer:
    def sanitize(self, input_path):
        logger.info(f"Iniciando sanitização de: {input_path}")
        
        try:
            # 1. Extração de Metadados (Cabeçalho)
            # Lê as primeiras 25 linhas sem cabeçalho para achar informações
            # header=None garante que leiamos tudo como dados brutos
            df_raw = pd.read_excel(input_path, header=None, nrows=25)
            header_info = self._extract_header_info(df_raw)
            logger.info(f"Dados de cabeçalho encontrados: {header_info}")

            # 2. Localização e Extração da Tabela de Itens
            df_full = pd.read_excel(input_path, header=None)
            
            # Procura a linha que contém "Item", "Código" e "Descrição"
            header_row_index = None
            for idx, row in df_full.iterrows():
                # Converte linha para string e uppercase para busca insensível a maiúsculas
                row_str = row.astype(str).str.upper().values
                if 'ITEM' in str(row_str) and 'DESCRIÇÃO' in str(row_str):
                    header_row_index = idx
                    break
            
            if header_row_index is None:
                # Fallback: Tenta achar linha 4 ou 5 se não achar por texto
                logger.warning("Cabeçalho não encontrado por texto. Tentando linha 4 padrão.")
                header_row_index = 4 # Ajuste conforme seu arquivo padrão se necessário

            # Recarrega o dataframe usando a linha correta como cabeçalho
            df_items = pd.read_excel(input_path, header=header_row_index)
            
            # Limpeza das colunas (remove colunas vazias e normaliza nomes)
            df_items = self._clean_table_columns(df_items)
            
            # Filtra apenas linhas onde a coluna ITEM parece um número (1, 1.1, 2, etc)
            # Isso remove linhas de rodapé ou lixo
            df_items = df_items[df_items['ITEM'].astype(str).str.match(r'^\d+(\.\d+)*$')]

            logger.info(f"Tabela de itens extraída com {len(df_items)} linhas.")
            
            return df_items, header_info

        except Exception as e:
            logger.error(f"Erro na sanitização: {str(e)}")
            raise e

    def _extract_header_info(self, df):
        """Varre o dataframe cru em busca de palavras-chave"""
        info = {
            'setor': '',
            'servidor': '',
            'telefone': '',
            'email': '',
            'os_num': '',
            'data_emissao': ''
        }

        # Converte tudo para string para facilitar busca
        df = df.fillna('')
        
        # Função auxiliar para buscar valor na célula vizinha à direita
        def get_neighbor(row_idx, col_idx):
            try:
                # Tenta pegar coluna +1 (vizinho direito)
                val = df.iloc[row_idx, col_idx+1]
                if val and str(val).strip():
                    return str(val).strip()
                # Se não, tenta coluna +2 (caso tenha uma celula vazia no meio)
                val2 = df.iloc[row_idx, col_idx+2]
                if val2 and str(val2).strip():
                    return str(val2).strip()
            except:
                pass
            return ''

        for r_idx, row in df.iterrows():
            for c_idx, cell in enumerate(row):
                val_str = str(cell).upper().strip()
                
                # Busca por palavras chave e pega o vizinho
                if 'SETOR:' in val_str:
                    clean_val = val_str.replace('SETOR:', '').strip()
                    info['setor'] = clean_val if clean_val else get_neighbor(r_idx, c_idx)
                
                if 'SERVIDOR:' in val_str:
                    clean_val = val_str.replace('SERVIDOR:', '').strip()
                    info['servidor'] = clean_val if clean_val else get_neighbor(r_idx, c_idx)
                
                if 'TELEFONE:' in val_str:
                    clean_val = val_str.replace('TELEFONE:', '').strip()
                    info['telefone'] = clean_val if clean_val else get_neighbor(r_idx, c_idx)

                if 'EMAIL:' in val_str:
                    clean_val = val_str.replace('EMAIL:', '').strip()
                    info['email'] = clean_val if clean_val else get_neighbor(r_idx, c_idx)

        return info

    def _clean_table_columns(self, df):
        """Padroniza nomes das colunas"""
        # Remove colunas totalmente vazias (Unnamed)
        df = df.dropna(how='all', axis=1)
        
        # Mapa de renomeação para garantir padrão
        cols_map = {}
        for col in df.columns:
            c_upper = str(col).upper().strip()
            if 'ITEM' in c_upper: cols_map[col] = 'ITEM'
            elif 'CÓDIGO' in c_upper or 'CODIGO' in c_upper: cols_map[col] = 'CÓDIGO'
            elif 'BANCO' in c_upper: cols_map[col] = 'BANCO'
            elif 'DESCRIÇÃO' in c_upper: cols_map[col] = 'DESCRIÇÃO'
            elif 'UND' in c_upper or 'UNID' in c_upper: cols_map[col] = 'UND'
            elif 'QUANT' in c_upper: cols_map[col] = 'QUANT.'
            elif 'UNIT' in c_upper: cols_map[col] = 'VALOR UNIT'
        
        df = df.rename(columns=cols_map)
        
        # Garante que as colunas essenciais existam
        required = ['ITEM', 'CÓDIGO', 'BANCO', 'DESCRIÇÃO', 'UND', 'QUANT.', 'VALOR UNIT']
        for req in required:
            if req not in df.columns:
                df[req] = '' # Cria vazia se não existir
                
        return df[required]