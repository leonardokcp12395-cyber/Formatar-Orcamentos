import sqlite3
from datetime import datetime
from utils.logger import Logger

class DatabaseManager:
    """Gerenciador de Banco de Dados SQLite"""
    
    def __init__(self, config: dict):
        # Pega o nome do arquivo corretamente do dicionário config
        self.db_name = config.get('database', {}).get('nome_arquivo', 'sisorc_history.db')
        self._init_db()

    def _init_db(self):
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS orcamentos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    data_geracao TEXT,
                    nome_obra TEXT,
                    local TEXT,
                    bdi REAL,
                    valor_total REAL,
                    arquivo_saida TEXT,
                    num_itens INTEGER,
                    num_titulos INTEGER,
                    duracao_processamento REAL
                )
            ''')
            conn.commit()
            conn.close()
        except Exception as e:
            Logger.error(f"Erro DB Init: {e}")

    def inserir_orcamento(self, dados: dict):
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO orcamentos 
                (data_geracao, nome_obra, local, bdi, valor_total, arquivo_saida, num_itens, num_titulos, duracao_processamento)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                dados['data_geracao'], dados['nome_obra'], dados['local'], 
                dados['bdi'], dados['valor_total'], dados['arquivo_saida'],
                dados['num_itens'], dados['num_titulos'], dados['duracao_processamento']
            ))
            conn.commit()
            conn.close()
        except Exception as e:
            Logger.error(f"Erro DB Insert: {e}")

    def buscar_estatisticas(self):
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            
            cursor.execute("SELECT COUNT(*), SUM(valor_total), AVG(num_itens) FROM orcamentos")
            total, valor, media = cursor.fetchone()
            
            cursor.execute("SELECT nome_obra FROM orcamentos ORDER BY id DESC LIMIT 1")
            ultimo = cursor.fetchone()
            
            conn.close()
            
            return {
                'total_orcamentos': total or 0,
                'valor_total_processado': valor or 0.0,
                'media_itens': media or 0.0,
                'ultimo_orcamento': {'nome': ultimo[0]} if ultimo else None
            }
        except Exception:
            return {'total_orcamentos': 0, 'valor_total_processado': 0, 'media_itens': 0, 'ultimo_orcamento': None}

    def buscar_orcamentos(self, limite=20):
        try:
            conn = sqlite3.connect(self.db_name)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM orcamentos ORDER BY id DESC LIMIT ?", (limite,))
            rows = [dict(row) for row in cursor.fetchall()]
            conn.close()
            return rows
        except:
            return []

class LogHandler:
    def __init__(self, db_manager):
        self.db = db_manager