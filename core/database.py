import sqlite3
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional
from utils.logger import Logger
from core.exceptions import PlanifyError

class DatabaseManager:
    """Gerenciador de Banco de Dados SQLite"""

    def __init__(self, config: Dict[str, Any]):
        self.db_name: str = config.get('database', {}).get('nome_arquivo', 'planify_history.db')
        self._migrar_db_antigo()
        self._init_db()

    def _migrar_db_antigo(self) -> None:
        """Migra a base de dados antiga (sisorc_history.db) para o novo nome."""
        try:
            novo = Path(self.db_name)
            antigo = novo.parent / "sisorc_history.db"
            if antigo.exists() and not novo.exists():
                os.rename(str(antigo), str(novo))
                Logger.info("📦 Base de dados migrada: sisorc_history.db → planify_history.db")
        except Exception as e:
            Logger.warning(f"Aviso na migração do DB: {e}")

    def _init_db(self) -> None:
        """Inicializa o banco de memória e tabelas do SQLite."""
        try:
            with sqlite3.connect(self.db_name) as conn:
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
        except sqlite3.Error as e:
            Logger.error(f"Erro Crítico de DB Init: {e}")
            raise PlanifyError("Falha na inicialização do Banco de Dados", e)

    def inserir_orcamento(self, dados: Dict[str, Any]) -> None:
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO orcamentos 
                    (data_geracao, nome_obra, local, bdi, valor_total, arquivo_saida, num_itens, num_titulos, duracao_processamento)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    dados.get('data_geracao'), dados.get('nome_obra'), dados.get('local'),
                    dados.get('bdi', 0.0), dados.get('valor_total', 0.0), dados.get('arquivo_saida'),
                    dados.get('num_itens', 0), dados.get('num_titulos', 0), dados.get('duracao_processamento', 0.0)
                ))
                conn.commit()
        except sqlite3.Error as e:
            Logger.error(f"Erro de Inserção de Relatório DB: {e}")

    def buscar_estatisticas(self) -> Dict[str, Any]:
        """Agrupa e retorna as stats para renderizar em UI e Relatórios locais."""
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()

                cursor.execute("SELECT COUNT(*), SUM(valor_total), AVG(num_itens) FROM orcamentos")
                total, valor, media = cursor.fetchone()

                cursor.execute("SELECT nome_obra FROM orcamentos ORDER BY id DESC LIMIT 1")
                ultimo = cursor.fetchone()

                return {
                    'total_orcamentos': total or 0,
                    'valor_total_processado': valor or 0.0,
                    'media_itens': media or 0.0,
                    'ultimo_orcamento': {'nome': ultimo[0]} if ultimo else None
                }
        except sqlite3.Error as e:
            Logger.error(f"Erro em DB buscar_estatisticas: {e}")
            return {'total_orcamentos': 0, 'valor_total_processado': 0.0, 'media_itens': 0.0, 'ultimo_orcamento': None}

    def buscar_orcamentos(self, limite: int = 20) -> List[Dict[str, Any]]:
        try:
            with sqlite3.connect(self.db_name) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM orcamentos ORDER BY id DESC LIMIT ?", (limite,))
                return [dict(row) for row in cursor.fetchall()]
        except sqlite3.Error:
            return []

class LogHandler:
    """Delegator de logs persistentes (reservado) """
    def __init__(self, db_manager: DatabaseManager):
        self.db = db_manager