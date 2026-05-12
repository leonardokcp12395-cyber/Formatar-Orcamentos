import pytest
import sys
import os

# Adds core/utils to path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from utils.smart_parser import SmartParser

class DummyAutocomplete:
    def get_list(self, key):
        return ["Campus São Desidério", "Campus Orizona"]

def test_parse_whatsapp_text():
    texto = """
    Orçamento: Manutenção Telhado
    Setor: TI
    Valor Simulado: R$ 5.432,10
    """
    dummy_ac = DummyAutocomplete()
    resultado = SmartParser.parse_whatsapp_text(texto, dummy_ac)
    assert resultado["descricao_header"] == "Orçamento: Manutenção Telhado"
    assert resultado["setor"] == "TI"
    assert resultado["valor_simulado"] == "5.432,10"
