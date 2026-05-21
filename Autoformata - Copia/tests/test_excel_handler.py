import pytest
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from core.excel_handler import OrcamentoEngine

def test_parse_num():
    engine = OrcamentoEngine({})
    
    # Int and Float native detection
    assert engine._parse_num(5) == 5.0
    assert engine._parse_num(10.55) == 10.55
    
    # String conversions
    assert engine._parse_num("R$ 1.500,20") == 1500.20
    assert engine._parse_num("  R$  30 ") == 30.0
    assert engine._parse_num("1.234.567,89") == 1234567.89
    
    # Edge cases (nan, none, empty)
    assert engine._parse_num("nan") is None
    assert engine._parse_num("None") is None
    assert engine._parse_num("") is None
    assert engine._parse_num(None) is None
    
def test_aplicar_precisao():
    engine = OrcamentoEngine({})
    
    assert engine._aplicar_precisao(10.559, "TRUNC") == 10.55
    assert engine._aplicar_precisao(10.559, "ROUND") == 10.56
    assert engine._aplicar_precisao(10.559, "EXACT") == 10.559
