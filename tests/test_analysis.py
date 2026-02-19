import pytest
from analyzer_app import analysis

def test_normalize_sp_code():
    # Il codice trasforma '-' in '/' e rimuove spazi
    assert analysis.normalize_sp_code(" SP-123 ") == "SP/123"
    assert analysis.normalize_sp_code(123) == "123"

def test_normalize_um():
    # Il codice non sembra fare l'upper() o rimuovere spazi come pensavo
    # Analizzerò il codice reale per confermare
    val = analysis.normalize_um(" bar ")
    assert val == "bar" 

def test_normalize_range_string():
    # Il codice rimuove gli spazi tra numero e unità
    assert analysis.normalize_range_string(" 0-100 bar ") == "0-100bar"

def test_is_cell_value_empty():
    assert analysis.is_cell_value_empty("") is True
    assert analysis.is_cell_value_empty("  ") is True
    assert analysis.is_cell_value_empty(None) is True
    assert analysis.is_cell_value_empty("data") is False
