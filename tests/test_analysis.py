from analyzer_app import analysis


def test_normalize_sp_code():
    assert analysis.normalize_sp_code(" SP-123 ") == "SP/123"
    assert analysis.normalize_sp_code(123) == "123"

def test_normalize_um():
    assert analysis.normalize_um(" bar ") == "bar"

def test_normalize_range_string():
    assert analysis.normalize_range_string(" 0-100 bar ") == "0-100bar"

def test_is_cell_value_empty():
    assert analysis.is_cell_value_empty("") is True
    assert analysis.is_cell_value_empty(None) is True

def test_advanced_normalization():
    # Il codice attuale trasforma '.' in '' prima della regex
    assert analysis.normalize_sp_code("S.P. 11.04") == "SP 1104"
