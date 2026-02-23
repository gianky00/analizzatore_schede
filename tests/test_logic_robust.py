from datetime import datetime

from analyzer_app import analysis, config


def create_raw_data(file_type="analogico", **overrides):
    data = {
        "file_path": "test_path.xlsx",
        "base_filename": "test_file",
        "file_type": file_type,
        "odc": "ODC-001",
        "card_date": "19/02/2026",
        "pdl": "PDL-XYZ",
        "esecutore": "Mario Rossi",
        "supervisore": "Luigi Bianchi",
        "contratto": config.VALORE_ATTESO_CONTRATTO_COEMI,
        "sp_code": "SP 01/01",
        "modello_l9": "TEST-L9",
        "cert_ids": ["CERT-001"],
        "cert_expiries": ["31/12/2026"],
        "cert_models": ["MODEL-A"],
        "cert_ranges": ["0-10 bar"]
    }
    if file_type == "analogico":
        data.update({
            "range_ing": "0-100", "um_ing": "bar",
            "range_usc": "4-20", "um_usc": "mA",
            "range_dcs": "0-100", "um_dcs": "bar"
        })
    else:
        data.update({"range_um_processo": "0-10 bar"})
    data.update(overrides)
    return data

def test_analyze_perfect_analog_sheet(sample_standard):
    raw_data = create_raw_data(file_type="analogico")
    raw_data["cert_ids"] = [sample_standard.id_certificato]
    result = analysis.analyze_sheet_data(raw_data, [sample_standard])
    assert result.is_valid is True
    assert len(result.human_errors) == 0

def test_analyze_missing_fields_digital():
    raw_data = create_raw_data(file_type="digitale", odc=None, pdl="")
    result = analysis.analyze_sheet_data(raw_data, [])
    assert result.is_valid is False
    error_keys = [e.key for e in result.human_errors]
    assert config.KEY_COMP_DIG_ODC_MANCANTE in error_keys

def test_analyze_expired_certificate(sample_standard):
    raw_data = create_raw_data(card_date="19/02/2026", cert_expiries=["01/01/2025"])
    raw_data["cert_ids"] = [sample_standard.id_certificato]
    result = analysis.analyze_sheet_data(raw_data, [sample_standard])
    assert result.certificate_usages[0].is_expired_at_use is True

def test_analyze_formula_error_handling():
    raw_data = create_raw_data(odc="#FORMULA_ERROR#")
    result = analysis.analyze_sheet_data(raw_data, [])
    error_keys = [e.key for e in result.human_errors]
    assert config.KEY_FORMULA_ERROR in error_keys

def test_verifica_congruita_logic(monkeypatch):
    monkeypatch.setattr(config, "REGOLE_CONGRUITA_CERTIFICATI_NORMALIZZATE", {
        "TEMPERATURA": {
            "modelli_campione_congrui": ["SITRANS"],
            "modelli_campione_incongrui": ["VECCHIO-MOD"]
        }
    })
    is_cong, _ = analysis._verifica_congruita_certificato("TEMPERATURA", "N/A", "SITRANS TH300")
    assert is_cong is True
    is_cong, _ = analysis._verifica_congruita_certificato("TEMPERATURA", "N/A", "VECCHIO-MOD X")
    assert is_cong is False

def test_trova_strumenti_alternativi(sample_standard):
    sample_standard.scadenza = datetime(2025, 12, 31)
    sample_standard.range = "0-100 bar"
    res = analysis.trova_strumenti_alternativi("0-100 bar", datetime(2024, 1, 1), [sample_standard])
    assert len(res) == 1
    res = analysis.trova_strumenti_alternativi("0-100 bar", datetime(2026, 1, 1), [sample_standard])
    assert len(res) == 0

def test_sp_mapping_logic():
    raw = create_raw_data(sp_code="SP 11/04")
    res = analysis.analyze_sheet_data(raw, [])
    assert res.tipologia_strumento == "LIVELLO"
    raw = create_raw_data(sp_code="SP 11/03")
    res = analysis.analyze_sheet_data(raw, [])
    assert res.tipologia_strumento == "TEMPERATURA"

def test_l9_subtype_determination():
    assert analysis._determina_sottotipo_l9("PT100") == "TEMPERATURA_RTD"
    assert analysis._determina_sottotipo_l9("RADAR") == "LIVELLO"
    assert analysis._determina_sottotipo_l9("CONVERTITORE") == "TEMPERATURA_CONVERTITORE"
    assert analysis._determina_sottotipo_l9("DP") == "PRESSIONE"

def test_temperature_converter_validation_errors():
    raw = create_raw_data(sp_code="SP 11/03", modello_l9="CONVERTITORE", um_ing="C", um_dcs="F")
    res = analysis.analyze_sheet_data(raw, [])
    error_keys = [e.key for e in res.human_errors]
    assert config.KEY_ERR_ANA_TEMP_CONV_C9F9_UM_DIVERSE in error_keys
    raw = create_raw_data(sp_code="SP 11/03", modello_l9="CONVERTITORE", um_usc="volt")
    res = analysis.analyze_sheet_data(raw, [])
    error_keys = [e.key for e in res.human_errors]
    assert config.KEY_ERR_ANA_TEMP_CONV_F12_UM_NON_MA in error_keys

def test_digital_unit_validation():
    raw = create_raw_data(file_type="digitale", sp_code="SP 11/02", range_um_processo="0-10 metri")
    res = analysis.analyze_sheet_data(raw, [])
    error_keys = [e.key for e in res.human_errors]
    assert config.KEY_ERR_DIG_PRESS_D22_UM_NON_PRESSIONE in error_keys
    raw = create_raw_data(file_type="digitale", sp_code="SP 11/04", range_um_processo="0-1000 mm")
    res = analysis.analyze_sheet_data(raw, [])
    error_keys = [e.key for e in res.human_errors]
    assert config.KEY_ERR_DIG_LIVELLO_D22_UM_NON_PERCENTO in error_keys

def test_skin_point_incomplete_error():
    raw = create_raw_data(sp_code="SP 11/03", modello_l9="SKIN POINT")
    res = analysis.analyze_sheet_data(raw, [])
    error_keys = [e.key for e in res.human_errors]
    assert config.KEY_L9_SKINPOINT_INCOMPLETO in error_keys

def test_congruity_complex_scenarios():
    is_cong, _ = analysis._verifica_congruita_certificato("TEMPERATURA", "CONVERTITORE", "MULTIMETRO DIGITALE")
    assert is_cong is True
    is_cong, _ = analysis._verifica_congruita_certificato("TEMPERATURA", "CONVERTITORE", "MANOMETRO DIGITALE")
    assert is_cong is False

def test_empty_value_detection():
    assert analysis.is_cell_value_empty(None) is True
    assert analysis.is_cell_value_empty("NaN") is True
    assert analysis.is_cell_value_empty("   ") is True

def test_dynamic_validation_rules_operators(monkeypatch):
    rules = [
        {"TipologiaStrumento": "PRESSIONE", "ModelloL9": "*", "CampoA": "odc", "Operatore": "in", "CampoB_o_Costante": "ODC-1,ODC-2", "ChiaveErrore": "ERR_IN"},
        {"TipologiaStrumento": "*", "ModelloL9": "*", "CampoA": "pdl", "Operatore": "is_empty", "CampoB_o_Costante": "", "ChiaveErrore": "ERR_EMPTY"}
    ]
    monkeypatch.setattr(config, "VALIDATION_RULES", rules)
    raw = create_raw_data(sp_code="SP 11/02", odc="ODC-1")
    res = analysis.analyze_sheet_data(raw, [])
    assert any(e.key == "ERR_IN" for e in res.human_errors)
    raw = create_raw_data(pdl=None)
    res = analysis.analyze_sheet_data(raw, [])
    assert any(e.key == "ERR_EMPTY" for e in res.human_errors)

def test_contract_variant_numeric():
    raw = create_raw_data(contratto="4600002254")
    res = analysis.analyze_sheet_data(raw, [])
    error_keys = [e.key for e in res.human_errors]
    assert config.KEY_COMP_ANA_CONTRATTO_DIVERSO not in error_keys
    raw = create_raw_data(contratto="12345")
    res = analysis.analyze_sheet_data(raw, [])
    assert config.KEY_COMP_ANA_CONTRATTO_DIVERSO in [e.key for e in res.human_errors]
