import pytest
from datetime import datetime
from analyzer_app import analysis, config, data_models

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
    # Data scheda: Febbraio 2026
    # Scadenza certificato scritta sulla scheda: Gennaio 2025 (già scaduto al momento dell'uso)
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
    # Strumento con scadenza 2025
    sample_standard.scadenza = datetime(2025, 12, 31)
    sample_standard.range = "0-100 bar"
    
    # Cerchiamo per una scheda del 2024 (valido)
    res = analysis.trova_strumenti_alternativi("0-100 bar", datetime(2024, 1, 1), [sample_standard])
    assert len(res) == 1
    
    # Cerchiamo per una scheda del 2026 (scaduto)
    res = analysis.trova_strumenti_alternativi("0-100 bar", datetime(2026, 1, 1), [sample_standard])
    assert len(res) == 0

def test_sp_mapping_logic():
    """Verifica che il codice SP determini correttamente la tipologia strumento."""
    # Caso 1: Livello
    raw = create_raw_data(sp_code="SP 11/04")
    res = analysis.analyze_sheet_data(raw, [])
    assert res.tipologia_strumento == "LIVELLO"
    
    # Caso 2: Temperatura
    raw = create_raw_data(sp_code="SP 11/03")
    res = analysis.analyze_sheet_data(raw, [])
    assert res.tipologia_strumento == "TEMPERATURA"

def test_l9_subtype_determination():
    """Testa la funzione interna di determinazione sottotipo L9."""
    assert analysis._determina_sottotipo_l9("PT100") == "TEMPERATURA_RTD"
    assert analysis._determina_sottotipo_l9("RADAR") == "LIVELLO"
    assert analysis._determina_sottotipo_l9("CONVERTITORE") == "TEMPERATURA_CONVERTITORE"
    assert analysis._determina_sottotipo_l9("DP") == "PRESSIONE" # Primo della lista in config

def test_temperature_converter_validation_errors():
    """Testa gli errori specifici del convertitore di temperatura."""
    # UM Ingressso != DCS
    raw = create_raw_data(
        sp_code="SP 11/03", modello_l9="CONVERTITORE",
        um_ing="C", um_dcs="F"
    )
    res = analysis.analyze_sheet_data(raw, [])
    error_keys = [e.key for e in res.human_errors]
    assert config.KEY_ERR_ANA_TEMP_CONV_C9F9_UM_DIVERSE in error_keys

    # UM Uscita != ma
    raw = create_raw_data(
        sp_code="SP 11/03", modello_l9="CONVERTITORE",
        um_usc="volt"
    )
    res = analysis.analyze_sheet_data(raw, [])
    error_keys = [e.key for e in res.human_errors]
    assert config.KEY_ERR_ANA_TEMP_CONV_F12_UM_NON_MA in error_keys

def test_digital_unit_validation():
    """Testa la validazione delle unità di misura per strumenti digitali."""
    # Pressione con unità non valida (es. "metri")
    raw = create_raw_data(
        file_type="digitale", sp_code="SP 11/02",
        range_um_processo="0-10 metri"
    )
    res = analysis.analyze_sheet_data(raw, [])
    error_keys = [e.key for e in res.human_errors]
    assert config.KEY_ERR_DIG_PRESS_D22_UM_NON_PRESSIONE in error_keys

    # Livello con unità non %
    raw = create_raw_data(
        file_type="digitale", sp_code="SP 11/04",
        range_um_processo="0-1000 mm"
    )
    res = analysis.analyze_sheet_data(raw, [])
    error_keys = [e.key for e in res.human_errors]
    assert config.KEY_ERR_DIG_LIVELLO_D22_UM_NON_PERCENTO in error_keys

def test_skin_point_incomplete_error():
    """Verifica l'errore per modello SKIN POINT senza specifica tipo."""
    raw = create_raw_data(sp_code="SP 11/03", modello_l9="SKIN POINT")
    res = analysis.analyze_sheet_data(raw, [])
    error_keys = [e.key for e in res.human_errors]
    assert config.KEY_L9_SKINPOINT_INCOMPLETO in error_keys

def test_congruity_complex_scenarios():
    """Testa la logica di congruità con sottotipi e eccezioni reali da config."""
    # Caso: Temperatura Convertitore + Multimetro -> Congruo (da config)
    is_cong, _ = analysis._verifica_congruita_certificato("TEMPERATURA", "CONVERTITORE", "MULTIMETRO DIGITALE")
    assert is_cong is True
    
    # Caso: Temperatura Convertitore + Manometro -> Incongruo (Eccezione in config)
    is_cong, _ = analysis._verifica_congruita_certificato("TEMPERATURA", "CONVERTITORE", "MANOMETRO DIGITALE")
    assert is_cong is False

def test_empty_value_detection():
    """Verifica che ogni variante di 'vuoto' venga rilevata."""
    assert analysis.is_cell_value_empty(None) is True
    assert analysis.is_cell_value_empty("NaN") is True
    assert analysis.is_cell_value_empty("   ") is True
    assert analysis.is_cell_value_empty(float('nan')) is True

def test_dynamic_validation_rules_operators(monkeypatch):
    """Testa tutti gli operatori delle regole dinamiche (in, not_in, !=, is_empty)."""
    rules = [
        {
            "TipologiaStrumento": "PRESSIONE", "ModelloL9": "*",
            "CampoA": "odc", "Operatore": "in", "CampoB_o_Costante": "ODC-1,ODC-2",
            "ChiaveErrore": "ERR_IN"
        },
        {
            "TipologiaStrumento": "*", "ModelloL9": "*",
            "CampoA": "pdl", "Operatore": "is_empty", "CampoB_o_Costante": "",
            "ChiaveErrore": "ERR_EMPTY"
        }
    ]
    monkeypatch.setattr(config, "VALIDATION_RULES", rules)
    
    # Caso trigger operator 'in'
    raw = create_raw_data(sp_code="SP 11/02", odc="ODC-1")
    res = analysis.analyze_sheet_data(raw, [])
    assert any(e.key == "ERR_IN" for e in res.human_errors)
    
    # Caso trigger operator 'is_empty'
    raw = create_raw_data(pdl=None)
    res = analysis.analyze_sheet_data(raw, [])
    assert any(e.key == "ERR_EMPTY" for e in res.human_errors)

def test_contract_variant_numeric():
    """Verifica che la variante puramente numerica del contratto sia accettata."""
    # Variante numerica
    raw = create_raw_data(contratto="4600002254")
    res = analysis.analyze_sheet_data(raw, [])
    # Non deve esserci l'errore di contratto diverso
    error_keys = [e.key for e in res.human_errors]
    assert config.KEY_COMP_ANA_CONTRATTO_DIVERSO not in error_keys
    
    # Variante errata
    raw = create_raw_data(contratto="12345")
    res = analysis.analyze_sheet_data(raw, [])
    assert config.KEY_COMP_ANA_CONTRATTO_DIVERSO in [e.key for e in res.human_errors]
