import re
import logging
from datetime import datetime, timezone
from typing import List, Optional

import pandas as pd

from . import config
from .data_models import CalibrationStandard

logger = logging.getLogger(__name__)

# --- Funzioni di Normalizzazione ---

def normalize_sp_code(sp_code_raw) -> str:
    """Normalizza il codice SP (es. 'S.P. 11/04' -> 'SP 11/04')."""
    if pd.isna(sp_code_raw):
        return ""
    s_norm = str(sp_code_raw).strip().upper()
    s_norm = s_norm.replace("S.P.", "SP").replace(".", "")
    s_norm = s_norm.replace("-", "/")
    s_norm = " ".join(s_norm.split())
    s_norm = re.sub(r'\s*SP\s*(\d+)\s*/\s*(\d+)', r'SP \1/\2', s_norm)
    return s_norm

def normalize_um(um_str_raw) -> str:
    """Normalizza una stringa di unità di misura (es. 'mm H2O' -> 'mmh2o')."""
    if pd.isna(um_str_raw):
        return ""
    s_norm = str(um_str_raw).strip().lower()
    s_norm = " ".join(s_norm.split())
    for k_map, v_map in config.MAPPA_NORMALIZZAZIONE_UM.items():
        s_norm = s_norm.replace(k_map, v_map)
    s_norm = s_norm.replace(" ", "")
    return s_norm

def normalize_range_string(range_str_raw) -> str:
    """Normalizza una stringa di range (es. '0 / 100' -> '0-100')."""
    if pd.isna(range_str_raw):
        return ""
    if isinstance(range_str_raw, (int, float)):
        range_str_raw = str(int(range_str_raw)) if range_str_raw == int(range_str_raw) else str(range_str_raw)

    norm_str = str(range_str_raw).lower()
    norm_str = " ".join(norm_str.split())
    norm_str = re.sub(r'\s*[\/÷]\s*', '-', norm_str)
    norm_str = re.sub(r'\s*to\s*', '-', norm_str, flags=re.IGNORECASE)
    norm_str = re.sub(r'\s*-\s*', '-', norm_str)
    norm_str = re.sub(r'\s+', '', norm_str)
    return norm_str

def is_cell_value_empty(cell_value) -> bool:
    """Controlla se il valore di una cella è da considerarsi vuoto."""
    if pd.isna(cell_value):
        return True
    if isinstance(cell_value, str) and not cell_value.strip():
        return True
    if isinstance(cell_value, str) and cell_value.strip().lower() == "nan":
        return True
    return False

def is_um_pressione_valida(um_str_normalized_input: str) -> bool:
    """Verifica se l'unità di misura è una di quelle di pressione riconosciute."""
    return um_str_normalized_input in config.LISTA_UM_PRESSIONE_RICONOSCIUTE


# --- Logica di Business ---

def trova_strumenti_alternativi(
    range_richiesto_raw: str,
    data_riferimento_scheda: datetime,
    strumenti_campione_list: List[CalibrationStandard]
) -> List[CalibrationStandard]:
    """
    Trova strumenti campione alternativi validi per un dato range e data.
    """
    logger.debug(f"Inizio trova_strumenti_alternativi. Range Richiesto: '{range_richiesto_raw}', Data: {data_riferimento_scheda}")
    if not strumenti_campione_list:
        logger.warning("Lista strumenti campione vuota in trova_strumenti_alternativi.")
        return []

    alternative_valide = []
    range_richiesto_norm = normalize_range_string(range_richiesto_raw)

    # Rendi la data di riferimento 'naive' per confronti omogenei
    data_riferimento_naive = data_riferimento_scheda.astimezone(timezone.utc).replace(tzinfo=None) \
        if data_riferimento_scheda.tzinfo is not None else data_riferimento_scheda

    for strumento in strumenti_campione_list:
        if not strumento.scadenza or not strumento.data_emissione:
            continue

        # Rendi naive anche le date dello strumento
        scadenza_naive = strumento.scadenza.astimezone(timezone.utc).replace(tzinfo=None) \
            if strumento.scadenza.tzinfo is not None else strumento.scadenza
        emissione_naive = strumento.data_emissione.astimezone(timezone.utc).replace(tzinfo=None) \
            if strumento.data_emissione.tzinfo is not None else strumento.data_emissione

        # Controlla validità temporale
        if not (emissione_naive <= data_riferimento_naive < scadenza_naive):
            continue

        # Controlla il range
        range_campione_norm = normalize_range_string(strumento.range)
        if range_richiesto_norm == range_campione_norm:
            alternative_valide.append(strumento)

    # Ordina per data di scadenza più lontana
    alternative_valide.sort(key=lambda x: x.scadenza, reverse=True)
    logger.debug(f"Trovate {len(alternative_valide)} alternative valide con range '{range_richiesto_norm}'.")
    return alternative_valide


from .data_models import InstrumentSheet, CertificateUsage, CompilationData
from .excel_io import parse_date_robust

def analyze_sheet_data(
    raw_data: dict,
    strumenti_campione_list: List[CalibrationStandard]
) -> InstrumentSheet:
    """
    Analizza i dati grezzi estratti da una scheda, applica la logica di business
    e restituisce un oggetto InstrumentSheet completo.
    """
    file_path = raw_data['file_path']
    base_filename = raw_data['base_filename']
    file_type = raw_data.get('file_type')
    human_error_keys = set()

    card_date = parse_date_robust(raw_data.get('card_date'), base_filename)
    if not card_date:
        return InstrumentSheet(
            file_path=file_path, base_filename=base_filename,
            status=f"Data scheda non valida: '{raw_data.get('card_date')}'",
            is_valid=False
        )

    # Estrai e normalizza SP code / Tipologia
    sp_code_raw_val = raw_data.get('sp_code')
    sp_code_normalizzato_letto = normalize_sp_code(sp_code_raw_val)
    if not sp_code_normalizzato_letto:
        human_error_keys.add(config.KEY_SP_VUOTO)
        tipologia_strumento_scheda = "SP MANCANTE"
    else:
        tipologia_strumento_scheda = config.MAPPA_SP_TIPOLOGIA.get(sp_code_normalizzato_letto, f"SP NON MAPPATO: {sp_code_raw_val}")
        if tipologia_strumento_scheda.startswith("SP NON MAPPATO"):
             logger.warning(f"{base_filename}: {tipologia_strumento_scheda} (Norm='{sp_code_normalizzato_letto}')")


    # Estrai e normalizza Modello L9 (solo per analogici)
    modello_l9_scheda_normalizzato = None
    if file_type == "analogico":
        modello_l9_raw_value = raw_data.get('modello_l9')
        if is_cell_value_empty(modello_l9_raw_value):
            human_error_keys.add(config.KEY_L9_VUOTO)
            modello_l9_scheda_normalizzato = "L9 VUOTO"
        else:
            modello_l9_temp = str(modello_l9_raw_value).strip().upper()
            modello_l9_temp = modello_l9_temp.replace('ΔP', 'DP').replace('DELTA P', 'DP').replace("SKINPOINT", "SKIN POINT")
            modello_l9_scheda_normalizzato = " ".join(modello_l9_temp.split())
            if modello_l9_scheda_normalizzato == "SKIN POINT":
                human_error_keys.add(config.KEY_L9_SKINPOINT_INCOMPLETO)

    # Logica di validazione Range/UM
    # ... (questa parte è molto complessa e verrà implementata in un secondo momento)
    # Per ora, ci concentriamo sulla struttura e sull'estrazione dei certificati.

    # Estrazione dati certificati
    extracted_certs_data = []
    cert_ids = raw_data.get('cert_ids', [])
    cert_expiries = raw_data.get('cert_expiries', [])
    cert_models = raw_data.get('cert_models', [])
    cert_ranges = raw_data.get('cert_ranges', [])

    for i in range(len(cert_ids)):
        cert_id = str(cert_ids[i]).strip() if not is_cell_value_empty(cert_ids[i]) else None
        if not cert_id:
            continue

        exp_raw = cert_expiries[i]
        cert_exp_dt = parse_date_robust(exp_raw, base_filename)
        is_exp = bool(cert_exp_dt and card_date and cert_exp_dt < card_date)

        # Verifica congruità
        is_congr = None
        congr_notes = "Verifica non implementata"
        mod_camp_reg = "N/D"
        used_before_em = False

        found_camp = next((sc for sc in strumenti_campione_list if sc.id_certificato == cert_id), None)
        if not found_camp:
            is_congr = None
            congr_notes = f"Cert.ID '{cert_id}' NON TROVATO nel registro."
        else:
            mod_camp_reg = found_camp.modello_strumento
            dt_em_camp = found_camp.data_emissione
            if dt_em_camp and card_date and card_date < dt_em_camp:
                used_before_em = True
                is_congr = False
                congr_notes = f"Usato prima dell'emissione (Scheda:{card_date:%d/%m/%Y}, Emiss:{dt_em_camp:%d/%m/%Y})"
            else:
                # La logica di congruità dettagliata andrà qui. Per ora la semplifichiamo.
                is_congr = True # Placeholder
                congr_notes = "OK (Logica dettagliata da implementare)"


        extracted_certs_data.append(
            CertificateUsage(
                file_name=base_filename,
                file_path=file_path,
                card_type=file_type,
                card_date=card_date,
                certificate_id=cert_id,
                certificate_expiry_raw=str(exp_raw),
                certificate_expiry=cert_exp_dt,
                instrument_model_on_card=str(cert_models[i]),
                instrument_range_on_card=str(cert_ranges[i]),
                is_expired_at_use=is_exp,
                tipologia_strumento_scheda=tipologia_strumento_scheda,
                modello_L9_scheda=modello_l9_scheda_normalizzato if file_type == 'analogico' else "N/A",
                modello_strumento_campione_usato=mod_camp_reg,
                is_congruent=is_congr,
                congruency_notes=congr_notes,
                used_before_emission=used_before_em
            )
        )

    # Dati per la compilazione
    comp_data = CompilationData(
        file_path=file_path,
        base_filename=base_filename,
        file_type=file_type,
        pdl_val=str(raw_data.get('pdl')).strip() if not is_cell_value_empty(raw_data.get('pdl')) else None,
        odc_val_scheda=str(raw_data.get('odc')).strip() if not is_cell_value_empty(raw_data.get('odc')) else None
    )
    # ... (la logica per popolare `campi_mancanti` andrà qui)

    return InstrumentSheet(
        file_path=file_path,
        base_filename=base_filename,
        status=f"{file_type} - {len(extracted_certs_data)} cert.",
        is_valid=True,
        card_date=card_date,
        file_type=file_type,
        tipologia_strumento=tipologia_strumento_scheda,
        modello_l9=modello_l9_scheda_normalizzato,
        certificate_usages=extracted_certs_data,
        human_error_keys=human_error_keys,
        compilation_data=comp_data
    )
