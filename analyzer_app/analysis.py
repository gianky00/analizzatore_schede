import re
import logging
from datetime import datetime, timezone
from typing import List, Optional, Dict

import pandas as pd
from pandas.tseries.offsets import DateOffset

from . import config
from .data_models import CalibrationStandard, InstrumentSheet, CertificateUsage, CompilationData, SheetError
from .excel_io import parse_date_robust

logger = logging.getLogger(__name__)

# --- Funzioni di Normalizzazione ---
def normalize_sp_code(sp_code_raw) -> str:
    if pd.isna(sp_code_raw): return ""
    s_norm = str(sp_code_raw).strip().upper().replace("S.P.","SP").replace(".","").replace("-","/")
    s_norm = " ".join(s_norm.split())
    return re.sub(r'\s*SP\s*(\d+)\s*/\s*(\d+)', r'SP \1/\2', s_norm)

def normalize_um(um_str_raw) -> str:
    if pd.isna(um_str_raw): return ""
    s_norm = str(um_str_raw).strip().lower()
    s_norm = " ".join(s_norm.split())
    for k, v in config.MAPPA_NORMALIZZAZIONE_UM.items(): s_norm = s_norm.replace(k,v)
    return s_norm.replace(" ","")

def normalize_range_string(range_str_raw) -> str:
    if pd.isna(range_str_raw): return ""
    if isinstance(range_str_raw,(int,float)):
        range_str_raw = str(int(range_str_raw)) if range_str_raw == int(range_str_raw) else str(range_str_raw)
    norm_str = str(range_str_raw).lower()
    norm_str = " ".join(norm_str.split())
    norm_str = re.sub(r'\s*[\/÷]\s*','-',norm_str)
    norm_str = re.sub(r'\s*to\s*','-',norm_str,flags=re.IGNORECASE)
    norm_str = re.sub(r'\s*-\s*','-',norm_str)
    return re.sub(r'\s+','',norm_str)

def is_cell_value_empty(cell_value) -> bool:
    if pd.isna(cell_value): return True
    if isinstance(cell_value,str) and not cell_value.strip(): return True
    return isinstance(cell_value,str) and cell_value.strip().lower()=="nan"

def is_um_pressione_valida(um: str) -> bool:
    return um in config.LISTA_UM_PRESSIONE_RICONOSCIUTE

def trova_strumenti_alternativi(
    range_richiesto_raw: str,
    data_riferimento_scheda: datetime,
    strumenti_campione_list: List[CalibrationStandard]
) -> List[CalibrationStandard]:
    if not strumenti_campione_list: return []
    alternative_valide = []
    range_richiesto_norm = normalize_range_string(range_richiesto_raw)
    data_riferimento_naive = data_riferimento_scheda.astimezone(timezone.utc).replace(tzinfo=None) if data_riferimento_scheda.tzinfo is not None else data_riferimento_scheda
    for strumento in strumenti_campione_list:
        if not strumento.scadenza or not strumento.data_emissione: continue
        scadenza_naive = strumento.scadenza.astimezone(timezone.utc).replace(tzinfo=None) if strumento.scadenza.tzinfo is not None else strumento.scadenza
        emissione_naive = strumento.data_emissione.astimezone(timezone.utc).replace(tzinfo=None) if strumento.data_emissione.tzinfo is not None else strumento.data_emissione
        if not (emissione_naive <= data_riferimento_naive < scadenza_naive): continue
        if range_richiesto_norm == normalize_range_string(strumento.range):
            alternative_valide.append(strumento)
    alternative_valide.sort(key=lambda x: x.scadenza, reverse=True)
    return alternative_valide

def analyze_sheet_data(
    raw_data: Dict,
    strumenti_campione_list: List[CalibrationStandard]
) -> InstrumentSheet:
    file_path = raw_data['file_path']
    base_filename = raw_data['base_filename']
    file_type = raw_data.get('file_type')
    human_errors: List[SheetError] = []

    def add_error(key, cell=None, suggestion=None):
        human_errors.append(SheetError(key=key, description=config.human_error_messages_map_descriptive.get(key, "Errore non definito"), cell=cell, suggestion=suggestion))

    card_date = parse_date_robust(raw_data.get('card_date'), base_filename)

    if not file_type:
        add_error(config.KEY_TIPO_SCHEDA_SCONOSCIUTO, cell="E2")
        return InstrumentSheet(
            file_path=file_path, base_filename=base_filename,
            status="Tipo scheda non riconosciuto",
            is_valid=False,
            human_errors=human_errors,
            compilation_data=CompilationData(file_path=file_path, base_filename=base_filename)
        )

    sp_code_cell = config.SCHEDA_ANA_CELL_TIPOLOGIA_STRUM if file_type == 'analogico' else config.SCHEDA_DIG_CELL_TIPOLOGIA_STRUM
    sp_code_raw_val = raw_data.get('sp_code')
    sp_code_normalizzato_letto = normalize_sp_code(sp_code_raw_val)
    tipologia_strumento_scheda = config.MAPPA_SP_TIPOLOGIA.get(sp_code_normalizzato_letto, "N/D")
    if tipologia_strumento_scheda == "N/D":
        add_error(config.KEY_SP_VUOTO, cell=sp_code_cell)

    modello_l9_scheda_normalizzato = "N/A"
    if file_type == "analogico":
        modello_l9_raw_value = raw_data.get('modello_l9')
        if is_cell_value_empty(modello_l9_raw_value):
            add_error(config.KEY_L9_VUOTO, cell=config.SCHEDA_ANA_CELL_MODELLO_STRUM)
            modello_l9_scheda_normalizzato = "L9 VUOTO"
        else:
            modello_l9_temp = str(modello_l9_raw_value).strip().upper().replace('ΔP', 'DP').replace('DELTA P', 'DP').replace("SKINPOINT", "SKIN POINT")
            modello_l9_scheda_normalizzato = " ".join(modello_l9_temp.split())
            if modello_l9_scheda_normalizzato == "SKIN POINT":
                add_error(config.KEY_L9_SKINPOINT_INCOMPLETO, cell=config.SCHEDA_ANA_CELL_MODELLO_STRUM)

    # --- Inizio Logica di Validazione ---

    # 1. Validazione campi di compilazione (ODC, Data, PDL, etc.)
    compilation_cells = {
        'analogico': {
            'ODC': (config.KEY_COMP_ANA_ODC_MANCANTE, config.SCHEDA_ANA_CELL_ODC),
            'CARD_DATE': (config.KEY_COMP_ANA_DATA_COMP_MANCANTE, config.SCHEDA_ANA_CELL_DATA_COMPILAZIONE),
            'PDL': (config.KEY_COMP_ANA_PDL_MANCANTE, config.SCHEDA_ANA_CELL_PDL),
            'ESECUTORE': (config.KEY_COMP_ANA_ESECUTORE_MANCANTE, config.SCHEDA_ANA_CELL_ESECUTORE),
            'SUPERVISORE': (config.KEY_COMP_ANA_SUPERVISORE_MANCANTE, config.SCHEDA_ANA_CELL_SUPERVISORE_ISAB),
            'CONTRATTO': (config.KEY_COMP_ANA_CONTRATTO_MANCANTE, config.SCHEDA_ANA_CELL_CONTRATTO_COEMI)
        },
        'digitale': {
            'ODC': (config.KEY_COMP_DIG_ODC_MANCANTE, config.SCHEDA_DIG_CELL_ODC),
            'CARD_DATE': (config.KEY_COMP_DIG_DATA_COMP_MANCANTE, config.SCHEDA_DIG_CELL_DATA_COMPILAZIONE),
            'PDL': (config.KEY_COMP_DIG_PDL_MANCANTE, config.SCHEDA_DIG_CELL_PDL),
            'ESECUTORE': (config.KEY_COMP_DIG_ESECUTORE_MANCANTE, config.SCHEDA_DIG_CELL_ESECUTORE),
            'SUPERVISORE': (config.KEY_COMP_DIG_SUPERVISORE_MANCANTE, config.SCHEDA_DIG_CELL_SUPERVISORE_ISAB),
            'CONTRATTO': (config.KEY_COMP_DIG_CONTRATTO_MANCANTE, config.SCHEDA_DIG_CELL_CONTRATTO_COEMI)
        }
    }
    if file_type in compilation_cells:
        for field, (key, cell) in compilation_cells[file_type].items():
            # Nota: raw_data.get('card_date') è già stato parsato. Per la validazione di vuoto,
            # dobbiamo controllare il valore grezzo originale, che è quello che fa questa logica.
            raw_value = raw_data.get(field.lower())
            if is_cell_value_empty(raw_value):
                add_error(key, cell)
            elif field == 'CONTRATTO':
                contratto_val = str(raw_value).strip()
                if not (contratto_val == config.VALORE_ATTESO_CONTRATTO_COEMI or contratto_val == config.VALORE_ATTESO_CONTRATTO_COEMI_VARIANTE_NUMERICA):
                    key_diverso = config.KEY_COMP_ANA_CONTRATTO_DIVERSO if file_type == 'analogico' else config.KEY_COMP_DIG_CONTRATTO_DIVERSO
                    add_error(key_diverso, cell, suggestion=config.VALORE_ATTESO_CONTRATTO_COEMI)

    # 2. Validazione Range/UM
    if tipologia_strumento_scheda != "N/D":
        if file_type == "analogico":
            range_ing_norm = normalize_range_string(raw_data.get('range_ing'))
            um_ing_norm = normalize_um(raw_data.get('um_ing'))
            range_usc_norm = normalize_range_string(raw_data.get('range_usc'))
            um_usc_norm = normalize_um(raw_data.get('um_usc'))
            range_dcs_norm = normalize_range_string(raw_data.get('range_dcs'))
            um_dcs_norm = normalize_um(raw_data.get('um_dcs'))

            if tipologia_strumento_scheda == "TEMPERATURA":
                if modello_l9_scheda_normalizzato == "CONVERTITORE":
                    if um_ing_norm != um_dcs_norm: add_error(config.KEY_ERR_ANA_TEMP_CONV_C9F9_UM_DIVERSE)
                    if um_usc_norm != config.UM_MA_NORMALIZZATA: add_error(config.KEY_ERR_ANA_TEMP_CONV_F12_UM_NON_MA)
                    if range_ing_norm != range_dcs_norm: add_error(config.KEY_ERR_ANA_TEMP_CONV_A9D9_RANGE_DIVERSI)
                    if range_usc_norm != config.RANGE_4_20_NORMALIZZATO: add_error(config.KEY_ERR_ANA_TEMP_CONV_D12_RANGE_NON_4_20)
                elif not modello_l9_scheda_normalizzato.startswith("L9 VUOTO"):
                    if not (um_ing_norm == um_dcs_norm and um_dcs_norm == um_usc_norm): add_error(config.KEY_ERR_ANA_TEMP_NOCONV_UM_NON_COINCIDENTI)
                    if not (range_ing_norm == range_dcs_norm and range_dcs_norm == range_usc_norm): add_error(config.KEY_ERR_ANA_TEMP_NOCONV_RANGE_NON_COINCIDENTI)

            elif tipologia_strumento_scheda == "LIVELLO":
                if modello_l9_scheda_normalizzato == "DP":
                    if not is_um_pressione_valida(um_ing_norm): add_error(config.KEY_ERR_ANA_LIVELLO_DP_C9_UM_NON_PRESSIONE)
                    if range_dcs_norm != config.RANGE_0_100_NORMALIZZATO: add_error(config.KEY_ERR_ANA_LIVELLO_DP_D9_RANGE_NON_0_100)
                    if um_dcs_norm != config.UM_PERCENTO_NORMALIZZATA: add_error(config.KEY_ERR_ANA_LIVELLO_DP_F9_UM_NON_PERCENTO)
                    if range_usc_norm != config.RANGE_4_20_NORMALIZZATO: add_error(config.KEY_ERR_ANA_LIVELLO_DP_D12_RANGE_NON_4_20)
                    if um_usc_norm != config.UM_MA_NORMALIZZATA: add_error(config.KEY_ERR_ANA_LIVELLO_DP_F12_UM_NON_MA)
                elif "BARRA DI TORSIONE" in modello_l9_scheda_normalizzato or ("TORSIONALE" in modello_l9_scheda_normalizzato and "PNEUMATICO" not in modello_l9_scheda_normalizzato and "LOCALE" not in modello_l9_scheda_normalizzato):
                    if not (um_ing_norm == config.UM_MMH2O_NORMALIZZATA or um_ing_norm == config.UM_MM_NORMALIZZATA): add_error(config.KEY_ERR_ANA_LIVELLO_TORS_C9_UM_INVALIDA)
                    if range_dcs_norm != config.RANGE_0_100_NORMALIZZATO: add_error(config.KEY_ERR_ANA_LIVELLO_TORS_D9_RANGE_NON_0_100)
                    if um_dcs_norm != config.UM_PERCENTO_NORMALIZZATA: add_error(config.KEY_ERR_ANA_LIVELLO_TORS_F9_UM_NON_PERCENTO)
                    if range_usc_norm != config.RANGE_4_20_NORMALIZZATO: add_error(config.KEY_ERR_ANA_LIVELLO_TORS_ELETTR_D12_RANGE_NON_4_20)
                    if um_usc_norm != config.UM_MA_NORMALIZZATA: add_error(config.KEY_ERR_ANA_LIVELLO_TORS_ELETTR_F12_UM_NON_MA)
                elif "TORSIONALE LOCALE" in modello_l9_scheda_normalizzato:
                     if not (um_usc_norm == config.UM_PSI_NORMALIZZATA or is_cell_value_empty(raw_data.get('um_usc'))): add_error(config.KEY_ERR_ANA_LIVELLO_TORS_LOCALE_F12_UM_NON_VUOTA)
                     if um_usc_norm != config.UM_PSI_NORMALIZZATA and not is_cell_value_empty(raw_data.get('range_usc')): add_error(config.KEY_ERR_ANA_LIVELLO_TORS_LOCALE_D12_RANGE_NON_VUOTO)
                elif modello_l9_scheda_normalizzato in ["RADAR", "ULTRASUONI", "ONDA GUIDATA"]:
                    if range_dcs_norm != config.RANGE_0_100_NORMALIZZATO: add_error(config.KEY_ERR_ANA_LIVELLO_RADARULTR_D9_RANGE_NON_0_100)
                    if um_dcs_norm != config.UM_PERCENTO_NORMALIZZATA: add_error(config.KEY_ERR_ANA_LIVELLO_RADARULTR_F9_UM_NON_PERCENTO)
                    if range_usc_norm != config.RANGE_4_20_NORMALIZZATO: add_error(config.KEY_ERR_ANA_LIVELLO_RADARULTR_D12_RANGE_NON_4_20)
                    if um_usc_norm != config.UM_MA_NORMALIZZATA: add_error(config.KEY_ERR_ANA_LIVELLO_RADARULTR_F12_UM_NON_MA)

            elif tipologia_strumento_scheda == "PRESSIONE":
                if modello_l9_scheda_normalizzato in ["DP", "TX", "TX PRESSIONE", "TX DI PRESSIONE", "CAPILLARE"]:
                    if um_ing_norm != um_dcs_norm: add_error(config.KEY_ERR_ANA_PRESS_DP_TX_C9F9_UM_DIVERSE)
                    if um_usc_norm != config.UM_MA_NORMALIZZATA: add_error(config.KEY_ERR_ANA_PRESS_DP_TX_F12_UM_NON_MA)
                    if range_ing_norm != range_dcs_norm: add_error(config.KEY_ERR_ANA_PRESS_DP_TX_A9D9_RANGE_DIVERSI)
                    if range_usc_norm != config.RANGE_4_20_NORMALIZZATO: add_error(config.KEY_ERR_ANA_PRESS_DP_TX_D12_RANGE_NON_4_20)

            elif tipologia_strumento_scheda == "PORTATA":
                if modello_l9_scheda_normalizzato == "DP" or "CAPILLARE" in modello_l9_scheda_normalizzato:
                    if range_usc_norm != config.RANGE_4_20_NORMALIZZATO: add_error(config.KEY_ERR_ANA_PORTATA_DP_D12_RANGE_NON_4_20)
                    if um_usc_norm != config.UM_MA_NORMALIZZATA: add_error(config.KEY_ERR_ANA_PORTATA_DP_F12_UM_NON_MA)

        elif file_type == "digitale":
            range_um_proc_raw = raw_data.get('range_um_processo', "")
            um_proc_norm = normalize_um(range_um_proc_raw)
            full_string_as_um_norm = normalize_um(range_um_proc_raw)

            if tipologia_strumento_scheda == "PRESSIONE":
                if not (is_um_pressione_valida(um_proc_norm) or is_um_pressione_valida(full_string_as_um_norm)):
                    add_error(config.KEY_ERR_DIG_PRESS_D22_UM_NON_PRESSIONE, config.SCHEDA_DIG_CELL_RANGE_UM_PROCESSO)
            elif tipologia_strumento_scheda == "LIVELLO":
                if config.UM_PERCENTO_NORMALIZZATA not in um_proc_norm:
                    add_error(config.KEY_ERR_DIG_LIVELLO_D22_UM_NON_PERCENTO, config.SCHEDA_DIG_CELL_RANGE_UM_PROCESSO)

    extracted_certs_data = []
    # La logica di validazione dei certificati viene eseguita solo se la data della scheda è presente.
    if card_date:
        cert_ids = raw_data.get('cert_ids', [])
        cert_expiries = raw_data.get('cert_expiries', [])
        cert_models = raw_data.get('cert_models', [])
        cert_ranges = raw_data.get('cert_ranges', [])

        for i in range(len(cert_ids)):
            cert_id = str(cert_ids[i]).strip() if not is_cell_value_empty(cert_ids[i]) else None
            if not cert_id: continue

            logger.debug(f"In sheet '{base_filename}', searching for cert ID: '{cert_id}'")

            exp_raw = cert_expiries[i]
            cert_exp_dt = parse_date_robust(exp_raw, base_filename)
            is_exp = bool(cert_exp_dt and card_date and cert_exp_dt < card_date)

            is_congr = None
            congr_notes = "Verifica non eseguita."
            mod_camp_reg = "N/D"
            used_before_em = False

            found_camp = next((sc for sc in strumenti_campione_list if sc.id_certificato == cert_id), None)
            if not found_camp:
                congr_notes = f"Cert.ID '{cert_id}' NON TROVATO nel registro."
            else:
                mod_camp_reg = found_camp.modello_strumento
                dt_em_camp = found_camp.data_emissione
                if dt_em_camp and card_date and card_date < dt_em_camp:
                    used_before_em = True; is_congr = False
                    congr_notes = f"Usato prima dell'emissione (Scheda:{card_date:%d/%m/%Y}, Emiss:{dt_em_camp:%d/%m/%Y})"
                else:
                    sott_l9_eff = "N/A"
                    if file_type == 'analogico' and not modello_l9_scheda_normalizzato.startswith(("L9 VUOTO", "Errore")):
                        if modello_l9_scheda_normalizzato in config.MAPPA_L9_SOTTOTIPO_NORMALIZZATA:
                            poss_l9_val = config.MAPPA_L9_SOTTOTIPO_NORMALIZZATA[modello_l9_scheda_normalizzato]
                            poss_l9_list = [poss_l9_val] if isinstance(poss_l9_val, str) else poss_l9_val
                            for cand_l9 in poss_l9_list:
                                if tipologia_strumento_scheda in cand_l9 or cand_l9 == tipologia_strumento_scheda:
                                    sott_l9_eff = cand_l9
                                    break
                        elif modello_l9_scheda_normalizzato: # Match parziale se non esatto
                            for l9_key_map in sorted(config.MAPPA_L9_SOTTOTIPO_NORMALIZZATA.keys(), key=len, reverse=True):
                                if l9_key_map in modello_l9_scheda_normalizzato:
                                    poss_l9_cand_val = config.MAPPA_L9_SOTTOTIPO_NORMALIZZATA[l9_key_map]
                                    poss_l9_cand_list = [poss_l9_cand_val] if isinstance(poss_l9_cand_val, str) else poss_l9_cand_val
                                    for cand_st_partial in poss_l9_cand_list:
                                        if tipologia_strumento_scheda in cand_st_partial or cand_st_partial == tipologia_strumento_scheda:
                                            sott_l9_eff = cand_st_partial
                                            break
                                    if sott_l9_eff != "N/A": break

                    if tipologia_strumento_scheda not in config.REGOLE_CONGRUITA_CERTIFICATI_NORMALIZZATE:
                        congr_notes = f"Regole congruità non definite per tipologia '{tipologia_strumento_scheda}'."
                    else:
                        reg_tip = config.REGOLE_CONGRUITA_CERTIFICATI_NORMALIZZATE[tipologia_strumento_scheda]
                        is_congr_prov, congr_notes_prov = False, f"INCONGRUO (default): '{mod_camp_reg}' per {tipologia_strumento_scheda} (L9:'{modello_l9_scheda_normalizzato}',SottL9Eff:'{sott_l9_eff}')."

                        # Caso speciale: LIVELLO con MANOMETRO DIGITALE
                        if tipologia_strumento_scheda == "LIVELLO" and mod_camp_reg == "MANOMETRO DIGITALE":
                            if file_type == 'digitale':
                                is_congr_prov, congr_notes_prov = True, "OK (LIV digitale con MAN DIG)."
                            elif file_type == 'analogico':
                                um_usc_norm = normalize_um(raw_data.get('um_usc'))
                                cond_dp_liv = (modello_l9_scheda_normalizzato == "DP")
                                cond_tors_pneu_liv = ("TORSIONALE PNEUMATICO" in modello_l9_scheda_normalizzato and um_usc_norm == config.UM_PSI_NORMALIZZATA)
                                cond_tors_locale_liv = ("TORSIONALE LOCALE" in modello_l9_scheda_normalizzato and um_usc_norm == config.UM_PSI_NORMALIZZATA)
                                cond_capillare_liv = ("CAPILLARE" in modello_l9_scheda_normalizzato and um_usc_norm == config.UM_MA_NORMALIZZATA)

                                if cond_dp_liv or cond_tors_pneu_liv or cond_tors_locale_liv or cond_capillare_liv:
                                    is_congr_prov, congr_notes_prov = True, f"OK (LIV {modello_l9_scheda_normalizzato} an. con MAN DIG e UM Uscita corretta)."
                                else:
                                    is_congr_prov, congr_notes_prov = False, f"INCONGRUO: MAN DIG per LIV an. L9='{modello_l9_scheda_normalizzato}' non supportato o UM uscita errata."

                        elif "eccezioni_l9_incongrui" in reg_tip and sott_l9_eff != "N/A" and sott_l9_eff in reg_tip["eccezioni_l9_incongrui"] and mod_camp_reg in reg_tip["eccezioni_l9_incongrui"][sott_l9_eff]:
                            is_congr_prov, congr_notes_prov = False, f"INCONGRUO (eccL9):'{mod_camp_reg}' per {tipologia_strumento_scheda}({sott_l9_eff})."
                        elif mod_camp_reg in reg_tip.get("modelli_campione_incongrui", []):
                            if "sottotipi_l9" in reg_tip and sott_l9_eff != "N/A" and sott_l9_eff in reg_tip["sottotipi_l9"] and mod_camp_reg in reg_tip["sottotipi_l9"][sott_l9_eff]:
                                is_congr_prov, congr_notes_prov = True, f"OK (sottL9 sovrascrive incongruo gen.):'{mod_camp_reg}' per {tipologia_strumento_scheda}({sott_l9_eff})."
                            else:
                                is_congr_prov, congr_notes_prov = False, f"INCONGRUO (lista gen):'{mod_camp_reg}' per {tipologia_strumento_scheda}."
                        elif "sottotipi_l9" in reg_tip and sott_l9_eff != "N/A" and sott_l9_eff in reg_tip["sottotipi_l9"] and mod_camp_reg in reg_tip["sottotipi_l9"][sott_l9_eff]:
                            is_congr_prov, congr_notes_prov = True, f"OK (sottL9):'{mod_camp_reg}' per {tipologia_strumento_scheda}({sott_l9_eff})."
                        elif mod_camp_reg in reg_tip.get("modelli_campione_congrui", []):
                            is_congr_prov, congr_notes_prov = True, "OK (regole base)."

                        is_congr, congr_notes = is_congr_prov, congr_notes_prov

            extracted_certs_data.append(
                CertificateUsage(
                    file_name=base_filename, file_path=file_path, card_type=file_type, card_date=card_date,
                    certificate_id=cert_id, certificate_expiry_raw=str(exp_raw), certificate_expiry=cert_exp_dt,
                    instrument_model_on_card=str(cert_models[i]), instrument_range_on_card=str(cert_ranges[i]),
                    is_expired_at_use=is_exp, tipologia_strumento_scheda=tipologia_strumento_scheda,
                    modello_L9_scheda=modello_l9_scheda_normalizzato, modello_strumento_campione_usato=mod_camp_reg,
                    is_congruent=is_congr, congruency_notes=congr_notes, used_before_emission=used_before_em
                )
            )

    comp_data = CompilationData(
        file_path=file_path, base_filename=base_filename, file_type=file_type,
        pdl_val=str(raw_data.get('pdl')).strip() if not is_cell_value_empty(raw_data.get('pdl')) else None,
        odc_val_scheda=str(raw_data.get('odc')).strip() if not is_cell_value_empty(raw_data.get('odc')) else None
    )

    return InstrumentSheet(
        file_path=file_path, base_filename=base_filename,
        status=f"{file_type} - {len(extracted_certs_data)} cert.",
        is_valid=True, card_date=card_date, file_type=file_type,
        tipologia_strumento=tipologia_strumento_scheda,
        modello_l9=modello_l9_scheda_normalizzato,
        certificate_usages=extracted_certs_data,
        human_errors=human_errors,
        compilation_data=comp_data
    )
