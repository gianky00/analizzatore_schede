# analyzer_app/excel_io.py
import os
import logging
from datetime import datetime, timedelta
from typing import List, Optional

import pandas as pd
from pandas.tseries.offsets import DateOffset

from . import config
from .data_models import CalibrationStandard

logger = logging.getLogger(__name__)

def parse_date_robust(date_val, context_filename: str = "N/A") -> Optional[datetime]:
    """
    Tenta di parsare una data da vari formati (stringa, timestamp, numero seriale Excel).
    """
    if pd.isna(date_val):
        return None
    if isinstance(date_val, datetime):
        return date_val
    if isinstance(date_val, pd.Timestamp):
        return date_val.to_pydatetime()

    s_date_str = str(date_val).strip()
    if not s_date_str:
        return None

    # Formati comuni
    expected_formats = ['%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d']
    common_formats = ['%m/%d/%Y', '%Y/%m/%d']
    all_formats_to_try = expected_formats + [f for f in common_formats if f not in expected_formats]

    for fmt in all_formats_to_try:
        try:
            return datetime.strptime(s_date_str.split(' ')[0], fmt)
        except ValueError:
            continue

    # Tentativo di conversione da numero seriale di Excel
    try:
        if isinstance(date_val, (int, float)) or (s_date_str.replace('.', '', 1).isdigit()):
            numeric_val = float(s_date_str)
            if 1 < numeric_val < 200000:  # Range ragionevole per date seriali
                return pd.to_datetime(numeric_val, unit='D', origin='1899-12-30').to_pydatetime()
    except (ValueError, TypeError, OverflowError) as e_num:
        logger.debug(f"File: {context_filename} - Parse numerico Excel fallito per '{s_date_str}': {e_num}")

    logger.warning(f"File: {context_filename} - Data '{s_date_str}' (raw: '{date_val}') non riconosciuta.")
    return None


def leggi_registro_strumenti() -> Optional[List[CalibrationStandard]]:
    """
    Legge il registro strumenti dal file Excel specificato nella configurazione.
    Restituisce una lista di oggetti CalibrationStandard o None in caso di errore critico.
    """
    if not config.FILE_REGISTRO_STRUMENTI:
        logger.error("Percorso FILE_REGISTRO_STRUMENTI non configurato. Impossibile leggere il registro.")
        return None

    logger.info(f"Tentativo lettura registro strumenti: {config.FILE_REGISTRO_STRUMENTI}")
    if not os.path.exists(config.FILE_REGISTRO_STRUMENTI):
        logger.error(f"File registro strumenti NON TROVATO: {config.FILE_REGISTRO_STRUMENTI}")
        return None

    try:
        cols_to_read_indices_map = {
            'modello_strumento_campione': config.REGISTRO_COL_IDX_MODELLO_STRUM_CAMPIONE,
            'id_cert_campione': config.REGISTRO_COL_IDX_ID_CERT_CAMPIONE,
            'range_campione': config.REGISTRO_COL_IDX_RANGE_CAMPIONE,
            'scadenza_cert_campione': config.REGISTRO_COL_IDX_SCADENZA_CAMPIONE
        }
        sorted_map_items = sorted(cols_to_read_indices_map.items(), key=lambda item: item[1])
        sorted_col_indices = [item[1] for item in sorted_map_items]
        sorted_col_names = [item[0] for item in sorted_map_items]

        df_registro = pd.read_excel(
            config.FILE_REGISTRO_STRUMENTI,
            sheet_name=config.REGISTRO_FOGLIO_NOME,
            header=None,
            skiprows=config.REGISTRO_RIGA_INIZIO_DATI - 1,
            usecols=sorted_col_indices,
            engine='openpyxl',
            dtype=str
        )
        df_registro.columns = sorted_col_names

        df_registro.dropna(subset=['id_cert_campione'], inplace=True)
        df_registro = df_registro[df_registro['id_cert_campione'].astype(str).str.strip() != ""]

        strumenti_campione = []
        for _, row in df_registro.iterrows():
            id_cert_strum = str(row['id_cert_campione']).strip()
            if not id_cert_strum:
                continue

            scadenza_val = row['scadenza_cert_campione']
            scadenza_dt = parse_date_robust(scadenza_val, config.FILE_REGISTRO_STRUMENTI)

            data_emissione_dt = None
            if scadenza_dt:
                try:
                    data_emissione_dt = scadenza_dt - DateOffset(years=1)
                except Exception:
                    logger.warning(f"Errore calcolo data emissione per {id_cert_strum}. Tentativo con timedelta.")
                    try:
                        data_emissione_dt = scadenza_dt - timedelta(days=365)
                    except Exception as e_delta:
                        logger.error(f"Fallito anche calcolo data emissione con timedelta per {id_cert_strum}: {e_delta}")

            modello_strum_raw = str(row['modello_strumento_campione'])
            modello_strum = modello_strum_raw.strip().upper() if modello_strum_raw.lower() != 'nan' else "N/D"

            range_strum_raw = str(row['range_campione'])
            range_strum = range_strum_raw.strip() if range_strum_raw.lower() != 'nan' else "N/D"

            scadenza_raw_str = str(scadenza_val) if str(scadenza_val).lower() != 'nan' else "N/D"

            strumenti_campione.append(
                CalibrationStandard(
                    modello_strumento=modello_strum,
                    id_certificato=id_cert_strum,
                    range=range_strum,
                    scadenza=scadenza_dt,
                    scadenza_raw=scadenza_raw_str,
                    data_emissione=data_emissione_dt
                )
            )

        logger.info(f"Letti {len(strumenti_campione)} strumenti validi dal registro.")
        return strumenti_campione

    except ValueError as ve:
        logger.error(f"Errore lettura registro (ValueError): {ve}. Verifica nome foglio, indici e nomi colonne.", exc_info=True)
        return None
    except Exception as e:
        logger.error(f"Errore imprevisto during lettura registro strumenti: {e}", exc_info=True)
        return None


from openpyxl import load_workbook

def read_instrument_sheet_raw_data(file_path: str) -> dict:
    """
    Legge i valori grezzi da un file di scheda strumento utilizzando openpyxl per efficienza.
    Restituisce un dizionario di dati grezzi. Solleva eccezioni in caso di errore.
    """
    try:
        # data_only=True per ottenere i valori calcolati delle formule
        # read_only=True per ottimizzazione
        wb = load_workbook(filename=file_path, data_only=True, read_only=True)
        ws = wb.active
    except Exception as e:
        # Gestisce file protetti da password, corrotti, etc.
        logger.error(f"Impossibile aprire o leggere il file Excel: {file_path}. Errore: {e}")
        raise IOError(f"Errore apertura file: {e}") from e

    raw_data = {'file_path': file_path, 'base_filename': os.path.basename(file_path)}

    # Funzione helper per leggere una cella in modo sicuro
    def get_cell_value(cell_coord):
        try:
            return ws[cell_coord].value
        except Exception:
            logger.warning(f"Impossibile leggere la cella '{cell_coord}' nel file {os.path.basename(file_path)}")
            return None

    # Determina tipo scheda
    model_indicator_e2 = get_cell_value('E2')
    model_indicator_e2_str = str(model_indicator_e2).strip().upper() if model_indicator_e2 else ""

    if "STRUMENTI DIGITALI" in model_indicator_e2_str:
        raw_data['file_type'] = "digitale"
        raw_data['sp_code'] = get_cell_value(config.SCHEDA_DIG_CELL_TIPOLOGIA_STRUM)
        raw_data['range_um_processo'] = get_cell_value(config.SCHEDA_DIG_CELL_RANGE_UM_PROCESSO)
        raw_data['card_date'] = get_cell_value(config.SCHEDA_DIG_CELL_DATA_COMPILAZIONE)

        # Campi anagrafici
        raw_data['odc'] = get_cell_value(config.SCHEDA_DIG_CELL_ODC)
        raw_data['pdl'] = get_cell_value(config.SCHEDA_DIG_CELL_PDL)
        raw_data['esecutore'] = get_cell_value(config.SCHEDA_DIG_CELL_ESECUTORE)
        raw_data['supervisore_isab'] = get_cell_value(config.SCHEDA_DIG_CELL_SUPERVISORE_ISAB)
        raw_data['contratto_coemi'] = get_cell_value(config.SCHEDA_DIG_CELL_CONTRATTO_COEMI)

        # Certificati
        raw_data['cert_ids'] = [get_cell_value(c) for c in ["C18", "E18", "G18"]]
        raw_data['cert_expiries'] = [get_cell_value(c) for c in ["C19", "E19", "G19"]]
        raw_data['cert_models'] = [get_cell_value(c) for c in ["C13", "E13", "G13"]]
        raw_data['cert_ranges'] = [get_cell_value(c) for c in ["C16", "E16", "G16"]]

    elif "STRUMENTI ANALOGICI" in model_indicator_e2_str:
        raw_data['file_type'] = "analogico"
        raw_data['sp_code'] = get_cell_value(config.SCHEDA_ANA_CELL_TIPOLOGIA_STRUM)
        raw_data['modello_l9'] = get_cell_value(config.SCHEDA_ANA_CELL_MODELLO_STRUM)
        raw_data['card_date'] = get_cell_value(config.SCHEDA_ANA_CELL_DATA_COMPILAZIONE)

        # Range/UM
        raw_data['range_ing'] = get_cell_value(config.SCHEDA_ANA_CELL_RANGE_INGRESSO)
        raw_data['um_ing'] = get_cell_value(config.SCHEDA_ANA_CELL_UM_INGRESSO)
        raw_data['range_usc'] = get_cell_value(config.SCHEDA_ANA_CELL_RANGE_USCITA)
        raw_data['um_usc'] = get_cell_value(config.SCHEDA_ANA_CELL_UM_USCITA)
        raw_data['range_dcs'] = get_cell_value(config.SCHEDA_ANA_CELL_RANGE_DCS)
        raw_data['um_dcs'] = get_cell_value(config.SCHEDA_ANA_CELL_UM_DCS)

        # Campi anagrafici
        raw_data['odc'] = get_cell_value(config.SCHEDA_ANA_CELL_ODC)
        raw_data['pdl'] = get_cell_value(config.SCHEDA_ANA_CELL_PDL)
        raw_data['esecutore'] = get_cell_value(config.SCHEDA_ANA_CELL_ESECUTORE)
        raw_data['supervisore_isab'] = get_cell_value(config.SCHEDA_ANA_CELL_SUPERVISORE_ISAB)
        raw_data['contratto_coemi'] = get_cell_value(config.SCHEDA_ANA_CELL_CONTRATTO_COEMI)

        # Certificati
        raw_data['cert_ids'] = [get_cell_value(c) for c in ["K43", "K44", "K45"]]
        raw_data['cert_expiries'] = [get_cell_value(c) for c in ["M43", "M44", "M45"]]
        raw_data['cert_models'] = [get_cell_value(c) for c in ["A43", "A44", "A45"]]
        raw_data['cert_ranges'] = [get_cell_value(c) for c in ["G43", "G44", "G45"]]
    else:
        raise ValueError(f"Tipo scheda non riconosciuto in E2: '{model_indicator_e2}' in {file_path}")

    wb.close()
    return raw_data
