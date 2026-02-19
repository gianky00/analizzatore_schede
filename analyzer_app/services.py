"""
Service Layer - Logica di business scorporata dalla GUI.
"""
import logging
import multiprocessing
import os
import threading
from collections import Counter, defaultdict
from typing import List, Dict, Optional

import pandas as pd

from . import analysis, excel_io, config
from .data_models import InstrumentSheet, CertificateUsage

logger = logging.getLogger(__name__)

def _read_file_worker_internal(q, file_path):
    """Worker interno per multiprocessing (deve essere top-level)."""
    try:
        raw_data = excel_io.read_instrument_sheet_raw_data(file_path)
        q.put(('success', raw_data))
    except Exception as e:
        q.put(('error', e))

class AnalysisService:
    """Gestisce l'orchestrazione asincrona dell'analisi dei file."""
    
    def __init__(self, message_queue: multiprocessing.Queue):
        self.message_queue = message_queue
        self.stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None

    def start_analysis(self, folder_path: str, strumenti_campione: List):
        """Avvia l'analisi in un thread separato."""
        self.stop_event.clear()
        self._thread = threading.Thread(
            target=self._run_analysis_loop,
            args=(folder_path, strumenti_campione),
            daemon=True
        )
        self._thread.start()

    def _run_analysis_loop(self, folder_path: str, strumenti_campione: List):
        """Loop principale di analisi (eseguito in background)."""
        try:
            if not folder_path or not os.path.isdir(folder_path):
                raise NotADirectoryError(f"Cartella non valida: {folder_path}")

            candidate_files = [
                f for f in os.listdir(folder_path)
                if f.lower().endswith(('.xls', '.xlsx')) and not f.startswith('~')
            ]
            
            self.message_queue.put(('total_files', len(candidate_files)))
            results = []

            for i, filename in enumerate(candidate_files):
                if self.stop_event.is_set():
                    break
                    
                file_path = os.path.join(folder_path, filename)
                self.message_queue.put(('log', (f"--- INIZIO file {i+1}/{len(candidate_files)}: {filename} ---", "FILE")))
                self.message_queue.put(('progress', (i + 1, f"Analisi: {filename}")))

                try:
                    # Isolamento processo per lettura Excel (previene crash Tcl/Tk)
                    q = multiprocessing.Queue()
                    p = multiprocessing.Process(target=_read_file_worker_internal, args=(q, file_path))
                    p.start()
                    p.join(30) # Timeout 30s

                    if p.is_alive():
                        p.terminate()
                        p.join()
                        raise TimeoutError("Timeout lettura file (>30s)")

                    status, result = q.get()
                    if status == 'error':
                        raise result

                    sheet_result = analysis.analyze_sheet_data(result, strumenti_campione)
                    results.append(sheet_result)
                    
                    status_msg = "Valida" if sheet_result.is_valid else f"{len(sheet_result.human_errors)} errori"
                    self.message_queue.put(('log', (f"--- FINE: {status_msg} ---", "SUCCESS" if sheet_result.is_valid else "WARNING")))

                except Exception as e:
                    logger.error(f"Errore analisi {filename}: {e}", exc_info=True)
                    results.append(InstrumentSheet(
                        file_path=file_path, base_filename=filename,
                        status=f"Errore: {e}", is_valid=False
                    ))
                    self.message_queue.put(('log', (f"--- ERRORE: {str(e)[:50]} ---", "ERROR")))

            self.message_queue.put(('done', results))

        except Exception as e:
            logger.critical(f"Errore fatale AnalysisService: {e}", exc_info=True)
            self.message_queue.put(('error', str(e)))

class StatisticsService:
    """Gestisce il calcolo di statistiche e aggregati dai risultati di analisi."""
    
    @staticmethod
    def calculate_cert_details(analysis_results: List[InstrumentSheet]) -> Dict:
        """
        Calcola la mappa dei dettagli dei certificati dai risultati dell'analisi.
        Restituisce un dizionario compatibile con la visualizzazione Treeview.
        """
        cert_details_map = defaultdict(lambda: {
            'id': "", 'utilizzi': 0, 'date_utilizzo_obj_set': set(),
            'range_su_scheda_counter': Counter(), 'tipologie_scheda_associate_counter': Counter(),
            'usi_congrui': 0, 'usi_total_incongrui': 0, 'usi_prima_emissione': 0, 'usi_scaduti_puri': 0,
            'dettaglio_usi_list': []
        })

        all_valid_usages = [
            usage for res in analysis_results
            if res.is_valid
            for usage in res.certificate_usages
        ]

        for usage in all_valid_usages:
            if not usage.certificate_id:
                continue
                
            details = cert_details_map[usage.certificate_id]
            if not details['id']:
                details['id'] = usage.certificate_id
                
            details['utilizzi'] += 1
            details['dettaglio_usi_list'].append(usage)
            
            if usage.card_date:
                details['date_utilizzo_obj_set'].add(usage.card_date)
                
            if usage.instrument_range_on_card and usage.instrument_range_on_card != "N/D":
                details['range_su_scheda_counter'][usage.instrument_range_on_card] += 1
                
            if usage.tipologia_strumento_scheda and usage.tipologia_strumento_scheda != "N/D":
                details['tipologie_scheda_associate_counter'][usage.tipologia_strumento_scheda] += 1
                
            if usage.is_congruent is True:
                details['usi_congrui'] += 1
            elif usage.is_congruent is False:
                details['usi_total_incongrui'] += 1
                
            if usage.used_before_emission:
                details['usi_prima_emissione'] += 1
            elif usage.is_expired_at_use:
                details['usi_scaduti_puri'] += 1
                
        return cert_details_map

class AutofillService:
    """Gestisce la logica di compilazione automatica dei campi mancanti."""

    @staticmethod
    def run_autofill(
        analysis_results: List[InstrumentSheet],
        source_excel_path: str,
        log_callback: callable
    ) -> int:
        """
        Esegue la compilazione automatica basata sui dati di un file Excel sorgente.
        Returns: Numero di schede modificate.
        """
        try:
            df_source = pd.read_excel(
                source_excel_path, 
                sheet_name=config.NOME_FOGLIO_DATI_COMPILAZIONE, 
                engine='openpyxl', 
                header=0
            )
            log_callback(f"Caricati {len(df_source)} record dal file sorgente.", "SUCCESS")
        except Exception as e:
            log_callback(f"Errore lettura file sorgente: {e}", "ERROR")
            raise e

        schede_da_compilare = [
            res for res in analysis_results 
            if any(e.key.startswith("COMP_") for e in res.human_errors) 
            and res.file_path.lower().endswith('.xlsx')
        ]

        if not schede_da_compilare:
            return 0

        log_callback(f"Trovate {len(schede_da_compilare)} schede da compilare.", "INFO")

        col_mapping = {
            'data': config.COL_IDX_COMP_DATA, 
            'esecutore': config.COL_IDX_COMP_ESECUTORE, 
            'supervisore': config.COL_IDX_COMP_SUPERVISORE, 
            'odc': config.COL_IDX_COMP_ODC, 
            'pdl': config.COL_IDX_COMP_PDL
        }
        modifiche = 0

        for sheet in schede_da_compilare:
            log_callback(f"Elaborazione: {sheet.base_filename}", "FILE")
            pdl_scheda = sheet.compilation_data.pdl_val if sheet.compilation_data else None
            odc_scheda = sheet.compilation_data.odc_val_scheda if sheet.compilation_data else None
            match_row = None

            if pdl_scheda:
                try:
                    pdl_col = df_source.columns[col_mapping['pdl']]
                    matches = df_source[df_source[pdl_col].astype(str).str.strip() == str(pdl_scheda).strip()]
                    if not matches.empty:
                        match_row = matches.iloc[0]
                        log_callback(f"  Match trovato per PDL: {pdl_scheda}", "SUCCESS")
                except Exception as e:
                    log_callback(f"  Errore ricerca PDL: {e}", "WARNING")

            if match_row is None and odc_scheda:
                try:
                    odc_col = df_source.columns[col_mapping['odc']]
                    matches = df_source[df_source[odc_col].astype(str).str.strip() == str(odc_scheda).strip()]
                    if not matches.empty:
                        match_row = matches.iloc[0]
                        log_callback(f"  Match trovato per ODC: {odc_scheda}", "SUCCESS")
                except Exception as e:
                    log_callback(f"  Errore ricerca ODC: {e}", "WARNING")

            if match_row is None:
                log_callback("  Nessuna corrispondenza trovata, skip.", "WARNING")
                continue

            corrections_made = False
            for error in sheet.human_errors:
                if not error.key.startswith("COMP_") or not error.cell:
                    continue
                
                value_to_write = None
                if "ODC" in error.key:
                    value_to_write = match_row.iloc[col_mapping['odc']]
                elif "DATA" in error.key:
                    value_to_write = match_row.iloc[col_mapping['data']]
                elif "ESECUTORE" in error.key:
                    value_to_write = match_row.iloc[col_mapping['esecutore']]
                elif "SUPERVISORE" in error.key:
                    value_to_write = match_row.iloc[col_mapping['supervisore']]
                elif "PDL" in error.key:
                    value_to_write = match_row.iloc[col_mapping['pdl']]
                elif "CONTRATTO" in error.key:
                    value_to_write = config.VALORE_ATTESO_CONTRATTO_COEMI

                if value_to_write is not None and not pd.isna(value_to_write):
                    if excel_io.write_cell(sheet.file_path, error.cell, value_to_write):
                        log_callback(f"  Scritto {error.cell}: {value_to_write}", "SUCCESS")
                        corrections_made = True

            if corrections_made:
                modifiche += 1

        return modifiche

class ConfigService:
    """Gestisce l'interazione tra la GUI e il modulo di configurazione."""

    @staticmethod
    def get_all_config_paths() -> Dict[str, str]:
        """Restituisce un dizionario con tutti i percorsi configurati."""
        return {
            "FILE_REGISTRO_STRUMENTI": config.FILE_REGISTRO_STRUMENTI or "",
            "FOLDER_PATH_DEFAULT": config.FOLDER_PATH_DEFAULT or "",
            "FILE_DATI_COMPILAZIONE_SCHEDE": config.FILE_DATI_COMPILAZIONE_SCHEDE or "",
            "FILE_MASTER_DIGITALE_XLSX": config.FILE_MASTER_DIGITALE_XLSX or "",
            "FILE_MASTER_ANALOGICO_XLSX": config.FILE_MASTER_ANALOGICO_XLSX or ""
        }

    @staticmethod
    def save_and_reload(new_config: Dict[str, str]) -> bool:
        """Salva la nuova configurazione e ricarica i moduli."""
        if config.save_config(new_config):
            config.load_config_from_json()
            return True
        return False

    @staticmethod
    def is_ready_for_analysis() -> bool:
        """Verifica se i percorsi critici sono configurati e validi."""
        return config.is_config_valid()
