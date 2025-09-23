# analyzer_app/config.py

import os
import sys
import re
from datetime import datetime, timezone
import pandas as pd

# --- Variabili di configurazione che verranno popolate da load_config() ---
FILE_REGISTRO_STRUMENTI = None
FOLDER_PATH_DEFAULT = None
FILE_DATI_COMPILAZIONE_SCHEDE = None
FILE_MASTER_DIGITALE_XLSX = None
FILE_MASTER_ANALOGICO_XLSX = None
SCRIPT_DIR = None
ANALYSIS_DATETIME = datetime.now(timezone.utc)

# --- Costanti per Logging ---
LOG_FILENAME_TEMPLATE = "log_analisi_schede_{timestamp}.txt"
LOGS_DIR = None # Verrà impostato in load_config
LOG_FILEPATH = None # Verrà determinato da _determine_log_filepath

# --- Costanti per il Registro Strumenti ---
REGISTRO_COL_IDX_MODELLO_STRUM_CAMPIONE = 6
REGISTRO_COL_IDX_ID_CERT_CAMPIONE = 16
REGISTRO_COL_IDX_RANGE_CAMPIONE = 12
REGISTRO_COL_IDX_SCADENZA_CAMPIONE = 18
REGISTRO_RIGA_INIZIO_DATI = 7
REGISTRO_FOGLIO_NOME = "strumenti campione ISAB SUD"
SOGLIA_PER_SUGGERIMENTO_ALTERNATIVO = 5

# --- Costanti per le Schede (Coordinate Celle) ---
SCHEDA_DIG_CELL_TIPOLOGIA_STRUM = "N10"
SCHEDA_DIG_CELL_RANGE_UM_PROCESSO = "D22"
SCHEDA_ANA_CELL_TIPOLOGIA_STRUM = "N9"
SCHEDA_ANA_CELL_MODELLO_STRUM = "L9"
SCHEDA_ANA_CELL_RANGE_INGRESSO = "A9"; SCHEDA_ANA_CELL_UM_INGRESSO = "C9"
SCHEDA_ANA_CELL_RANGE_USCITA = "D12"; SCHEDA_ANA_CELL_UM_USCITA = "F12"
SCHEDA_ANA_CELL_RANGE_DCS = "D9"; SCHEDA_ANA_CELL_UM_DCS = "F9"
SCHEDA_ANA_CELL_ODC = "L50"
SCHEDA_ANA_CELL_DATA_COMPILAZIONE = "B50"
SCHEDA_ANA_CELL_PDL = "F50"
SCHEDA_ANA_CELL_ESECUTORE = "F52"
SCHEDA_ANA_CELL_SUPERVISORE_ISAB = "L52"
SCHEDA_ANA_CELL_CONTRATTO_COEMI = "B52"
SCHEDA_DIG_CELL_ODC = "L45"
SCHEDA_DIG_CELL_DATA_COMPILAZIONE = "B45"
SCHEDA_DIG_CELL_PDL = "F45"
SCHEDA_DIG_CELL_ESECUTORE = "F47"
SCHEDA_DIG_CELL_SUPERVISORE_ISAB = "L47"
SCHEDA_DIG_CELL_CONTRATTO_COEMI = "B47"

# --- Costanti per la Logica di Business ---
VALORE_ATTESO_CONTRATTO_COEMI = "COEMI 4600002254"
VALORE_ATTESO_CONTRATTO_COEMI_VARIANTE_NUMERICA = "4600002254"
NOME_FOGLIO_DATI_COMPILAZIONE = "RIASSUNTO"
COL_IDX_COMP_DATA = 0
COL_IDX_COMP_ESECUTORE = 1
COL_IDX_COMP_SUPERVISORE = 3
COL_IDX_COMP_ODC = 4
COL_IDX_COMP_PDL = 5

# --- Chiavi di Errore ---
KEY_SP_VUOTO = "SP_VUOTO_O_ILLEGGIBILE"
KEY_L9_VUOTO = "L9_VUOTO_O_ILLEGGIBILE"
KEY_L9_SKINPOINT_INCOMPLETO = "L9_SKINPOINT_INCOMPLETO"
KEY_CELL_RANGE_UM_NON_LEGGIBILE = "RANGE_UM_CELL_ILLEGGIBILE"
KEY_ERR_ANA_TEMP_CONV_C9F9_UM_DIVERSE = "ERR_ANA_TEMP_CONV_C9F9_UM_DIVERSE"
KEY_ERR_ANA_TEMP_CONV_F12_UM_NON_MA = "ERR_ANA_TEMP_CONV_F12_UM_NON_MA"
KEY_ERR_ANA_TEMP_CONV_A9D9_RANGE_DIVERSI = "ERR_ANA_TEMP_CONV_A9D9_RANGE_DIVERSI"
KEY_ERR_ANA_TEMP_CONV_D12_RANGE_NON_4_20 = "ERR_ANA_TEMP_CONV_D12_RANGE_NON_4_20"
KEY_ERR_ANA_TEMP_NOCONV_UM_NON_COINCIDENTI = "ERR_ANA_TEMP_NOCONV_UM_NON_COINCIDENTI"
KEY_ERR_ANA_TEMP_NOCONV_RANGE_NON_COINCIDENTI = "ERR_ANA_TEMP_NOCONV_RANGE_NON_COINCIDENTI"
KEY_ERR_ANA_LIVELLO_DP_C9_UM_NON_PRESSIONE = "ERR_ANA_LIVELLO_DP_C9_UM_NON_PRESSIONE"
KEY_ERR_ANA_LIVELLO_DP_D9_RANGE_NON_0_100 = "ERR_ANA_LIVELLO_DP_D9_RANGE_NON_0_100"
KEY_ERR_ANA_LIVELLO_DP_F9_UM_NON_PERCENTO = "ERR_ANA_LIVELLO_DP_F9_UM_NON_PERCENTO"
KEY_ERR_ANA_LIVELLO_DP_D12_RANGE_NON_4_20 = "ERR_ANA_LIVELLO_DP_D12_RANGE_NON_4_20"
KEY_ERR_ANA_LIVELLO_DP_F12_UM_NON_MA = "ERR_ANA_LIVELLO_DP_F12_UM_NON_MA"
KEY_ERR_ANA_LIVELLO_TORS_C9_UM_INVALIDA = "ERR_ANA_LIVELLO_TORS_C9_UM_INVALIDA"
KEY_ERR_ANA_LIVELLO_TORS_D9_RANGE_NON_0_100 = "ERR_ANA_LIVELLO_TORS_D9_RANGE_NON_0_100"
KEY_ERR_ANA_LIVELLO_TORS_F9_UM_NON_PERCENTO = "ERR_ANA_LIVELLO_TORS_F9_UM_NON_PERCENTO"
KEY_ERR_ANA_LIVELLO_TORS_ELETTR_D12_RANGE_NON_4_20 = "ERR_ANA_LIVELLO_TORS_ELETTR_D12_RANGE_NON_4_20"
KEY_ERR_ANA_LIVELLO_TORS_ELETTR_F12_UM_NON_MA = "ERR_ANA_LIVELLO_TORS_ELETTR_F12_UM_NON_MA"
KEY_ERR_ANA_LIVELLO_TORS_LOCALE_D12_RANGE_NON_VUOTO = "ERR_ANA_LIVELLO_TORS_LOCALE_D12_RANGE_NON_VUOTO"
KEY_ERR_ANA_LIVELLO_TORS_LOCALE_F12_UM_NON_VUOTA = "ERR_ANA_LIVELLO_TORS_LOCALE_F12_UM_NON_VUOTA"
KEY_ERR_ANA_LIVELLO_RADARULTR_D9_RANGE_NON_0_100 = "ERR_ANA_LIVELLO_RADARULTR_D9_RANGE_NON_0_100"
KEY_ERR_ANA_LIVELLO_RADARULTR_F9_UM_NON_PERCENTO = "ERR_ANA_LIVELLO_RADARULTR_F9_UM_NON_PERCENTO"
KEY_ERR_ANA_LIVELLO_RADARULTR_D12_RANGE_NON_4_20 = "ERR_ANA_LIVELLO_RADARULTR_D12_RANGE_NON_4_20"
KEY_ERR_ANA_LIVELLO_RADARULTR_F12_UM_NON_MA = "ERR_ANA_LIVELLO_RADARULTR_F12_UM_NON_MA"
KEY_ERR_ANA_PRESS_DP_TX_C9F9_UM_DIVERSE = "ERR_ANA_PRESS_DP_TX_C9F9_UM_DIVERSE"
KEY_ERR_ANA_PRESS_DP_TX_F12_UM_NON_MA = "ERR_ANA_PRESS_DP_TX_F12_UM_NON_MA"
KEY_ERR_ANA_PRESS_DP_TX_A9D9_RANGE_DIVERSI = "ERR_ANA_PRESS_DP_TX_A9D9_RANGE_DIVERSI"
KEY_ERR_ANA_PRESS_DP_TX_D12_RANGE_NON_4_20 = "ERR_ANA_PRESS_DP_TX_D12_RANGE_NON_4_20"
KEY_ERR_ANA_PORTATA_DP_D12_RANGE_NON_4_20 = "ERR_ANA_PORTATA_DP_D12_RANGE_NON_4_20"
KEY_ERR_ANA_PORTATA_DP_F12_UM_NON_MA = "ERR_ANA_PORTATA_DP_F12_UM_NON_MA"
KEY_ERR_DIG_PRESS_D22_UM_NON_PRESSIONE = "ERR_DIG_PRESS_D22_UM_NON_PRESSIONE"
KEY_ERR_DIG_LIVELLO_D22_UM_NON_PERCENTO = "ERR_DIG_LIVELLO_D22_UM_NON_PERCENTO"
KEY_COMP_ANA_ODC_MANCANTE = "COMP_ANA_ODC_MANCANTE"
KEY_COMP_ANA_DATA_COMP_MANCANTE = "COMP_ANA_DATA_COMP_MANCANTE"
KEY_COMP_ANA_PDL_MANCANTE = "COMP_ANA_PDL_MANCANTE"
KEY_COMP_ANA_ESECUTORE_MANCANTE = "COMP_ANA_ESECUTORE_MANCANTE"
KEY_COMP_ANA_SUPERVISORE_MANCANTE = "COMP_ANA_SUPERVISORE_MANCANTE"
KEY_COMP_ANA_CONTRATTO_MANCANTE = "COMP_ANA_CONTRATTO_MANCANTE"
KEY_COMP_ANA_CONTRATTO_DIVERSO = "COMP_ANA_CONTRATTO_DIVERSO"
KEY_COMP_DIG_ODC_MANCANTE = "COMP_DIG_ODC_MANCANTE"
KEY_COMP_DIG_DATA_COMP_MANCANTE = "COMP_DIG_DATA_COMP_MANCANTE"
KEY_COMP_DIG_PDL_MANCANTE = "COMP_DIG_PDL_MANCANTE"
KEY_COMP_DIG_ESECUTORE_MANCANTE = "COMP_DIG_ESECUTORE_MANCANTE"
KEY_COMP_DIG_SUPERVISORE_MANCANTE = "COMP_DIG_SUPERVISORE_MANCANTE"
KEY_COMP_DIG_CONTRATTO_MANCANTE = "COMP_DIG_CONTRATTO_MANCANTE"
KEY_COMP_DIG_CONTRATTO_DIVERSO = "COMP_DIG_CONTRATTO_DIVERSO"
KEY_COMP_CAMPI_MANCANTI_NON_LEGGIBILI = "COMP_CAMPI_MANCANTI_NON_LEGGIBILI"

# --- Mappe e Regole ---
MAPPA_SP_TIPOLOGIA = {"SP 11/04":"LIVELLO","SP 11-04":"LIVELLO","SP 11/03":"TEMPERATURA","SP 11-03":"TEMPERATURA","SP 11/02":"PRESSIONE","SP 11-02":"PRESSIONE","SP 11/01":"PORTATA","SP 11-01":"PORTATA"}
MAPPA_L9_SOTTOTIPO_NORMALIZZATA = {"DP":["PRESSIONE","PORTATA","LIVELLO"],"CAPILLARE":["PRESSIONE","PORTATA","LIVELLO"],"TX":["PRESSIONE"],"TX PRESSIONE":["PRESSIONE"],"TX DI PRESSIONE":["PRESSIONE"],"TORSIONALE":["LIVELLO"],"TORSIONALE PNEUMATICO":["LIVELLO"],"TORSIONALE LOCALE":["LIVELLO"],"BARRA DI TORSIONE":["LIVELLO"],"ONDA GUIDATA":["LIVELLO"],"RADAR":["LIVELLO"],"ULTRASUONI":["LIVELLO","PORTATA"],"K":["TEMPERATURA_TERMOCOPPIA"],"J":["TEMPERATURA_TERMOCOPPIA"],"SKIN POINT K":["TEMPERATURA_TERMOCOPPIA"],"SKIN POINT J":["TEMPERATURA_TERMOCOPPIA"],"TERMOCOPPIA K":["TEMPERATURA_TERMOCOPPIA"],"TERMOCOPPIA J":["TEMPERATURA_TERMOCOPPIA"],"TERMOCOPPIA":["TEMPERATURA_TERMOCOPPIA"],"TEMOCOPPIA J":["TEMPERATURA_TERMOCOPPIA"],"TC K":["TEMPERATURA_TERMOCOPPIA"],"TC J":["TEMPERATURA_TERMOCOPPIA"],"K TC":["TEMPERATURA_TERMOCOPPIA"],"J TC":["TEMPERATURA_TERMOCOPPIA"],"TERMOCOPPIA TIPO K":["TEMPERATURA_TERMOCOPPIA"],"TERMOCOPPIA TIPO J":["TEMPERATURA_TERMOCOPPIA"],"RTD 3W":["TEMPERATURA_RTD"],"RTD":["TEMPERATURA_RTD"],"RTD 2W":["TEMPERATURA_RTD"],"TERMORESISTENZA":["TEMPERATURA_RTD"],"PT100":["TEMPERATURA_RTD"],"CONVERTITORE":["TEMPERATURA_CONVERTITORE"],"INDICATORE LOCALE":["LIVELLO","PORTATA","PRESSIONE","TEMPERATURA"]}
REGOLE_CONGRUITA_CERTIFICATI_NORMALIZZATE = {"TEMPERATURA":{"modelli_campione_congrui":["CALIBR. TEMPERATURA"],"sottotipi_l9":{"TEMPERATURA_TERMOCOPPIA":["TERMOCOPPIA CAMPIONE","MULTIMETRO DIGITALE"],"TEMPERATURA_RTD":["TERMORESISTENZA CAMPIONE","MULTIMETRO DIGITALE"],"TEMPERATURA_CONVERTITORE":["MULTIMETRO DIGITALE","CALIBRATORE DI LOOP"]},"modelli_campione_incongrui":["MANOMETRO DIGITALE","CALIBRATORE DI LOOP"],"eccezioni_l9_incongrui":{"TEMPERATURA_CONVERTITORE":["MANOMETRO DIGITALE","TERMOCOPPIA CAMPIONE","TERMORESISTENZA CAMPIONE"]}},"PRESSIONE":{"modelli_campione_congrui":["MANOMETRO DIGITALE","MULTIMETRO DIGITALE","CALIBRATORE DI LOOP"],"modelli_campione_incongrui":["CALIBR. TEMPERATURA","TERMOCOPPIA CAMPIONE","TERMORESISTENZA CAMPIONE"]},"PORTATA":{"modelli_campione_congrui":["MANOMETRO DIGITALE","MULTIMETRO DIGITALE","CALIBRATORE DI LOOP"],"modelli_campione_incongrui":["CALIBR. TEMPERATURA","TERMOCOPPIA CAMPIONE","TERMORESISTENZA CAMPIONE"]},"LIVELLO":{"modelli_campione_congrui":["MULTIMETRO DIGITALE","CALIBRATORE DI LOOP","COMPARATORE", "MANOMETRO DIGITALE"],"modelli_campione_incongrui":["CALIBR. TEMPERATURA","TERMOCOPPIA CAMPIONE","TERMORESISTENZA CAMPIONE"]}}
for tipologia_regola, item_regola in REGOLE_CONGRUITA_CERTIFICATI_NORMALIZZATE.items():
    if "modelli_campione_congrui" in item_regola: item_regola["modelli_campione_congrui"] = [m.strip().upper() for m in item_regola.get("modelli_campione_congrui", [])]
    if "modelli_campione_incongrui" in item_regola: item_regola["modelli_campione_incongrui"] = [m.strip().upper() for m in item_regola.get("modelli_campione_incongrui", [])]
    if "sottotipi_l9" in item_regola and isinstance(item_regola["sottotipi_l9"], dict):
        for sottotipo_key, modelli_sottotipo_list in item_regola["sottotipi_l9"].items():
            if isinstance(modelli_sottotipo_list, list): item_regola["sottotipi_l9"][sottotipo_key] = [m.strip().upper() for m in modelli_sottotipo_list]
    if "eccezioni_l9_incongrui" in item_regola and isinstance(item_regola["eccezioni_l9_incongrui"], dict):
        for eccezione_l9_key, modelli_eccezione_list in item_regola["eccezioni_l9_incongrui"].items():
            if isinstance(modelli_eccezione_list, list): item_regola["eccezioni_l9_incongrui"][eccezione_l9_key] = [m.strip().upper() for m in modelli_eccezione_list]

LISTA_UM_PRESSIONE_RICONOSCIUTE = sorted(["bar","barg","bara","mbar","mbarg","mbara","pa","kpa","mpa","psi","psig","psia","mmh2o","cmh2o","mh2o","mmhg","cmhg","mhg","kg/cm2"])
MAPPA_NORMALIZZAZIONE_UM = {"mm h2o":"mmh2o","mmh₂o":"mmh2o","mm H₂O":"mmh2o","mm H2O":"mmh2o","kg/cm²":"kg/cm2","kg/cm^2":"kg/cm2","milliampere":"ma","milli ampere":"ma","milliamperes":"ma","mamp":"ma","percent":"%","percentage":"%"}
RANGE_0_100_NORMALIZZATO="0-100";RANGE_4_20_NORMALIZZATO="4-20";UM_MA_NORMALIZZATA="ma";UM_PERCENTO_NORMALIZZATA="%";UM_MMH2O_NORMALIZZATA="mmh2o";UM_MM_NORMALIZZATA="mm";UM_PSI_NORMALIZZATA="psi"

# --- Mappa Messaggi di Errore Umani ---
human_error_messages_map_descriptive = {
    KEY_SP_VUOTO: f"Codice SP (Tipologia Strumento da N9/N10) mancante o illeggibile.", KEY_L9_VUOTO: f"Modello Strumento (L9, solo per analogici) mancante o illeggibile.", KEY_L9_SKINPOINT_INCOMPLETO: f"Modello L9 (analogico) 'SKIN POINT' incompleto (manca tipo K, J, ecc.).", KEY_CELL_RANGE_UM_NON_LEGGIBILE: f"Impossibile leggere una o più celle necessarie per Range/UM (A9,C9,D9,F9,D12,F12 per analogici; D22 per digitali).",
    KEY_ERR_ANA_TEMP_CONV_C9F9_UM_DIVERSE: f"Temp./Convertitore (ANA): Unità Ingresso ({SCHEDA_ANA_CELL_UM_INGRESSO}) diversa da Unità DCS ({SCHEDA_ANA_CELL_UM_DCS}). Devono coincidere.", KEY_ERR_ANA_TEMP_CONV_F12_UM_NON_MA: f"Temp./Convertitore (ANA): Unità Uscita ({SCHEDA_ANA_CELL_UM_USCITA}) deve essere '{UM_MA_NORMALIZZATA}'.", KEY_ERR_ANA_TEMP_CONV_A9D9_RANGE_DIVERSI: f"Temp./Convertitore (ANA): Range Ingresso ({SCHEDA_ANA_CELL_RANGE_INGRESSO}) diverso da Range DCS ({SCHEDA_ANA_CELL_RANGE_DCS}). Devono coincidere.", KEY_ERR_ANA_TEMP_CONV_D12_RANGE_NON_4_20: f"Temp./Convertitore (ANA): Range Uscita ({SCHEDA_ANA_CELL_RANGE_USCITA}) deve essere '{RANGE_4_20_NORMALIZZATO}'.", KEY_ERR_ANA_TEMP_NOCONV_UM_NON_COINCIDENTI: f"Temp./NO-Convertitore (RTD, TC) (ANA): Unità Ingresso ({SCHEDA_ANA_CELL_UM_INGRESSO}), DCS ({SCHEDA_ANA_CELL_UM_DCS}) e Uscita ({SCHEDA_ANA_CELL_UM_USCITA}) devono coincidere.", KEY_ERR_ANA_TEMP_NOCONV_RANGE_NON_COINCIDENTI: f"Temp./NO-Convertitore (RTD, TC) (ANA): Range Ingresso ({SCHEDA_ANA_CELL_RANGE_INGRESSO}), DCS ({SCHEDA_ANA_CELL_RANGE_DCS}) e Uscita ({SCHEDA_ANA_CELL_RANGE_USCITA}) devono coincidere.",
    KEY_ERR_ANA_LIVELLO_DP_C9_UM_NON_PRESSIONE: f"Livello/DP (ANA): Unità Ingresso ({SCHEDA_ANA_CELL_UM_INGRESSO}) deve essere un'unità di misura di pressione valida (es. mmH2O, bar).", KEY_ERR_ANA_LIVELLO_DP_D9_RANGE_NON_0_100: f"Livello/DP (ANA): Range DCS ({SCHEDA_ANA_CELL_RANGE_DCS}) deve essere '{RANGE_0_100_NORMALIZZATO}'.", KEY_ERR_ANA_LIVELLO_DP_F9_UM_NON_PERCENTO: f"Livello/DP (ANA): Unità DCS ({SCHEDA_ANA_CELL_UM_DCS}) deve essere '{UM_PERCENTO_NORMALIZZATA}'.", KEY_ERR_ANA_LIVELLO_DP_D12_RANGE_NON_4_20: f"Livello/DP (ANA): Range Uscita ({SCHEDA_ANA_CELL_RANGE_USCITA}) deve essere '{RANGE_4_20_NORMALIZZATO}'.", KEY_ERR_ANA_LIVELLO_DP_F12_UM_NON_MA: f"Livello/DP (ANA): Unità Uscita ({SCHEDA_ANA_CELL_UM_USCITA}) deve essere '{UM_MA_NORMALIZZATA}'.", KEY_ERR_ANA_LIVELLO_TORS_C9_UM_INVALIDA: f"Livello/BarraTors. (ANA): Unità Ingresso ({SCHEDA_ANA_CELL_UM_INGRESSO}) deve essere '{UM_MMH2O_NORMALIZZATA}' o '{UM_MM_NORMALIZZATA}'.", KEY_ERR_ANA_LIVELLO_TORS_D9_RANGE_NON_0_100: f"Livello/BarraTors. (ANA): Range DCS ({SCHEDA_ANA_CELL_RANGE_DCS}) deve essere '{RANGE_0_100_NORMALIZZATO}'.", KEY_ERR_ANA_LIVELLO_TORS_F9_UM_NON_PERCENTO: f"Livello/BarraTors. (ANA): Unità DCS ({SCHEDA_ANA_CELL_UM_DCS}) deve essere '{UM_PERCENTO_NORMALIZZATA}'.", KEY_ERR_ANA_LIVELLO_TORS_ELETTR_D12_RANGE_NON_4_20: f"Livello/BarraTors.Elettronico (ANA): Range Uscita ({SCHEDA_ANA_CELL_RANGE_USCITA}) deve essere '{RANGE_4_20_NORMALIZZATO}'.", KEY_ERR_ANA_LIVELLO_TORS_ELETTR_F12_UM_NON_MA: f"Livello/BarraTors.Elettronico (ANA): Unità Uscita ({SCHEDA_ANA_CELL_UM_USCITA}) deve essere '{UM_MA_NORMALIZZATA}'.", KEY_ERR_ANA_LIVELLO_TORS_LOCALE_D12_RANGE_NON_VUOTO: f"Livello/BarraTors.Locale (ANA): Range Uscita ({SCHEDA_ANA_CELL_RANGE_USCITA}) deve essere vuoto (o '{UM_PSI_NORMALIZZATA}' se pneumatico).", KEY_ERR_ANA_LIVELLO_TORS_LOCALE_F12_UM_NON_VUOTA: f"Livello/BarraTors.Locale (ANA): Unità Uscita ({SCHEDA_ANA_CELL_UM_USCITA}) deve essere vuota (o '{UM_PSI_NORMALIZZATA}' se pneumatico).", KEY_ERR_ANA_LIVELLO_RADARULTR_D9_RANGE_NON_0_100: f"Livello/Radar-Ultrasuoni (ANA): Range DCS ({SCHEDA_ANA_CELL_RANGE_DCS}) deve essere '{RANGE_0_100_NORMALIZZATO}'.", KEY_ERR_ANA_LIVELLO_RADARULTR_F9_UM_NON_PERCENTO: f"Livello/Radar-Ultrasuoni (ANA): Unità DCS ({SCHEDA_ANA_CELL_UM_DCS}) deve essere '{UM_PERCENTO_NORMALIZZATA}'.", KEY_ERR_ANA_LIVELLO_RADARULTR_D12_RANGE_NON_4_20: f"Livello/Radar-Ultrasuoni (ANA): Range Uscita ({SCHEDA_ANA_CELL_RANGE_USCITA}) deve essere '{RANGE_4_20_NORMALIZZATO}'.", KEY_ERR_ANA_LIVELLO_RADARULTR_F12_UM_NON_MA: f"Livello/Radar-Ultrasuoni (ANA): Unità Uscita ({SCHEDA_ANA_CELL_UM_USCITA}) deve essere '{UM_MA_NORMALIZZATA}'.",
    KEY_ERR_ANA_PRESS_DP_TX_C9F9_UM_DIVERSE: f"Pressione/DP-TX-Capillare (ANA): Unità Ingresso ({SCHEDA_ANA_CELL_UM_INGRESSO}) e DCS ({SCHEDA_ANA_CELL_UM_DCS}) devono coincidere.", KEY_ERR_ANA_PRESS_DP_TX_F12_UM_NON_MA: f"Pressione/DP-TX-Capillare (ANA): Unità Uscita ({SCHEDA_ANA_CELL_UM_USCITA}) deve essere '{UM_MA_NORMALIZZATA}'.", KEY_ERR_ANA_PRESS_DP_TX_A9D9_RANGE_DIVERSI: f"Pressione/DP-TX-Capillare (ANA): Range Ingresso ({SCHEDA_ANA_CELL_RANGE_INGRESSO}) e DCS ({SCHEDA_ANA_CELL_RANGE_DCS}) devono coincidere.", KEY_ERR_ANA_PRESS_DP_TX_D12_RANGE_NON_4_20: f"Pressione/DP-TX-Capillare (ANA): Range Uscita ({SCHEDA_ANA_CELL_RANGE_USCITA}) deve essere '{RANGE_4_20_NORMALIZZATO}'.",
    KEY_ERR_ANA_PORTATA_DP_D12_RANGE_NON_4_20: f"Portata/DP (ANA): Range Uscita ({SCHEDA_ANA_CELL_RANGE_USCITA}) deve essere '{RANGE_4_20_NORMALIZZATO}'.", KEY_ERR_ANA_PORTATA_DP_F12_UM_NON_MA: f"Portata/DP (ANA): Unità Uscita ({SCHEDA_ANA_CELL_UM_USCITA}) deve essere '{UM_MA_NORMALIZZATA}'.",
    KEY_ERR_DIG_PRESS_D22_UM_NON_PRESSIONE: f"Digitale/Pressione: Unità Processo ({SCHEDA_DIG_CELL_RANGE_UM_PROCESSO}) non è UM di pressione valida.", KEY_ERR_DIG_LIVELLO_D22_UM_NON_PERCENTO: f"Digitale/Livello: Unità Processo ({SCHEDA_DIG_CELL_RANGE_UM_PROCESSO}) deve essere '{UM_PERCENTO_NORMALIZZATA}'.",
    KEY_COMP_ANA_ODC_MANCANTE: f"Scheda Analogica: ODC ({SCHEDA_ANA_CELL_ODC}) mancante.", KEY_COMP_ANA_DATA_COMP_MANCANTE: f"Scheda Analogica: Data Compilazione ({SCHEDA_ANA_CELL_DATA_COMPILAZIONE}) mancante.", KEY_COMP_ANA_PDL_MANCANTE: f"Scheda Analogica: Numero PDL ({SCHEDA_ANA_CELL_PDL}) mancante.", KEY_COMP_ANA_ESECUTORE_MANCANTE: f"Scheda Analogica: Esecutore ({SCHEDA_ANA_CELL_ESECUTORE}) mancante.", KEY_COMP_ANA_SUPERVISORE_MANCANTE: f"Scheda Analogica: Supervisore ISAB ({SCHEDA_ANA_CELL_SUPERVISORE_ISAB}) mancante.", KEY_COMP_ANA_CONTRATTO_MANCANTE: f"Scheda Analogica: Contratto Coemi ({SCHEDA_ANA_CELL_CONTRATTO_COEMI}) mancante.", KEY_COMP_ANA_CONTRATTO_DIVERSO: f"Scheda Analogica: Contratto Coemi ({SCHEDA_ANA_CELL_CONTRATTO_COEMI}) diverso da '{VALORE_ATTESO_CONTRATTO_COEMI}' o '{VALORE_ATTESO_CONTRATTO_COEMI_VARIANTE_NUMERICA}'.",
    KEY_COMP_DIG_ODC_MANCANTE: f"Scheda Digitale: ODC ({SCHEDA_DIG_CELL_ODC}) mancante.", KEY_COMP_DIG_DATA_COMP_MANCANTE: f"Scheda Digitale: Data Compilazione ({SCHEDA_DIG_CELL_DATA_COMPILAZIONE}) mancante.", KEY_COMP_DIG_PDL_MANCANTE: f"Scheda Digitale: Numero PDL ({SCHEDA_DIG_CELL_PDL}) mancante.", KEY_COMP_DIG_ESECUTORE_MANCANTE: f"Scheda Digitale: Esecutore ({SCHEDA_DIG_CELL_ESECUTORE}) mancante.", KEY_COMP_DIG_SUPERVISORE_MANCANTE: f"Scheda Digitale: Supervisore ISAB ({SCHEDA_DIG_CELL_SUPERVISORE_ISAB}) mancante.", KEY_COMP_DIG_CONTRATTO_MANCANTE: f"Scheda Digitale: Contratto Coemi ({SCHEDA_DIG_CELL_CONTRATTO_COEMI}) mancante.", KEY_COMP_DIG_CONTRATTO_DIVERSO: f"Scheda Digitale: Contratto Coemi ({SCHEDA_DIG_CELL_CONTRATTO_COEMI}) diverso da '{VALORE_ATTESO_CONTRATTO_COEMI}' o '{VALORE_ATTESO_CONTRATTO_COEMI_VARIANTE_NUMERICA}'.", KEY_COMP_CAMPI_MANCANTI_NON_LEGGIBILI: "Impossibile leggere uno o più campi per la verifica di completezza (ODC, Data, PDL, etc.).",
}

# --- Funzioni di Utilità e Indici Derivati ---
def excel_coord_to_indices(coord_str):
    match=re.match(r"([A-Z]+)([0-9]+)",coord_str.upper());
    if not match:raise ValueError(f"Coordinata Excel non valida: {coord_str}")
    col_s,row_s=match.groups();col_idx=0
    for char_i,char_v in enumerate(reversed(col_s)):col_idx+=(ord(char_v)-ord('A')+1)*(26**char_i)
    return int(row_s)-1,col_idx-1

IDX_DIG_TIPOLOGIA_STRUM=excel_coord_to_indices(SCHEDA_DIG_CELL_TIPOLOGIA_STRUM);IDX_ANA_TIPOLOGIA_STRUM=excel_coord_to_indices(SCHEDA_ANA_CELL_TIPOLOGIA_STRUM)
IDX_ANA_MODELLO_STRUM=excel_coord_to_indices(SCHEDA_ANA_CELL_MODELLO_STRUM);IDX_ANA_RANGE_INGRESSO=excel_coord_to_indices(SCHEDA_ANA_CELL_RANGE_INGRESSO)
IDX_ANA_UM_INGRESSO=excel_coord_to_indices(SCHEDA_ANA_CELL_UM_INGRESSO);IDX_ANA_RANGE_USCITA=excel_coord_to_indices(SCHEDA_ANA_CELL_RANGE_USCITA)
IDX_ANA_UM_USCITA=excel_coord_to_indices(SCHEDA_ANA_CELL_UM_USCITA);IDX_ANA_RANGE_DCS=excel_coord_to_indices(SCHEDA_ANA_CELL_RANGE_DCS)
IDX_ANA_UM_DCS=excel_coord_to_indices(SCHEDA_ANA_CELL_UM_DCS);IDX_DIG_RANGE_UM_PROCESSO=excel_coord_to_indices(SCHEDA_DIG_CELL_RANGE_UM_PROCESSO)
IDX_ANA_UM_USCITA_ROW, IDX_ANA_UM_USCITA_COL = IDX_ANA_UM_USCITA
IDX_ANA_ODC = excel_coord_to_indices(SCHEDA_ANA_CELL_ODC)
IDX_ANA_DATA_COMPILAZIONE = excel_coord_to_indices(SCHEDA_ANA_CELL_DATA_COMPILAZIONE)
IDX_ANA_PDL = excel_coord_to_indices(SCHEDA_ANA_CELL_PDL)
IDX_ANA_ESECUTORE = excel_coord_to_indices(SCHEDA_ANA_CELL_ESECUTORE)
IDX_ANA_SUPERVISORE_ISAB = excel_coord_to_indices(SCHEDA_ANA_CELL_SUPERVISORE_ISAB)
IDX_ANA_CONTRATTO_COEMI = excel_coord_to_indices(SCHEDA_ANA_CELL_CONTRATTO_COEMI)
IDX_DIG_ODC = excel_coord_to_indices(SCHEDA_DIG_CELL_ODC)
IDX_DIG_DATA_COMPILAZIONE = excel_coord_to_indices(SCHEDA_DIG_CELL_DATA_COMPILAZIONE)
IDX_DIG_PDL = excel_coord_to_indices(SCHEDA_DIG_CELL_PDL)
IDX_DIG_ESECUTORE = excel_coord_to_indices(SCHEDA_DIG_CELL_ESECUTORE)
IDX_DIG_SUPERVISORE_ISAB = excel_coord_to_indices(SCHEDA_DIG_CELL_SUPERVISORE_ISAB)
IDX_DIG_CONTRATTO_COEMI = excel_coord_to_indices(SCHEDA_DIG_CELL_CONTRATTO_COEMI)

# --- Funzione di Caricamento Configurazione ---
def load_config():
    """
    Legge il file parametri.xlsm e popola le variabili di configurazione globali.
    Solleva FileNotFoundError, NotADirectoryError, o ValueError in caso di problemi critici.
    """
    global FILE_REGISTRO_STRUMENTI, FOLDER_PATH_DEFAULT, FILE_DATI_COMPILAZIONE_SCHEDE
    global FILE_MASTER_DIGITALE_XLSX, FILE_MASTER_ANALOGICO_XLSX, LOGS_DIR

    try:
        SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        SCRIPT_DIR = os.getcwd()

    PATH_FILE_PARAMETRI = os.path.normpath(os.path.join(SCRIPT_DIR, '..', "parametri.xlsm"))
    NOME_FOGLIO_PARAMETRI = "parametri"
    LOGS_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, '..', 'logs'))


    print(f"INFO: Tentativo di lettura percorsi da: {PATH_FILE_PARAMETRI}")

    if not os.path.exists(PATH_FILE_PARAMETRI):
        raise FileNotFoundError(f"ERRORE CRITICO: File parametri '{PATH_FILE_PARAMETRI}' non trovato.")

    df_params = pd.read_excel(PATH_FILE_PARAMETRI, sheet_name=NOME_FOGLIO_PARAMETRI, header=None, engine='openpyxl', dtype=str)

    # Cella B2: FILE_REGISTRO_STRUMENTI
    try:
        path_registro_letto = df_params.iloc[1, 1]
        if pd.isna(path_registro_letto) or str(path_registro_letto).strip() == "":
            raise ValueError("Cella B2 (FILE_REGISTRO_STRUMENTI) nel file parametri è vuota o non valida.")
        FILE_REGISTRO_STRUMENTI = str(path_registro_letto).strip()
        if not os.path.exists(FILE_REGISTRO_STRUMENTI):
            raise FileNotFoundError(f"File registro strumenti specificato in B2 non trovato: {FILE_REGISTRO_STRUMENTI}")
    except IndexError:
        raise ValueError(f"Cella B2 non trovata nel foglio '{NOME_FOGLIO_PARAMETRI}'.")

    # Cella B3: FOLDER_PATH_DEFAULT
    try:
        path_schede_letto = df_params.iloc[2, 1]
        if pd.isna(path_schede_letto) or str(path_schede_letto).strip() == "":
            raise ValueError("Cella B3 (FOLDER_PATH_DEFAULT) nel file parametri è vuota o non valida.")
        FOLDER_PATH_DEFAULT = str(path_schede_letto).strip()
        if not os.path.isdir(FOLDER_PATH_DEFAULT):
            raise NotADirectoryError(f"Cartella schede specificata in B3 non trovata: {FOLDER_PATH_DEFAULT}")
    except IndexError:
        raise ValueError(f"Cella B3 non trovata nel foglio '{NOME_FOGLIO_PARAMETRI}'.")

    # Cella B4: FILE_DATI_COMPILAZIONE_SCHEDE (Opzionale)
    try:
        path_compilazione_letto = df_params.iloc[3, 1]
        if not (pd.isna(path_compilazione_letto) or str(path_compilazione_letto).strip() == ""):
            path = str(path_compilazione_letto).strip()
            if os.path.exists(path):
                FILE_DATI_COMPILAZIONE_SCHEDE = path
            else:
                print(f"AVVISO: File dati compilazione specificato in B4 non trovato: {path}", file=sys.stderr)
    except IndexError:
        pass

    # Cella B5: FILE_MASTER_DIGITALE_XLSX (Opzionale)
    try:
        path_master_dig_letto = df_params.iloc[4, 1]
        if not (pd.isna(path_master_dig_letto) or str(path_master_dig_letto).strip() == ""):
            path = str(path_master_dig_letto).strip()
            if os.path.exists(path) and path.lower().endswith(".xlsx"):
                FILE_MASTER_DIGITALE_XLSX = path
            else:
                 print(f"AVVISO: File master digitale B5 non trovato o non .xlsx: {path}", file=sys.stderr)
    except IndexError:
        pass

    # Cella B6: FILE_MASTER_ANALOGICO_XLSX (Opzionale)
    try:
        path_master_ana_letto = df_params.iloc[5, 1]
        if not (pd.isna(path_master_ana_letto) or str(path_master_ana_letto).strip() == ""):
            path = str(path_master_ana_letto).strip()
            if os.path.exists(path) and path.lower().endswith(".xlsx"):
                FILE_MASTER_ANALOGICO_XLSX = path
            else:
                print(f"AVVISO: File master analogico B6 non trovato o non .xlsx: {path}", file=sys.stderr)
    except IndexError:
        pass


def _determine_log_filepath():
    """Determina il percorso per il file di log."""
    global LOG_FILEPATH, LOGS_DIR
    if LOGS_DIR is None:
        try:
            SCRIPT_DIR_TEMP = os.path.dirname(os.path.abspath(__file__))
            LOGS_DIR = os.path.normpath(os.path.join(SCRIPT_DIR_TEMP, '..', 'logs'))
        except NameError:
             LOGS_DIR = "logs"

    os.makedirs(LOGS_DIR, exist_ok=True)
    timestamp_str = ANALYSIS_DATETIME.astimezone().strftime("%Y%m%d_%H%M%S")
    log_filename = LOG_FILENAME_TEMPLATE.format(timestamp=timestamp_str)
    LOG_FILEPATH = os.path.join(LOGS_DIR, log_filename)

_determine_log_filepath()
