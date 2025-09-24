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
VALIDATION_RULES = []
ANALYSIS_DATETIME = datetime.now(timezone.utc)

# --- Nomi dei fogli ---
NOME_FOGLIO_PARAMETRI = "parametri"
NOME_FOGLIO_REGOLE = "RegoleValidazione"

# --- Percorsi e Costanti Fondamentali ---
try:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    SCRIPT_DIR = os.getcwd()

PATH_FILE_PARAMETRI = os.path.normpath(os.path.join(SCRIPT_DIR, '..', "parametri.xlsm"))

# --- Costanti per Logging ---
LOG_FILENAME_TEMPLATE = "log_analisi_schede_{timestamp}.txt"
LOGS_DIR = None
LOG_FILEPATH = None

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
KEY_TIPO_SCHEDA_SCONOSCIUTO = "TIPO_SCHEDA_SCONOSCIUTO"
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
KEY_FORMULA_ERROR = "FORMULA_ERROR"

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

human_error_messages_map_descriptive = {
    KEY_TIPO_SCHEDA_SCONOSCIUTO: "Tipo scheda non riconosciuto.",
    KEY_FORMULA_ERROR: "La cella contiene un errore di formula (es. #N/A, #VALORE!).",
    KEY_CELL_RANGE_UM_NON_LEGGIBILE: "Impossibile leggere una o più celle di Range/UM.",
    KEY_COMP_ANA_ODC_MANCANTE: "ODC mancante.", KEY_COMP_ANA_DATA_COMP_MANCANTE: "Data compilazione mancante.",
    KEY_COMP_ANA_PDL_MANCANTE: "Numero PDL mancante.", KEY_COMP_ANA_ESECUTORE_MANCANTE: "Esecutore mancante.",
    KEY_COMP_ANA_SUPERVISORE_MANCANTE: "Supervisore ISAB mancante.", KEY_COMP_ANA_CONTRATTO_MANCANTE: "Contratto Coemi mancante.",
    KEY_COMP_DIG_ODC_MANCANTE: "ODC mancante.", KEY_COMP_DIG_DATA_COMP_MANCANTE: "Data compilazione mancante.",
    KEY_COMP_DIG_PDL_MANCANTE: "Numero PDL mancante.", KEY_COMP_DIG_ESECUTORE_MANCANTE: "Esecutore mancante.",
    KEY_COMP_DIG_SUPERVISORE_MANCANTE: "Supervisore ISAB mancante.", KEY_COMP_DIG_CONTRATTO_MANCANTE: "Contratto Coemi mancante.",
    KEY_COMP_ANA_CONTRATTO_DIVERSO: f"Contratto Coemi non valido. Valore atteso: '{VALORE_ATTESO_CONTRATTO_COEMI}'.",
    KEY_COMP_DIG_CONTRATTO_DIVERSO: f"Contratto Coemi non valido. Valore atteso: '{VALORE_ATTESO_CONTRATTO_COEMI}'.",
    KEY_SP_VUOTO: "Codice SP (Tipologia Strumento) mancante o illeggibile.",
    KEY_L9_VUOTO: "Modello Strumento mancante o illeggibile.",
    KEY_L9_SKINPOINT_INCOMPLETO: "Modello Strumento 'SKIN POINT' incompleto (manca tipo K, J, ecc.).",
    KEY_ERR_ANA_TEMP_CONV_C9F9_UM_DIVERSE: "Temp./Convertitore: Unità Ingresso e Unità DCS devono coincidere.",
    KEY_ERR_ANA_TEMP_CONV_F12_UM_NON_MA: f"Temp./Convertitore: Unità Uscita deve essere '{UM_MA_NORMALIZZATA}'.",
    KEY_ERR_ANA_TEMP_CONV_A9D9_RANGE_DIVERSI: "Temp./Convertitore: Range Ingresso e Range DCS devono coincidere.",
    KEY_ERR_ANA_TEMP_CONV_D12_RANGE_NON_4_20: f"Temp./Convertitore: Range Uscita deve essere '{RANGE_4_20_NORMALIZZATO}'.",
    KEY_ERR_ANA_TEMP_NOCONV_UM_NON_COINCIDENTI: "Temp./No-Conv: Unità Ingresso, DCS e Uscita devono coincidere.",
    KEY_ERR_ANA_TEMP_NOCONV_RANGE_NON_COINCIDENTI: "Temp./No-Conv: Range Ingresso, DCS e Uscita devono coincidere.",
    KEY_ERR_DIG_PRESS_D22_UM_NON_PRESSIONE: "Unità di Processo non è un'unità di pressione valida.",
    KEY_ERR_DIG_LIVELLO_D22_UM_NON_PERCENTO: f"Unità di Processo deve essere '{UM_PERCENTO_NORMALIZZATA}'.",
}

def excel_coord_to_indices(coord_str):
    match=re.match(r"([A-Z]+)([0-9]+)",coord_str.upper());
    if not match:raise ValueError(f"Coordinata Excel non valida: {coord_str}")
    col_s,row_s=match.groups();col_idx=0
    for char_i,char_v in enumerate(reversed(col_s)):col_idx+=(ord(char_v)-ord('A')+1)*(26**char_i)
    return int(row_s)-1,col_idx-1

def load_config():
    global FILE_REGISTRO_STRUMENTI, FOLDER_PATH_DEFAULT, FILE_DATI_COMPILAZIONE_SCHEDE
    global FILE_MASTER_DIGITALE_XLSX, FILE_MASTER_ANALOGICO_XLSX, LOGS_DIR
    LOGS_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, '..', 'logs'))
    if not os.path.exists(PATH_FILE_PARAMETRI):
        raise FileNotFoundError(f"ERRORE CRITICO: File parametri '{PATH_FILE_PARAMETRI}' non trovato.")
    df_params = pd.read_excel(PATH_FILE_PARAMETRI, sheet_name=NOME_FOGLIO_PARAMETRI, header=None, engine='openpyxl', dtype=str)
    try:
        path_registro_letto = df_params.iloc[1, 1]
        if pd.isna(path_registro_letto) or str(path_registro_letto).strip() == "": raise ValueError("Cella B2 (FILE_REGISTRO_STRUMENTI) nel file parametri è vuota o non valida.")
        FILE_REGISTRO_STRUMENTI = str(path_registro_letto).strip()
        if not os.path.exists(FILE_REGISTRO_STRUMENTI): raise FileNotFoundError(f"File registro strumenti specificato in B2 non trovato: {FILE_REGISTRO_STRUMENTI}")
    except IndexError: raise ValueError(f"Cella B2 non trovata nel foglio '{NOME_FOGLIO_PARAMETRI}'.")
    try:
        path_schede_letto = df_params.iloc[2, 1]
        if pd.isna(path_schede_letto) or str(path_schede_letto).strip() == "": raise ValueError("Cella B3 (FOLDER_PATH_DEFAULT) nel file parametri è vuota o non valida.")
        FOLDER_PATH_DEFAULT = str(path_schede_letto).strip()
        if not os.path.isdir(FOLDER_PATH_DEFAULT): raise NotADirectoryError(f"Cartella schede specificata in B3 non trovata: {FOLDER_PATH_DEFAULT}")
    except IndexError: raise ValueError(f"Cella B3 non trovata nel foglio '{NOME_FOGLIO_PARAMETRI}'.")
    try:
        path_compilazione_letto = df_params.iloc[3, 1]
        if not (pd.isna(path_compilazione_letto) or str(path_compilazione_letto).strip() == ""):
            path = str(path_compilazione_letto).strip()
            if os.path.exists(path): FILE_DATI_COMPILAZIONE_SCHEDE = path
            else: print(f"AVVISO: File dati compilazione specificato in B4 non trovato: {path}", file=sys.stderr)
    except IndexError: pass
    try:
        path_master_dig_letto = df_params.iloc[4, 1]
        if not (pd.isna(path_master_dig_letto) or str(path_master_dig_letto).strip() == ""):
            path = str(path_master_dig_letto).strip()
            if os.path.exists(path) and path.lower().endswith(".xlsx"): FILE_MASTER_DIGITALE_XLSX = path
            else: print(f"AVVISO: File master digitale B5 non trovato o non .xlsx: {path}", file=sys.stderr)
    except IndexError: pass
    try:
        path_master_ana_letto = df_params.iloc[5, 1]
        if not (pd.isna(path_master_ana_letto) or str(path_master_ana_letto).strip() == ""):
            path = str(path_master_ana_letto).strip()
            if os.path.exists(path) and path.lower().endswith(".xlsx"): FILE_MASTER_ANALOGICO_XLSX = path
            else: print(f"AVVISO: File master analogico B6 non trovato o non .xlsx: {path}", file=sys.stderr)
    except IndexError: pass
    _load_validation_rules()

def _load_validation_rules():
    global VALIDATION_RULES
    VALIDATION_RULES = []
    try:
        df_rules = pd.read_excel(PATH_FILE_PARAMETRI, sheet_name=NOME_FOGLIO_REGOLE, header=0, engine='openpyxl')
        df_rules.columns = [col.strip() for col in df_rules.columns]
        df_rules = df_rules.dropna(how='all')
        for index, row in df_rules.iterrows():
            is_active = str(row.get('IsActive', 'FALSE')).strip().upper()
            if is_active != 'TRUE': continue
            rule = { 'TipologiaStrumento': str(row.get('TipologiaStrumento', '*')).strip().upper() or '*', 'ModelloL9': str(row.get('ModelloL9 (Optional)', '*')).strip().upper() or '*', 'CampoA': str(row.get('CampoA', '')).strip(), 'Operatore': str(row.get('Operatore', '')).strip(), 'CampoB_o_Costante': str(row.get('CampoB_o_Costante', '')).strip(), 'ChiaveErrore': str(row.get('ChiaveErrore', '')).strip() }
            if rule['CampoA'] and rule['Operatore'] and rule['ChiaveErrore']: VALIDATION_RULES.append(rule)
            else: print(f"AVVISO: Regola alla riga {index + 2} nel foglio '{NOME_FOGLIO_REGOLE}' ignorata perché incompleta.", file=sys.stderr)
        print(f"INFO: Caricate {len(VALIDATION_RULES)} regole di validazione personalizzate dal foglio '{NOME_FOGLIO_REGOLE}'.")
    except ValueError: print(f"INFO: Foglio '{NOME_FOGLIO_REGOLE}' non trovato nel file dei parametri. Verrà usata la logica di validazione hardcoded.")
    except Exception as e: print(f"AVVISO: Impossibile caricare le regole di validazione personalizzate dal foglio '{NOME_FOGLIO_REGOLE}'. Errore: {e}", file=sys.stderr)

def _determine_log_filepath():
    global LOG_FILEPATH, LOGS_DIR
    if LOGS_DIR is None:
        LOGS_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, '..', 'logs'))
    os.makedirs(LOGS_DIR, exist_ok=True)
    timestamp_str = ANALYSIS_DATETIME.astimezone().strftime("%Y%m%d_%H%M%S")
    log_filename = LOG_FILENAME_TEMPLATE.format(timestamp=timestamp_str)
    LOG_FILEPATH = os.path.join(LOGS_DIR, log_filename)

_determine_log_filepath()