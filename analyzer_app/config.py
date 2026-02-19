# analyzer_app/config.py
"""
Modulo di configurazione per l'Analizzatore Schede Taratura.
Salva e carica la configurazione da file JSON.
"""

import json
import logging
import os
import re
import sys
from datetime import UTC, datetime

logger = logging.getLogger(__name__)

# ============================================================================
# VARIABILI DI CONFIGURAZIONE
# ============================================================================
FILE_REGISTRO_STRUMENTI: str | None = None
FOLDER_PATH_DEFAULT: str | None = None
FILE_DATI_COMPILAZIONE_SCHEDE: str | None = None
FILE_MASTER_DIGITALE_XLSX: str | None = None
FILE_MASTER_ANALOGICO_XLSX: str | None = None
VALIDATION_RULES: list[dict] = []
ANALYSIS_DATETIME = datetime.now(UTC)

# ============================================================================
# PERCORSI E FILE
# ============================================================================
try:
    # Se eseguito come exe PyInstaller
    if getattr(sys, 'frozen', False):
        SCRIPT_DIR = os.path.dirname(sys.executable)
    else:
        SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
        SCRIPT_DIR = os.path.dirname(SCRIPT_DIR)  # Vai alla root del progetto
except Exception:
    SCRIPT_DIR = os.getcwd()

CONFIG_FILE_PATH = os.path.join(SCRIPT_DIR, "config.json")

# ============================================================================
# NOMI FOGLI EXCEL
# ============================================================================
NOME_FOGLIO_DATI_COMPILAZIONE = "RIASSUNTO"

# ============================================================================
# REGISTRO STRUMENTI - INDICI COLONNE (0-based)
# ============================================================================
REGISTRO_COL_IDX_MODELLO_STRUM_CAMPIONE = 6
REGISTRO_COL_IDX_ID_CERT_CAMPIONE = 16
REGISTRO_COL_IDX_RANGE_CAMPIONE = 12
REGISTRO_COL_IDX_SCADENZA_CAMPIONE = 18
REGISTRO_RIGA_INIZIO_DATI = 7
REGISTRO_FOGLIO_NOME = "strumenti campione ISAB SUD"
SOGLIA_PER_SUGGERIMENTO_ALTERNATIVO = 5

# ============================================================================
# CELLE SCHEDA DIGITALE
# ============================================================================
SCHEDA_DIG_CELL_TIPOLOGIA_STRUM = "N10"
SCHEDA_DIG_CELL_RANGE_UM_PROCESSO = "D22"
SCHEDA_DIG_CELL_ODC = "L45"
SCHEDA_DIG_CELL_DATA_COMPILAZIONE = "B45"
SCHEDA_DIG_CELL_PDL = "F45"
SCHEDA_DIG_CELL_ESECUTORE = "F47"
SCHEDA_DIG_CELL_SUPERVISORE_ISAB = "L47"
SCHEDA_DIG_CELL_CONTRATTO_COEMI = "B47"

# ============================================================================
# CELLE SCHEDA ANALOGICA
# ============================================================================
SCHEDA_ANA_CELL_TIPOLOGIA_STRUM = "N9"
SCHEDA_ANA_CELL_MODELLO_STRUM = "L9"
SCHEDA_ANA_CELL_RANGE_INGRESSO = "A9"
SCHEDA_ANA_CELL_UM_INGRESSO = "C9"
SCHEDA_ANA_CELL_RANGE_USCITA = "D12"
SCHEDA_ANA_CELL_UM_USCITA = "F12"
SCHEDA_ANA_CELL_RANGE_DCS = "D9"
SCHEDA_ANA_CELL_UM_DCS = "F9"
SCHEDA_ANA_CELL_ODC = "L50"
SCHEDA_ANA_CELL_DATA_COMPILAZIONE = "B50"
SCHEDA_ANA_CELL_PDL = "F50"
SCHEDA_ANA_CELL_ESECUTORE = "F52"
SCHEDA_ANA_CELL_SUPERVISORE_ISAB = "L52"
SCHEDA_ANA_CELL_CONTRATTO_COEMI = "B52"

# ============================================================================
# CELLE CERTIFICATI
# ============================================================================
# Analogico
SCHEDA_ANA_CERT_IDS = ["K43", "K44", "K45"]
SCHEDA_ANA_CERT_EXPIRIES = ["M43", "M44", "M45"]
SCHEDA_ANA_CERT_MODELS = ["A43", "A44", "A45"]
SCHEDA_ANA_CERT_RANGES = ["G43", "G44", "G45"]

# Digitale
SCHEDA_DIG_CERT_IDS = ["C18", "E18", "G18"]
SCHEDA_DIG_CERT_EXPIRIES = ["C19", "E19", "G19"]
SCHEDA_DIG_CERT_MODELS = ["C13", "E13", "G13"]
SCHEDA_DIG_CERT_RANGES = ["E13", "G13", "I13"] # Nota: da verificare se corretti per il template

# ============================================================================
# INDICI COLONNE FILE DATI COMPILAZIONE
# ============================================================================
COL_IDX_COMP_DATA = 0
COL_IDX_COMP_ESECUTORE = 1
COL_IDX_COMP_SUPERVISORE = 2
COL_IDX_COMP_ODC = 3
COL_IDX_COMP_PDL = 4

# ============================================================================
# COSTANTI E MAPPATURE
# ============================================================================
INDICATORE_STRUMENTI_DIGITALI = "STRUMENTI DIGITALI"
INDICATORE_STRUMENTI_ANALOGICI = "STRUMENTI ANALOGICI"
VALORE_ATTESO_CONTRATTO_COEMI = "3400006767"
VALORE_ATTESO_CONTRATTO_COEMI_VARIANTE_NUMERICA = 3400006767

# Mappatura Tipologia -> Modello L9 atteso/possibile
# (Solo come riferimento o per estensioni future, la logica principale è nelle regole congruenza)
MAPPA_L9_SOTTOTIPO_NORMALIZZATA = {
    "CONVERTITORE": ["CONVERTITORE"],
    "TERMOCOPPIA": ["TERMOCOPPIA"],
    "TERMORESISTENZA": ["TERMORESISTENZA"],
    "TRASMETTITORE": ["TRASMETTITORE"],
    "MANOMETRO": ["MANOMETRO"],
    "TERMOMETRO": ["TERMOMETRO"],
    "LIVELLO": ["LIVELLO"], # Generico
    "PORTATA": ["PORTATA"], # Generico
    "SKIN": ["SKIN POINT"]
}

# Regole di congruità per Tipologia Strumento (cella N9/N10)
# Struttura:
# Tipologia -> {
#   modelli_campione_congrui: [lista modelli certificato validi],
#   modelli_campione_incongrui: [lista modelli certificato invalidi],
#   sottotipi_l9: {
#       SOTTOTIPO_L9: [lista modelli certificato validi specifici per questo sottotipo]
#   },
#   eccezioni_l9_incongrui: {
#       SOTTOTIPO_L9: [lista modelli certificato invalidi specifici per questo sottotipo]
#   }
# }
REGOLE_CONGRUITA_CERTIFICATI_NORMALIZZATE = {
    "TEMPERATURA": {
        "modelli_campione_congrui": ["CALIBR. TEMPERATURA", "MULTIMETRO DIGITALE", "DECADE DI RESISTENZA", "CALIBRATORE DI LOOP"],
        "modelli_campione_incongrui": ["MANOMETRO DIGITALE", "COMPARATORE"],
        "sottotipi_l9": {
            "TERMOCOPPIA": ["CALIBR. TEMPERATURA"],
            "TERMORESISTENZA": ["DECADE DI RESISTENZA", "CALIBR. TEMPERATURA"],
            "CONVERTITORE": ["CALIBR. TEMPERATURA", "MULTIMETRO DIGITALE", "CALIBRATORE DI LOOP"]
        }
    },
    "PRESSIONE": {
        "modelli_campione_congrui": ["MANOMETRO DIGITALE", "COMPARATORE", "CALIBRATORE DI LOOP", "MULTIMETRO DIGITALE"],
        "modelli_campione_incongrui": ["CALIBR. TEMPERATURA", "TERMOCOPPIA CAMPIONE", "TERMORESISTENZA CAMPIONE"]
    },
    "PORTATA": {
        "modelli_campione_congrui": ["MANOMETRO DIGITALE", "MULTIMETRO DIGITALE", "CALIBRATORE DI LOOP"],
        "modelli_campione_incongrui": ["CALIBR. TEMPERATURA", "TERMOCOPPIA CAMPIONE", "TERMORESISTENZA CAMPIONE"]
    },
    "LIVELLO": {
        "modelli_campione_congrui": ["MULTIMETRO DIGITALE", "CALIBRATORE DI LOOP", "COMPARATORE", "MANOMETRO DIGITALE"],
        "modelli_campione_incongrui": ["CALIBR. TEMPERATURA", "TERMOCOPPIA CAMPIONE", "TERMORESISTENZA CAMPIONE"]
    }
}

# Normalizza le regole
for regole in REGOLE_CONGRUITA_CERTIFICATI_NORMALIZZATE.values():
    if "modelli_campione_congrui" in regole:
        regole["modelli_campione_congrui"] = [m.strip().upper() for m in regole["modelli_campione_congrui"]]
    if "modelli_campione_incongrui" in regole:
        regole["modelli_campione_incongrui"] = [m.strip().upper() for m in regole["modelli_campione_incongrui"]]
    if "sottotipi_l9" in regole:
        for sottotipo, modelli in regole["sottotipi_l9"].items():
            regole["sottotipi_l9"][sottotipo] = [m.strip().upper() for m in modelli]
    if "eccezioni_l9_incongrui" in regole:
        for eccezione, modelli in regole["eccezioni_l9_incongrui"].items():
            regole["eccezioni_l9_incongrui"][eccezione] = [m.strip().upper() for m in modelli]

LISTA_UM_PRESSIONE_RICONOSCIUTE = sorted([
    "bar", "barg", "bara", "mbar", "mbarg", "mbara",
    "pa", "kpa", "mpa", "psi", "psig", "psia",
    "mmh2o", "cmh2o", "mh2o", "mmhg", "cmhg", "mhg", "kg/cm2"
])

MAPPA_NORMALIZZAZIONE_UM = {
    "mm h2o": "mmh2o", "mmh2o": "mmh2o", "mm H2O": "mmh2o",
    "kg/cm2": "kg/cm2", "kg/cm^2": "kg/cm2",
    "milliampere": "ma", "milli ampere": "ma", "milliamperes": "ma", "mamp": "ma",
    "percent": "%", "percentage": "%"
}

RANGE_0_100_NORMALIZZATO = "0-100"
RANGE_4_20_NORMALIZZATO = "4-20"
UM_MA_NORMALIZZATA = "ma"
UM_PERCENTO_NORMALIZZATA = "%"
UM_MMH2O_NORMALIZZATA = "mmh2o"
UM_MM_NORMALIZZATA = "mm"
UM_PSI_NORMALIZZATA = "psi"

# Messaggi errore
human_error_messages_map_descriptive = {
    KEY_TIPO_SCHEDA_SCONOSCIUTO: "Tipo scheda non riconosciuto.",
    KEY_FORMULA_ERROR: "La cella contiene un errore di formula (#N/A, #VALORE!).",
    KEY_CELL_RANGE_UM_NON_LEGGIBILE: "Impossibile leggere una o piu celle di Range/UM.",
    KEY_COMP_ANA_ODC_MANCANTE: "ODC mancante.",
    KEY_COMP_ANA_DATA_COMP_MANCANTE: "Data compilazione mancante.",
    KEY_COMP_ANA_PDL_MANCANTE: "Numero PDL mancante.",
    KEY_COMP_ANA_ESECUTORE_MANCANTE: "Esecutore mancante.",
    KEY_COMP_ANA_SUPERVISORE_MANCANTE: "Supervisore ISAB mancante.",
    KEY_COMP_ANA_CONTRATTO_MANCANTE: "Contratto Coemi mancante.",
    KEY_COMP_DIG_ODC_MANCANTE: "ODC mancante.",
    KEY_COMP_DIG_DATA_COMP_MANCANTE: "Data compilazione mancante.",
    KEY_COMP_DIG_PDL_MANCANTE: "Numero PDL mancante.",
    KEY_COMP_DIG_ESECUTORE_MANCANTE: "Esecutore mancante.",
    KEY_COMP_DIG_SUPERVISORE_MANCANTE: "Supervisore ISAB mancante.",
    KEY_COMP_DIG_CONTRATTO_MANCANTE: "Contratto Coemi mancante.",
    KEY_COMP_ANA_CONTRATTO_DIVERSO: f"Contratto Coemi non valido. Atteso: '{VALORE_ATTESO_CONTRATTO_COEMI}'.",
    KEY_COMP_DIG_CONTRATTO_DIVERSO: f"Contratto Coemi non valido. Atteso: '{VALORE_ATTESO_CONTRATTO_COEMI}'.",
    KEY_SP_VUOTO: "Codice SP (Tipologia Strumento) mancante.",
    KEY_L9_VUOTO: "Modello Strumento mancante.",
    KEY_L9_SKINPOINT_INCOMPLETO: "Modello 'SKIN POINT' incompleto (manca tipo K, J).",
    KEY_ERR_ANA_TEMP_CONV_C9F9_UM_DIVERSE: "Temp./Convertitore: UM Ingresso e UM DCS devono coincidere.",
    KEY_ERR_ANA_TEMP_CONV_F12_UM_NON_MA: f"Temp./Convertitore: UM Uscita deve essere '{UM_MA_NORMALIZZATA}'.",
    KEY_ERR_ANA_TEMP_CONV_A9D9_RANGE_DIVERSI: "Temp./Convertitore: Range Ingresso e Range DCS devono coincidere.",
    KEY_ERR_ANA_TEMP_CONV_D12_RANGE_NON_4_20: f"Temp./Convertitore: Range Uscita deve essere '{RANGE_4_20_NORMALIZZATO}'.",
    KEY_ERR_ANA_TEMP_NOCONV_UM_NON_COINCIDENTI: "Temperatura: UM Ingresso, DCS e Uscita devono coincidere.",
    KEY_ERR_ANA_TEMP_NOCONV_RANGE_NON_COINCIDENTI: "Temperatura: Range Ingresso, DCS e Uscita devono coincidere.",
    KEY_ERR_DIG_PRESS_D22_UM_NON_PRESSIONE: "UM Processo non e un'unita di pressione valida.",
    KEY_ERR_DIG_LIVELLO_D22_UM_NON_PERCENTO: f"UM Processo deve essere '{UM_PERCENTO_NORMALIZZATA}'.",
}


# Dataclass per strumenti campione (usata in altri moduli)
class CalibrationStandard:
    """Rappresenta uno strumento campione."""
    def __init__(self, modello_strumento, id_certificato, range, scadenza, scadenza_raw, data_emissione):
        self.modello_strumento = modello_strumento
        self.id_certificato = id_certificato
        self.range = range
        self.scadenza = scadenza
        self.scadenza_raw = scadenza_raw
        self.data_emissione = data_emissione


def excel_coord_to_indices(coord_str: str) -> tuple:
    """Converte coordinate Excel (es. 'A1') in indici (riga, colonna) 0-based."""
    match = re.match(r"([A-Z]+)([0-9]+)", coord_str.upper())
    if not match:
        raise ValueError(f"Coordinata Excel non valida: {coord_str}")

    col_str, row_str = match.groups()
    col_idx = 0
    for i, char in enumerate(reversed(col_str)):
        col_idx += (ord(char) - ord('A') + 1) * (26 ** i)
    return int(row_s) - 1, col_idx - 1


def is_config_valid() -> bool:
    """Controlla se la configurazione attuale è valida per l'analisi."""
    if not FILE_REGISTRO_STRUMENTI or not os.path.exists(FILE_REGISTRO_STRUMENTI):
        return False
    if not FOLDER_PATH_DEFAULT or not os.path.exists(FOLDER_PATH_DEFAULT):
        return False
    return True


def save_config(new_config: dict) -> bool:
    """Salva la configurazione su file JSON."""
    try:
        with open(CONFIG_FILE_PATH, 'w') as f:
            json.dump(new_config, f, indent=4)
        return True
    except Exception as e:
        logger.error(f"Errore salvataggio config: {e}")
        return False


def load_config_from_json():
    """Carica la configurazione da file JSON."""
    global FILE_REGISTRO_STRUMENTI, FOLDER_PATH_DEFAULT, FILE_DATI_COMPILAZIONE_SCHEDE
    global FILE_MASTER_DIGITALE_XLSX, FILE_MASTER_ANALOGICO_XLSX, VALIDATION_RULES

    if os.path.exists(CONFIG_FILE_PATH):
        try:
            with open(CONFIG_FILE_PATH, 'r') as f:
                data = json.load(f)
                FILE_REGISTRO_STRUMENTI = data.get("FILE_REGISTRO_STRUMENTI")
                FOLDER_PATH_DEFAULT = data.get("FOLDER_PATH_DEFAULT")
                FILE_DATI_COMPILAZIONE_SCHEDE = data.get("FILE_DATI_COMPILAZIONE_SCHEDE")
                FILE_MASTER_DIGITALE_XLSX = data.get("FILE_MASTER_DIGITALE_XLSX")
                FILE_MASTER_ANALOGICO_XLSX = data.get("FILE_MASTER_ANALOGICO_XLSX")
                VALIDATION_RULES = data.get("VALIDATION_RULES", [])
        except Exception as e:
            logger.error(f"Errore caricamento config: {e}")

# Caricamento iniziale
load_config_from_json()
