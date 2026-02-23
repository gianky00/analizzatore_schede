from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class CalibrationStandard:
    """Rappresenta uno strumento campione letto dal registro."""
    modello_strumento: str
    id_certificato: str
    range: str
    scadenza: datetime | None
    scadenza_raw: str
    data_emissione: datetime | None

@dataclass
class CertificateUsage:
    """Rappresenta un singolo utilizzo di un certificato su una scheda."""
    file_name: str
    file_path: str
    card_type: str | None
    card_date: datetime | None
    certificate_id: str
    certificate_expiry_raw: str
    certificate_expiry: datetime | None
    instrument_model_on_card: str
    instrument_range_on_card: str
    is_expired_at_use: bool
    tipologia_strumento_scheda: str
    modello_L9_scheda: str
    modello_strumento_campione_usato: str
    is_congruent: bool | None
    congruency_notes: str
    used_before_emission: bool

@dataclass
class CompilationData:
    """Dati raccolti da una scheda per la successiva compilazione automatica."""
    file_path: str
    base_filename: str
    file_type: str | None
    campi_mancanti: set[str] = field(default_factory=set)
    pdl_val: str | None = None
    odc_val_scheda: str | None = None

@dataclass
class SheetError:
    """Rappresenta un singolo errore di compilazione trovato in una scheda."""
    key: str
    description: str
    cell: str | None = None
    suggestion: str | None = None

@dataclass
class InstrumentSheet:
    """Rappresenta il risultato completo dell'analisi di un file di scheda."""
    file_path: str
    base_filename: str
    status: str
    is_valid: bool
    card_date: datetime | None = None
    file_type: str | None = None
    tipologia_strumento: str | None = None
    modello_l9: str | None = None
    certificate_usages: list[CertificateUsage] = field(default_factory=list)
    human_errors: list[SheetError] = field(default_factory=list)
    compilation_data: CompilationData | None = None
