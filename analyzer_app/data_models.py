from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional, Set

@dataclass
class CalibrationStandard:
    """Rappresenta uno strumento campione letto dal registro."""
    modello_strumento: str
    id_certificato: str
    range: str
    scadenza: Optional[datetime]
    scadenza_raw: str
    data_emissione: Optional[datetime]

@dataclass
class CertificateUsage:
    """Rappresenta un singolo utilizzo di un certificato su una scheda."""
    file_name: str
    file_path: str
    card_type: Optional[str]
    card_date: Optional[datetime]
    certificate_id: str
    certificate_expiry_raw: str
    certificate_expiry: Optional[datetime]
    instrument_model_on_card: str
    instrument_range_on_card: str
    is_expired_at_use: bool
    tipologia_strumento_scheda: str
    modello_L9_scheda: str
    modello_strumento_campione_usato: str
    is_congruent: Optional[bool]
    congruency_notes: str
    used_before_emission: bool

@dataclass
class CompilationData:
    """Dati raccolti da una scheda per la successiva compilazione automatica."""
    file_path: str
    base_filename: str
    file_type: Optional[str]
    campi_mancanti: Set[str] = field(default_factory=set)
    pdl_val: Optional[str] = None
    odc_val_scheda: Optional[str] = None

@dataclass
class SheetError:
    """Rappresenta un singolo errore di compilazione trovato in una scheda."""
    key: str
    description: str
    cell: Optional[str] = None
    suggestion: Optional[str] = None

@dataclass
class InstrumentSheet:
    """Rappresenta il risultato completo dell'analisi di un file di scheda."""
    file_path: str
    base_filename: str
    status: str
    is_valid: bool
    card_date: Optional[datetime] = None
    file_type: Optional[str] = None
    tipologia_strumento: Optional[str] = None
    modello_l9: Optional[str] = None
    certificate_usages: List[CertificateUsage] = field(default_factory=list)
    human_errors: List[SheetError] = field(default_factory=list)
    compilation_data: Optional[CompilationData] = None
