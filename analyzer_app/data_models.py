from dataclasses import dataclass, field
from datetime import datetime

@dataclass
class CertificateUsage:
    """Rappresenta l'utilizzo di un certificato in una scheda."""
    file_name: str
    file_path: str
    card_type: str  # 'analogico' o 'digitale'
    card_date: datetime | None
    certificate_id: str
    certificate_expiry_raw: str
    certificate_expiry: datetime | None
    instrument_model_on_card: str
    instrument_range_on_card: str
    is_expired_at_use: bool
    tipologia_strumento_scheda: str
    modello_l9_scheda: str
    modello_strumento_campione_usato: str
    is_congruent: bool | None  # True=Congruo, False=Non congruo, None=Non verificabile
    congruency_notes: str
    used_before_emission: bool = False

@dataclass
class SheetError:
    """Rappresenta un errore trovato nella scheda."""
    key: str
    description: str
    cell: str | None = None
    suggestion: str | None = None

@dataclass
class CompilationData:
    """Dati di compilazione estratti dalla scheda."""
    odc_val_scheda: str | None
    pdl_val: str | None

@dataclass
class InstrumentSheet:
    """Rappresenta l'esito dell'analisi di una singola scheda."""
    file_path: str
    base_filename: str
    status: str
    is_valid: bool
    human_errors: list[SheetError] = field(default_factory=list)
    certificate_usages: list[CertificateUsage] = field(default_factory=list)
    compilation_data: CompilationData | None = None
