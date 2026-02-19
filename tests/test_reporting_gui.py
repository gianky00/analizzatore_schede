import sys
from datetime import datetime
from unittest.mock import MagicMock, patch

import pytest

sys.modules['tkinter'] = MagicMock()
sys.modules['tkinter.ttk'] = MagicMock()
sys.modules['tkinter.filedialog'] = MagicMock()
sys.modules['tkinter.messagebox'] = MagicMock()
sys.modules['tkinter.font'] = MagicMock()

from analyzer_app import data_models, gui, reporting  # noqa: E402


@pytest.fixture
def mock_error_dicts():
    return [
        {"key": "ERR_TEST", "description": "Descrizione", "cell": "A1", "suggestion": "Sugg", "file": "test.xlsx", "path": "C:/tmp/test.xlsx"},
        {"key": "COMP_TEST", "description": "Compilazione", "cell": "B2", "suggestion": "", "file": "test.xlsx", "path": "C:/tmp/test.xlsx"}
    ]

def test_reporting_word_generation(mock_error_dicts):
    with patch("analyzer_app.reporting.Document") as mock_doc_cls:
        mock_doc = mock_doc_cls.return_value
        with patch("subprocess.Popen"):
            reporting.crea_e_apri_report_anomalie_word(
                errors_list=mock_error_dicts,
                temporal_list=[],
                incongruent_list=[],
                candidate_files_count=10,
                validated_file_count=5
            )
        mock_doc.save.assert_called()

def test_gui_init_state():
    app = gui.App(MagicMock())
    assert hasattr(app, 'analysis_queue')
    assert hasattr(app, 'cert_details_map')
    assert hasattr(app, 'analysis_service')

def test_gui_start_analysis_logic():
    """Verifica che start_analysis deleghi correttamente al servizio."""
    app = gui.App(MagicMock())
    # Mocking config and registry reading
    with (
        patch("analyzer_app.config.is_config_valid", return_value=True),
        patch("analyzer_app.excel_io.leggi_registro_strumenti", return_value=[])
    ):
        # Mocking the service method
        app.analysis_service.start_analysis = MagicMock()
        app.start_analysis()
        app.analysis_service.start_analysis.assert_called_once()

def test_gui_update_cert_details_map_logic():
    app = gui.App(MagicMock())
    usage = data_models.CertificateUsage(
        file_name="f", file_path="p", card_type="a", card_date=datetime(2026, 1, 1),
        certificate_id="C1", certificate_expiry_raw="", certificate_expiry=datetime(2025, 1, 1),
        instrument_model_on_card="M1", instrument_range_on_card="R1",
        is_expired_at_use=True,
        tipologia_strumento_scheda="PRESSIONE", modello_l9_scheda="DP",
        modello_strumento_campione_usato="M1", is_congruent=True, congruency_notes="",
        used_before_emission=False
    )
    sheet = data_models.InstrumentSheet(
        file_path="p", base_filename="f", status="", is_valid=True,
        certificate_usages=[usage]
    )
    app.analysis_results = [sheet]
    app._update_cert_details_map()
    assert app.cert_details_map is not None
