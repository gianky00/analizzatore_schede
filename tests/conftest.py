import pytest
from datetime import datetime
from analyzer_app.data_models import CalibrationStandard

@pytest.fixture
def sample_standard():
    return CalibrationStandard(
        modello_strumento="TEST-MODEL",
        id_certificato="CERT-123",
        range="0-100 bar",
        scadenza=datetime(2025, 12, 31),
        scadenza_raw="31/12/2025",
        data_emissione=datetime(2024, 1, 1)
    )
