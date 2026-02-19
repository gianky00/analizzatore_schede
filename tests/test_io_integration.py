import multiprocessing
import os
from datetime import datetime
from unittest.mock import MagicMock, patch

import pytest
from openpyxl import Workbook

from analyzer_app import config, excel_io, services


@pytest.fixture
def mock_registry_file(tmp_path):
    f = tmp_path / "registro.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = config.REGISTRO_FOGLIO_NOME
    start_row = config.REGISTRO_RIGA_INIZIO_DATI
    ws.cell(row=start_row, column=config.REGISTRO_COL_IDX_MODELLO_STRUM_CAMPIONE + 1).value = "Modello Test"
    ws.cell(row=start_row, column=config.REGISTRO_COL_IDX_ID_CERT_CAMPIONE + 1).value = "CERT-123"
    ws.cell(row=start_row, column=config.REGISTRO_COL_IDX_RANGE_CAMPIONE + 1).value = "0-10 bar"
    ws.cell(row=start_row, column=config.REGISTRO_COL_IDX_SCADENZA_CAMPIONE + 1).value = "31/12/2026"
    wb.save(f)
    return str(f)

def test_leggi_registro_strumenti_success(mock_registry_file, monkeypatch):
    monkeypatch.setattr(config, "FILE_REGISTRO_STRUMENTI", mock_registry_file)
    res = excel_io.leggi_registro_strumenti()
    assert res is not None
    assert len(res) >= 1
    assert res[0].id_certificato == "CERT-123"

def test_leggi_registro_non_esistente(monkeypatch):
    monkeypatch.setattr(config, "FILE_REGISTRO_STRUMENTI", "non_esiste.xlsx")
    res = excel_io.leggi_registro_strumenti()
    assert res is None

def test_read_sheet_file_not_found():
    with pytest.raises(FileNotFoundError):
        excel_io.read_instrument_sheet_raw_data("missing.xlsx")

def test_read_instrument_sheet_raw_data_xls_legacy():
    mock_book = MagicMock()
    mock_sheet = MagicMock()
    mock_book.sheet_by_index.return_value = mock_sheet
    mock_sheet.merged_cells = []

    r_e2, c_e2 = excel_io.excel_coord_to_indices("E2")
    r_odc, c_odc = excel_io.excel_coord_to_indices(config.SCHEDA_ANA_CELL_ODC)

    def mock_cell_value(r, c):
        if r == r_e2 and c == c_e2:
            return "STRUMENTI ANALOGICI"
        if r == r_odc and c == c_odc:
            return "ODC-XLS"
        return None

    mock_sheet.cell_value.side_effect = mock_cell_value

    with patch("xlrd.open_workbook", return_value=mock_book):
        data = excel_io.read_instrument_sheet_raw_data("test.xls")
        assert data['base_filename'] == "test.xls"
        assert data['file_type'] == "analogico"
        assert data['odc'] == "ODC-XLS"

def test_read_instrument_sheet_raw_data_xlsx(tmp_path):
    f = tmp_path / "test.xlsx"
    wb = Workbook()
    ws = wb.active
    ws['E2'] = "SCHEDA TARATURA STRUMENTI ANALOGICI"
    ws[config.SCHEDA_ANA_CELL_ODC] = "ODC-OK"
    wb.save(f)

    data = excel_io.read_instrument_sheet_raw_data(str(f))
    assert data['file_type'] == "analogico"
    assert data['odc'] == "ODC-OK"

def test_parse_date_robust():
    assert isinstance(excel_io.parse_date_robust("19/02/2026"), datetime)
    assert excel_io.parse_date_robust("GUASTO") is None

def test_excel_coord_to_indices():
    assert excel_io.excel_coord_to_indices("A1") == (0, 0)
    assert excel_io.excel_coord_to_indices("B3") == (2, 1)

def test_write_cell_xlsx(tmp_path):
    f = tmp_path / "write.xlsx"
    wb = Workbook()
    wb.save(f)
    assert excel_io.write_cell(str(f), "A1", "TEST-VALUE") is True

def test_config_save_and_load(tmp_path):
    cfg_file = tmp_path / "config.json"
    with patch("analyzer_app.config.CONFIG_FILE_PATH", str(cfg_file)):
        test_data = {"FILE_REGISTRO_STRUMENTI": "test_path"}
        assert config.save_config(test_data) is True
        config.load_config_from_json()
        assert config.FILE_REGISTRO_STRUMENTI == "test_path"

def test_is_config_valid_logic(tmp_path, monkeypatch):
    reg = tmp_path / "reg.xlsx"
    reg.touch()
    fld = tmp_path / "schede"
    fld.mkdir()
    
    monkeypatch.setattr(config, "FILE_REGISTRO_STRUMENTI", str(reg))
    monkeypatch.setattr(config, "FOLDER_PATH_DEFAULT", str(fld))
    assert config.is_config_valid() is True

def test_internal_read_file_worker_success(tmp_path):
    f = tmp_path / "simple.xlsx"
    wb = Workbook()
    wb.active['E2'] = "DIGITALE"
    wb.save(f)

    q: multiprocessing.Queue = multiprocessing.Queue()
    services._read_file_worker_internal(q, str(f))

    status, res = q.get()
    assert status == 'success'
    assert res['file_type'] == 'digitale'
