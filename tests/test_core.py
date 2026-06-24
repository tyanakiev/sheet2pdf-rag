import pytest
from pathlib import Path
from sheet2pdf_rag import ExcelConverter

def test_engine_resolution():
    converter = ExcelConverter(engine="auto")
    # Will default to win32 or libreoffice depending on environment
    assert converter._engine in ["win32", "libreoffice"]

def test_invalid_engine():
    with pytest.raises(ValueError):
        converter = ExcelConverter(engine="auto")
        converter._engine = "invalid_engine"
        converter.convert("fake.xlsx", "fake_out")

def test_file_not_found():
    converter = ExcelConverter()
    with pytest.raises(FileNotFoundError):
        converter.convert("does_not_exist.xlsx", "out_dir")

def test_libreoffice_mocked(dummy_excel_path, tmp_path, mocker):
    """Test the logic pathway for LibreOffice sheet splitting without actually running soffice."""
    converter = ExcelConverter(engine="libreoffice")
    
    # Mock shutil.which to bypass the executable check
    mocker.patch("shutil.which", return_value="/usr/bin/soffice")
    # Mock subprocess.run to not actually invoke LibreOffice
    mock_run = mocker.patch("subprocess.run")
    
    # Create dummy pdf files to simulate the output of LibreOffice
    def side_effect(*args, **kwargs):
        # args[0] is the command list. The outdir is args[0][5], and the filename generated is temp_{sheet}.pdf
        # We will touch the expected output files so the rest of the function succeeds
        for sheet in ["Financials", "Guidelines"]:
            (tmp_path / f"temp_{sheet}.pdf").touch()
    
    mock_run.side_effect = side_effect
    
    results = converter.convert(dummy_excel_path, tmp_path, split_sheets=True)
    
    assert len(results) == 2
    assert results[0].sheet_name == "Financials"
    assert results[1].sheet_name == "Guidelines"
    assert results[0].engine_used == "libreoffice"
    assert "dummy_test__Financials.pdf" in results[0].pdf_path
