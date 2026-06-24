import pytest
import openpyxl
from pathlib import Path

@pytest.fixture
def dummy_excel_path(tmp_path):
    """Creates a dummy Excel file with multiple sheets for testing."""
    file_path = tmp_path / "dummy_test.xlsx"
    
    wb = openpyxl.Workbook()
    
    # Sheet 1
    ws1 = wb.active
    ws1.title = "Financials"
    ws1["A1"] = "Revenue"
    ws1["B1"] = 1000
    
    # Sheet 2
    ws2 = wb.create_sheet(title="Guidelines")
    ws2["A1"] = "Important Workflow Rules"
    
    wb.save(file_path)
    return file_path
