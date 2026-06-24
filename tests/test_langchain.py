import pytest
from pathlib import Path
from langchain_core.documents import Document
from sheet2pdf_rag import SheetPDFLoader
from sheet2pdf_rag.core import ConversionResult

def test_langchain_metadata_injection(mocker):
    """Test that the SheetPDFLoader correctly maps conversion results into LangChain Document metadata."""
    
    # We mock the ExcelConverter so we don't actually generate PDFs
    mock_converter = mocker.patch("sheet2pdf_rag.langchain_loader.ExcelConverter")
    mock_instance = mock_converter.return_value
    
    mock_instance.convert.return_value = [
        ConversionResult(
            original_excel_path="/fake/path/data.xlsx",
            pdf_path="/tmp/fake_data__Sheet1.pdf",
            sheet_name="Sheet1",
            engine_used="mocked_engine"
        )
    ]
    
    loader = SheetPDFLoader("/fake/path/data.xlsx")
    
    # We also mock PyPDFLoader to just return a dummy Document instead of parsing a real PDF
    mock_pypdf = mocker.patch("langchain_community.document_loaders.PyPDFLoader")
    mock_pypdf_instance = mock_pypdf.return_value
    
    # PyPDFLoader's lazy_load returns an iterator of Documents
    dummy_doc = Document(page_content="Mocked PDF Content", metadata={"source": "/tmp/fake_data__Sheet1.pdf", "page": 0})
    mock_pypdf_instance.lazy_load.return_value = iter([dummy_doc])
    
    # Consume the loader
    docs = list(loader.lazy_load())
    
    assert len(docs) == 1
    doc = docs[0]
    
    # Verify metadata was perfectly mapped!
    assert doc.metadata["source"] == "/fake/path/data.xlsx"
    assert doc.metadata["sheet_name"] == "Sheet1"
    assert doc.metadata["engine_used"] == "mocked_engine"
    assert doc.metadata["is_from_excel"] is True
    assert doc.metadata["page"] == 0
