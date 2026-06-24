import tempfile
from pathlib import Path
from typing import Iterator, List, Optional, Literal

from langchain_core.documents import Document
from langchain_core.document_loaders import BaseLoader

from .core import ExcelConverter, ConversionResult

class SheetPDFLoader(BaseLoader):
    """
    A LangChain Document Loader that converts an Excel workbook to PDF(s)
    and then parses the resulting PDFs, injecting rich lineage metadata 
    (source excel path, sheet name, engine used).
    """
    
    def __init__(
        self,
        file_path: str | Path,
        split_sheets: bool = True,
        engine: Literal["auto", "win32", "libreoffice"] = "auto",
        pdf_parser: str = "pypdf"
    ):
        """
        Initialize the loader.
        
        Args:
            file_path: Path to the Excel file.
            split_sheets: If True, each sheet is extracted as a separate PDF and thus separate Documents.
            engine: The engine to use for conversion ("auto", "win32", "libreoffice").
            pdf_parser: The underlying PDF parser to use ("pypdf", or "pdfplumber").
        """
        self.file_path = Path(file_path).resolve()
        self.split_sheets = split_sheets
        self.converter = ExcelConverter(engine=engine)
        self.pdf_parser = pdf_parser

    def lazy_load(self) -> Iterator[Document]:
        """Lazy load the documents from the Excel file by converting to PDF first."""
        
        # We use a temporary directory for the generated PDFs
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir_path = Path(temp_dir)
            
            # Convert Excel to PDF(s)
            conversion_results = self.converter.convert(
                excel_path=self.file_path,
                output_dir=temp_dir_path,
                split_sheets=self.split_sheets
            )
            
            for result in conversion_results:
                yield from self._parse_pdf(result)
                
    def _parse_pdf(self, result: ConversionResult) -> Iterator[Document]:
        """Internal helper to parse a single generated PDF and inject metadata."""
        if self.pdf_parser == "pypdf":
            try:
                from langchain_community.document_loaders import PyPDFLoader
                loader = PyPDFLoader(result.pdf_path)
                docs = loader.lazy_load()
            except ImportError:
                raise ImportError("Please install langchain-community and pypdf: pip install langchain-community pypdf")
        else:
            raise ValueError(f"Unsupported pdf_parser: {self.pdf_parser}")

        for doc in docs:
            # Augment the metadata
            doc.metadata["source"] = result.original_excel_path
            doc.metadata["sheet_name"] = result.sheet_name
            doc.metadata["engine_used"] = result.engine_used
            doc.metadata["is_from_excel"] = True
            
            # Remove the temporary PDF path from metadata to avoid confusion
            if "source" in doc.metadata and doc.metadata["source"] == result.pdf_path:
                doc.metadata["source"] = result.original_excel_path
                
            yield doc
