"""
sheet2pdf_rag: A production-grade Python library to convert complex Excel files to PDFs for RAG systems.
"""

__version__ = "0.1.0"

from .core import ExcelConverter, ConversionResult
from .langchain_loader import SheetPDFLoader
from .gcp_vertex import analyze_pdf_with_gemini

__all__ = [
    "ExcelConverter",
    "ConversionResult",
    "SheetPDFLoader",
    "analyze_pdf_with_gemini",
]
