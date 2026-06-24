import sys
from pathlib import Path

# Add src to python path for testing
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from sheet2pdf_rag import ExcelConverter, SheetPDFLoader
from sheet2pdf_rag.gcp_vertex import analyze_pdf_with_gemini

def test_imports():
    print("Testing ExcelConverter instantiation...")
    converter = ExcelConverter(engine="auto")
    print(f"Auto-selected engine: {converter._engine}")
    
    print("\nTesting LangChain Loader structure...")
    loader = SheetPDFLoader("dummy.xlsx", split_sheets=True)
    print(f"Loader engine selected: {loader.converter._engine}")
    print("Imports and class structures are valid.")

if __name__ == "__main__":
    test_imports()
