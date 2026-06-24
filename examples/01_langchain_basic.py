import sys
from pathlib import Path

# Add src to python path for testing without installing
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from sheet2pdf_rag import SheetPDFLoader

def main():
    print("Initializing LangChain Excel Loader...")
    
    # We assume 'sample_financial.xlsx' exists in your path.
    # Replace this with a real file when testing!
    loader = SheetPDFLoader("sample_financial.xlsx", split_sheets=True)
    
    try:
        documents = loader.load()
        
        for i, doc in enumerate(documents):
            print(f"\n--- Document {i+1} ---")
            print(f"Content snippet: {doc.page_content[:150]}...")
            print(f"Source file:     {doc.metadata['source']}")
            print(f"Sheet Name:      {doc.metadata['sheet_name']}")
            print(f"Engine Used:     {doc.metadata['engine_used']}")
            
    except FileNotFoundError:
        print("Error: Please provide a valid Excel file path in the script to test.")

if __name__ == "__main__":
    main()
