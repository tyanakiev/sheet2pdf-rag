# LangChain Integration

The library provides a native LangChain loader that converts sheets to PDFs and extracts metadata automatically.

```python
from sheet2pdf_rag import SheetPDFLoader

# Initialize the loader
loader = SheetPDFLoader("financial_report.xlsx", split_sheets=True)

# Load the documents
documents = loader.load()

for doc in documents:
    print(f"Content: {doc.page_content[:150]}")
    print(f"Source Path:  {doc.metadata['source']}")
    print(f"Sheet Name:   {doc.metadata['sheet_name']}")
```
