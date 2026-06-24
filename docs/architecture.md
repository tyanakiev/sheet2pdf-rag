# Architecture & Tradeoffs

`sheet2pdf-rag` is designed with a hybrid rendering architecture to balance **fidelity** with **scalability**.

## The Rendering Engines

### 1. `win32` (Microsoft Excel COM)
- **Fidelity**: 100% Exact. 
- **Requirements**: Windows OS, Microsoft Excel installed.
- **Use Case**: Best used on local developer workstations or Windows Server environments. It opens the raw Excel executable in headless mode and uses Microsoft's native layout engine to print the PDFs.

### 2. `libreoffice` (Headless Calc)
- **Fidelity**: ~95%. Excellent for tables, text, and basic charts. Sometimes struggles with highly complex SmartArt.
- **Requirements**: Cross-platform (Linux, macOS, Docker). Requires `libreoffice` or `soffice` installed.
- **Use Case**: Standard production deployment (e.g. AWS ECS, GCP Cloud Run, Docker containers). 

## Splitting Sheets
By default, `split_sheets=True` is recommended for RAG systems.
Instead of chunking a massive PDF where Page 3 might contain the end of the "Income Statement" and the beginning of the "Workflow Guidelines", splitting the sheets ensures that the LangChain chunks never cross semantic boundaries.
