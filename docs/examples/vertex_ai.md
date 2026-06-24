# GCP Vertex AI Integration

For complex workflow drawings, pipe the rendered sheets directly into Gemini 1.5 Pro.

```python
from sheet2pdf_rag import analyze_pdf_with_gemini

results = analyze_pdf_with_gemini(
    excel_path="guidelines.xlsx",
    prompt="Extract the core workflow topologies and decision trees shown in the images.",
    model_name="gemini-1.5-pro-001",
    project="your-gcp-project-id"
)

for res in results:
    print(f"Sheet: {res['sheet_name']}")
    print(f"Analysis: {res['gemini_analysis']}")
```
