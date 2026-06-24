import os
from pathlib import Path
from typing import List, Dict, Any, Optional

from .core import ExcelConverter, ConversionResult

def analyze_pdf_with_gemini(
    excel_path: str | Path,
    prompt: str = "Analyze the drawings, text, and data in this financial document. Extract the core rules, guidelines, and workflow topologies.",
    split_sheets: bool = True,
    model_name: str = "gemini-1.5-pro-001",
    project: Optional[str] = None,
    location: str = "us-central1"
) -> List[Dict[str, Any]]:
    """
    Converts an Excel file to PDF(s) and sends them to Google Cloud Vertex AI
    (Gemini 1.5 Pro) for multimodal analysis, retaining sheet lineage.
    
    Args:
        excel_path: Path to the original Excel workbook.
        prompt: The instructions to give to the Gemini model.
        split_sheets: If True, analyze each sheet individually.
        model_name: The Gemini model version to use.
        project: GCP Project ID. If None, it uses the environment default.
        location: GCP location.
        
    Returns:
        A list of dictionaries containing the sheet metadata and the model's text response.
    """
    try:
        import vertexai
        from vertexai.generative_models import GenerativeModel, Part
    except ImportError:
        raise ImportError("Please install google-cloud-aiplatform to use this feature: pip install google-cloud-aiplatform")

    vertexai.init(project=project, location=location)
    model = GenerativeModel(model_name)
    converter = ExcelConverter()
    
    # Use a temporary directory for the intermediate PDFs
    import tempfile
    
    results_output = []
    
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir_path = Path(temp_dir)
        
        conversion_results = converter.convert(
            excel_path=excel_path,
            output_dir=temp_dir_path,
            split_sheets=split_sheets
        )
        
        for result in conversion_results:
            # Load the PDF as a Part for Gemini
            with open(result.pdf_path, "rb") as f:
                pdf_data = f.read()
                
            pdf_part = Part.from_data(data=pdf_data, mime_type="application/pdf")
            
            # Call Gemini
            response = model.generate_content(
                [pdf_part, prompt]
            )
            
            results_output.append({
                "source_file": result.original_excel_path,
                "sheet_name": result.sheet_name,
                "engine_used": result.engine_used,
                "gemini_analysis": response.text
            })
            
    return results_output
