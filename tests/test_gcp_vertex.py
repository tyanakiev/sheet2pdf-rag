import pytest
from sheet2pdf_rag import analyze_pdf_with_gemini
from sheet2pdf_rag.core import ConversionResult

def test_gcp_vertex_payload_building(mocker, tmp_path):
    """Test that the GCP integration builds the correct payload for Gemini."""
    
    # Mock the Vertex AI library completely
    mock_vertexai = mocker.patch("sheet2pdf_rag.gcp_vertex.vertexai")
    mock_genmodel = mocker.patch("sheet2pdf_rag.gcp_vertex.GenerativeModel")
    mock_part = mocker.patch("sheet2pdf_rag.gcp_vertex.Part")
    
    # Mock the conversion engine
    mock_converter = mocker.patch("sheet2pdf_rag.gcp_vertex.ExcelConverter")
    mock_instance = mock_converter.return_value
    
    # Create a real dummy PDF file to satisfy the 'with open(pdf_path, "rb")' logic
    dummy_pdf = tmp_path / "dummy.pdf"
    dummy_pdf.write_bytes(b"dummy pdf binary data")
    
    mock_instance.convert.return_value = [
        ConversionResult(
            original_excel_path="/fake/path.xlsx",
            pdf_path=str(dummy_pdf),
            sheet_name="Workflow",
            engine_used="mock"
        )
    ]
    
    # Mock the GenerativeModel response
    mock_model_instance = mock_genmodel.return_value
    mock_model_instance.generate_content.return_value.text = "Extracted topologies here."
    
    results = analyze_pdf_with_gemini(
        excel_path="/fake/path.xlsx",
        project="test-project"
    )
    
    assert len(results) == 1
    assert results[0]["sheet_name"] == "Workflow"
    assert results[0]["gemini_analysis"] == "Extracted topologies here."
    
    # Ensure vertexai was initialized with the project
    mock_vertexai.init.assert_called_with(project="test-project", location="us-central1")
