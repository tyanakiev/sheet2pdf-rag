import sys
from pathlib import Path

# Add src to python path for testing without installing
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from sheet2pdf_rag import analyze_pdf_with_gemini

def main():
    print("Sending Excel workbook directly to Gemini 1.5 Pro...")
    
    try:
        # Requires Google Cloud Authentication (e.g. `gcloud auth application-default login`)
        # Replace 'sample_guidelines.xlsx' and 'your-gcp-project-id' with your details!
        results = analyze_pdf_with_gemini(
            excel_path="sample_guidelines.xlsx",
            prompt="Extract the core workflow topologies and decision trees shown in the images. Output as structured text.",
            split_sheets=True,
            model_name="gemini-1.5-pro-001",
            project="your-gcp-project-id"  # Set your GCP Project ID here
        )
        
        for i, res in enumerate(results):
            print(f"\n--- Sheet Result: {res['sheet_name']} ---")
            print(res["gemini_analysis"])
            
    except FileNotFoundError:
        print("Error: Please provide a valid Excel file path in the script.")
    except Exception as e:
        print(f"Failed to call Gemini: {e}")

if __name__ == "__main__":
    main()
