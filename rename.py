import os

files_to_update = [
    "README.md",
    "src/sheet2pdf_rag/__init__.py",
]

for file_path in files_to_update:
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()
        
        # Replace occurrences
        new_content = content.replace("excel2pdf-rag", "sheet2pdf-rag")
        new_content = new_content.replace("excel2pdf_rag", "sheet2pdf_rag")
        new_content = new_content.replace("ExcelPDFLoader", "SheetPDFLoader")
        new_content = new_content.replace("excel_path", "sheet_path")
        
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(new_content)

print("Renaming complete!")
