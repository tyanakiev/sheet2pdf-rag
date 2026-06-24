# Welcome to sheet2pdf-rag

**sheet2pdf-rag** is a production-grade Python library designed to bridge the gap between complex financial Excel files and AI Retrieval-Augmented Generation (RAG) systems.

When building RAG systems on financial documents, standard spreadsheet parsers only extract cell data. They completely ignore:
- Floating text boxes
- Embedded images and guidelines
- Complex workflow drawings and flowcharts

**sheet2pdf-rag** solves this by orchestrating a headless conversion of Excel sheets into high-fidelity PDFs, and natively injecting them into your LangChain or Google Cloud Vertex AI pipelines.

## Installation

```bash
pip install sheet2pdf-rag
```

To enable native Windows Excel rendering:
```bash
pip install sheet2pdf-rag[windows]
```
