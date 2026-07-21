# PDF Excel Automation Engine

> A Python automation engine that automatically monitors folders, detects multiple document types, extracts structured information from PDF and image files, enriches the data using local lookup tables, and generates standardized Excel workbooks.

---

## Overview

This project was developed to automate a repetitive document processing workflow that previously required significant manual work.

The engine continuously monitors dedicated input folders, automatically identifies the document type, extracts the required information, enriches the extracted data using local Excel lookup tables, and generates standardized Excel outputs using predefined templates.

The application was designed with performance and maintainability in mind by combining persistent caching, OCR support, template-based output generation, and automatic document classification.

---

# Features

- Automatic folder monitoring using **Watchdog**
- Automatic PDF document processing
- OCR support for scanned images using **Tesseract**
- Automatic document type detection
- Template-based Excel generation
- Local lookup table integration
- Persistent cache for improved startup performance
- Currency handling (HUF / EUR)
- Automatic formatting of generated Excel files
- Highlighting of unmatched products
- Modular processing logic for multiple document formats

---

# Supported document types

The engine currently supports multiple independent document formats:

| Document | Description |
|-----------|-------------|
| EK | Purchase price requests |
| GE | GE order documents |
| SAJATKESZLET | Internal stock documents |
| UNIMAS | Supplier order documents |

Each document type has its own parsing logic and output template.

---

# Workflow

```text
                  Input Folder
                        │
                        ▼
             PDF / Image Detection
                        │
        ┌───────────────┼────────────────┐
        │               │                │
       EK              GE            UNIMAS
        │               │                │
        └───────────────┼────────────────┘
                        │
                OCR / PDF Parsing
                        │
                        ▼
                 Lookup Tables
                        │
                        ▼
                Excel Templates
                        │
                        ▼
              Generated Excel File
                        │
                        ▼
                   Output Folder
```

---

# Project structure

```text
PDF-Excel-Automation-Engine
│
├── Engine.py
├── Templates/
├── Input/
├── Input_EK/
├── Output/
├── Cache/
├── requirements.txt
└── README.md
```

---

# Technologies

- Python
- OpenPyXL
- PDFPlumber
- Pillow
- PyTesseract
- Watchdog

---

# Installation

Clone the repository

```bash
git clone https://github.com/LFodor91/PDF-Excel-Automation-Engine.git
```

Install dependencies

```bash
pip install -r requirements.txt
```

Run the engine

```bash
python Engine.py
```

---

# Configuration

The engine expects several local files that are intentionally **not included** in this repository.

These include:

- `lookup.xlsx`
- `Vevők.xlsx`

These files contain company-specific master data and therefore have been excluded from the public repository.

The Excel templates included in the repository demonstrate the required output structure.

---

# Security

This repository intentionally excludes:

- Customer data
- Internal pricing
- Company lookup tables
- Supplier documents
- Generated outputs
- OCR cache files

The published version contains only the automation logic.

---

# Performance

The engine uses:

- Persistent lookup cache
- Optimized Excel loading
- Automatic file readiness detection
- Incremental processing
- Lightweight folder monitoring

to minimize startup time and processing latency.

---

# Future Improvements

- Configuration file support
- Batch processing mode
- Logging improvements
- Unit testing
- Docker support
- GUI version
- Plugin architecture for new document types

---

# Author

**László Fodor**

If you found this project interesting, feel free to connect with me on LinkedIn.
