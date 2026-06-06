# File Converter

A desktop file converter app built with Python and Tkinter. Convert between common image, document, and data formats — plus merge, trim, and split PDFs — all from a simple local GUI.

## Features

- **Convert** between 20+ formats across images, documents, and data files
- **PDF Merge** — combine multiple PDFs in any order
- **PDF Trim** — extract a page range from a PDF
- **PDF Split** — split a PDF into chunks, saved as a ZIP

## Supported Conversions

| From | To |
|------|----|
| PNG, JPG, JPEG, WEBP, BMP, GIF, TIFF, ICO | Any other image format |
| PDF | TXT |
| TXT | PDF |
| DOCX, DOC | TXT |
| MD | HTML |
| EPUB | PDF |
| CSV | JSON, XML |
| JSON | CSV, XML |
| XML | JSON |

## Requirements

- Python 3.10+
- Dependencies listed in `requirements.txt`

## Setup

```bash
pip install -r requirements.txt
python app.py
```

## Build

To build a standalone Windows executable:

```bash
build.bat
```

The output will be in `dist/FileConverter.exe`.
