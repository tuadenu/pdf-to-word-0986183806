# PDF to Word Converter

A simple web application that converts PDF files to editable Word (`.docx`) documents.

## Features

- Drag-and-drop or click-to-browse PDF upload
- Fast, server-side conversion powered by [pdf2docx](https://github.com/dothinking/pdf2docx)
- Automatic download of the converted `.docx` file
- Uploaded files are deleted from the server after conversion

## Getting Started

### Prerequisites

- Python 3.8+

### Installation

```bash
pip install -r requirements.txt
```

### Running the app

```bash
python app.py
```

Then open <http://localhost:5000> in your browser.

## Usage

1. Open the app in your browser.
2. Drag and drop a PDF file onto the upload area, or click **browse** to select one.
3. Click **Convert to Word**.
4. The converted `.docx` file will download automatically.

## Running Tests

```bash
python -m pytest tests/ -v
```
