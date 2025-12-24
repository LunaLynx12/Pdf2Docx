# PDF to DOCX Converter

A professional Python package for converting PDF files to DOCX format with excellent format preservation. Built on top of the [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) library for high-quality conversion.

![Python Version](https://img.shields.io/badge/python-3.7+-blue.svg)
![License](https://img.shields.io/badge/license-GPL--3.0-green.svg)

## ✨ Features

- ✅ **Excellent format preservation** - Uses the professional pdf2docx library
- ✅ **Text formatting** - Preserves bold, italic, font size, color, underline, strike-through
- ✅ **Layout preservation** - Maintains paragraph structure, alignment, and spacing
- ✅ **Image support** - Extracts and embeds images (Gray/RGB/CMYK, transparent images)
- ✅ **Table preservation** - Maintains tables with borders, shading, and merged cells
- ✅ **Multi-page support** - Handles documents of any size
- ✅ **Page layout** - Preserves margins, sections, and multi-column layouts
- ✅ **Command-line interface** - Full-featured CLI with advanced options
- ✅ **Configuration system** - Fine-tuned control over conversion process
- ✅ **Logging & progress** - Comprehensive logging and progress tracking
- ✅ **Batch conversion** - Convert multiple PDFs at once
- ✅ **Error handling** - Robust validation and error reporting

## 📦 Installation

### From Source

```bash
# Clone the repository
git clone https://github.com/LunaLynx12/Pdf2Docx.git
cd Pdf2Docx

# Install dependencies
pip install -r requirements.txt

# Install package (optional, for development)
pip install -e .
```

### Using pip

```bash
pip install pdf-to-docx
```

## 🚀 Quick Start

### Command Line

```bash
# Basic conversion
python -m pdf_to_docx input.pdf

# Specify output file
python -m pdf_to_docx input.pdf output.docx

# Convert specific pages
python -m pdf_to_docx input.pdf --start-page 0 --end-page 10

# With options
python -m pdf_to_docx input.pdf output.docx --overwrite --verbose
```

### Python API

```python
from pdf_to_docx import PDFToDOCXConverter

# Simple conversion
converter = PDFToDOCXConverter("input.pdf", "output.docx")
result = converter.convert()
print(f"Converted to: {result}")
```

## 📖 Usage

### Command Line Interface

#### Basic Usage

```bash
# Convert PDF to DOCX (output will be same name with .docx extension)
python -m pdf_to_docx input.pdf

# Specify output file name
python -m pdf_to_docx input.pdf output.docx
```

#### Advanced Options

```bash
# Convert specific page range (0-based indexing)
python -m pdf_to_docx input.pdf --start-page 0 --end-page 10

# Overwrite existing file
python -m pdf_to_docx input.pdf output.docx --overwrite

# Create backup before overwriting
python -m pdf_to_docx input.pdf output.docx --overwrite --backup

# Enable verbose output
python -m pdf_to_docx input.pdf --verbose

# Quiet mode
python -m pdf_to_docx input.pdf --quiet

# Save logs to file
python -m pdf_to_docx input.pdf --log-file conversion.log

# Multi-processing (experimental)
python -m pdf_to_docx input.pdf --multi-processing --cpu-count 4
```

#### Help

```bash
python -m pdf_to_docx --help
```

### Python API

#### Basic Usage

```python
from pdf_to_docx import PDFToDOCXConverter

# Simple conversion
converter = PDFToDOCXConverter("input.pdf", "output.docx")
result = converter.convert()
print(f"Converted to: {result}")
```

#### With Configuration

```python
from pdf_to_docx import PDFToDOCXConverter, ConversionConfig

# Create custom configuration
config = ConversionConfig(
    start_page=0,
    end_page=10,  # Convert first 11 pages
    overwrite=True,
    create_backup=True,
    verbose=True
)

# Convert with configuration
converter = PDFToDOCXConverter(
    pdf_path="input.pdf",
    docx_path="output.docx",
    config=config
)
result = converter.convert()
```

#### Using Context Manager

```python
from pdf_to_docx import PDFToDOCXConverter

with PDFToDOCXConverter("input.pdf", "output.docx") as converter:
    result = converter.convert()
    print(f"Converted to: {result}")
```

#### Progress Tracking

```python
from pdf_to_docx import PDFToDOCXConverter

def progress_callback(current, total):
    percentage = (current / total) * 100
    print(f"Progress: {current}/{total} pages ({percentage:.1f}%)")

converter = PDFToDOCXConverter("input.pdf", "output.docx")
converter.set_progress_callback(progress_callback)
result = converter.convert_with_progress()
```

#### Batch Conversion

```python
from pathlib import Path
from pdf_to_docx import PDFToDOCXConverter

pdf_files = ["doc1.pdf", "doc2.pdf", "doc3.pdf"]

for pdf_file in pdf_files:
    pdf_path = Path(pdf_file)
    converter = PDFToDOCXConverter(
        pdf_path=pdf_path,
        docx_path=pdf_path.with_suffix('.docx')
    )
    try:
        result = converter.convert()
        print(f"✓ Converted: {pdf_file}")
    except Exception as e:
        print(f"✗ Failed: {pdf_file} - {e}")
```

See `examples/` directory for more usage examples.

## 📁 Project Structure

```
pdf-to-docx/
├── pdf_to_docx/          # Main package
│   ├── __init__.py      # Package initialization
│   ├── __main__.py      # Module entry point
│   ├── converter.py     # Main converter class
│   ├── config.py        # Configuration system
│   ├── utils.py         # Utility functions
│   └── cli.py           # Command-line interface
├── examples/            # Usage examples
│   ├── basic_usage.py
│   └── advanced_usage.py
├── setup.py             # Package setup
├── requirements.txt     # Dependencies
├── LICENSE              # GPL-3.0 License
├── .gitignore          # Git ignore rules
└── README.md           # This file
```

## ⚙️ Configuration

The `ConversionConfig` class provides comprehensive configuration options:

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `start_page` | int | 0 | Start page index (0-based) |
| `end_page` | int | None | End page index (None = last page) |
| `overwrite` | bool | False | Overwrite output file if exists |
| `create_backup` | bool | False | Create backup before overwriting |
| `verbose` | bool | True | Enable verbose logging |
| `log_file` | Path | None | Path to log file |
| `multi_processing` | bool | False | Enable multi-processing |
| `cpu_count` | int | None | Number of CPU cores to use |

Example:

```python
from pdf_to_docx import ConversionConfig

config = ConversionConfig(
    start_page=5,
    end_page=15,
    overwrite=True,
    create_backup=True,
    verbose=True,
    log_file=Path("conversion.log")
)
```

## 🔧 Requirements

- Python 3.7+
- pdf2docx >= 0.5.8

## 🛠️ Development

```bash
# Clone repository
git clone https://github.com/LunaLynx12/Pdf2Docx.git
cd Pdf2Docx

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Install package in development mode
pip install -e .

# Install development dependencies
pip install -e ".[dev]"

# Run tests (if available)
pytest

# Format code
black pdf_to_docx/

# Lint code
flake8 pdf_to_docx/
```

## 📝 Examples

### Example 1: Simple Conversion

```python
from pdf_to_docx import PDFToDOCXConverter

converter = PDFToDOCXConverter("document.pdf")
result = converter.convert()
```

### Example 2: Custom Configuration

```python
from pdf_to_docx import PDFToDOCXConverter, ConversionConfig

config = ConversionConfig(
    start_page=0,
    end_page=10,
    overwrite=True,
    create_backup=True
)

converter = PDFToDOCXConverter("document.pdf", "output.docx", config=config)
result = converter.convert()
```

### Example 3: Batch Processing

```python
from pathlib import Path
from pdf_to_docx import PDFToDOCXConverter

for pdf_file in Path(".").glob("*.pdf"):
    converter = PDFToDOCXConverter(pdf_file, pdf_file.with_suffix('.docx'))
    converter.convert()
```

More examples can be found in the `examples/` directory.

## ⚠️ Limitations

- **Text-based PDFs** work best. Scanned PDFs may require OCR preprocessing.
- **Rule-based conversion** - The underlying pdf2docx library uses rule-based methods, so it may not achieve 100% perfect conversion for all PDF layouts.
- **Language support** - Best results with left-to-right languages and normal reading direction.
- **Complex layouts** - Very complex layouts may require manual adjustment in the output DOCX.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📚 About pdf2docx

This package uses the [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) library by Artifex Software, which is specifically designed for high-quality PDF to DOCX conversion. It:

- Extracts data from PDF with PyMuPDF (text, images, drawings)
- Parses layout with rules (sections, paragraphs, images, tables)
- Generates DOCX with python-docx

For more information, visit: https://pdf2docx.readthedocs.io/

## 📄 License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) - The underlying conversion library
- [PyMuPDF](https://github.com/pymupdf/PyMuPDF) - PDF processing
- [python-docx](https://github.com/python-openxml/python-docx) - DOCX generation

## 📧 Contact

For issues, questions, or contributions, please open an issue on GitHub.

---

**Note**: This package is a wrapper around pdf2docx that provides additional features like configuration management, logging, progress tracking, and a comprehensive CLI interface.

