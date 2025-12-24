"""
PDF to DOCX Converter Package
A professional Python package for converting PDF files to DOCX format
with excellent format preservation.
"""

from .converter import PDFToDOCXConverter
from .config import ConversionConfig, DEFAULT_CONFIG
from .utils import validate_pdf, validate_output_path, get_file_info

__version__ = "1.0.0"
__author__ = "Your Name"
__all__ = [
    "PDFToDOCXConverter",
    "ConversionConfig",
    "DEFAULT_CONFIG",
    "validate_pdf",
    "validate_output_path",
    "get_file_info",
]

