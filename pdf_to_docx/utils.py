"""
Utility functions for PDF to DOCX conversion.
"""

import logging
from pathlib import Path
from typing import Optional, Tuple
import fitz  # PyMuPDF


def validate_pdf(pdf_path: Path) -> Tuple[bool, Optional[str]]:
    """
    Validate that a file is a valid PDF.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        Tuple of (is_valid, error_message)
    """
    if not pdf_path.exists():
        return False, f"File not found: {pdf_path}"
    
    if not pdf_path.is_file():
        return False, f"Path is not a file: {pdf_path}"
    
    if pdf_path.suffix.lower() != ".pdf":
        return False, f"File is not a PDF: {pdf_path}"
    
    # Try to open with PyMuPDF to verify it's a valid PDF
    try:
        doc = fitz.open(str(pdf_path))
        page_count = len(doc)
        doc.close()
        
        if page_count == 0:
            return False, "PDF has no pages"
        
        return True, None
    except Exception as e:
        return False, f"Invalid PDF file: {str(e)}"


def validate_output_path(docx_path: Path, overwrite: bool = False) -> Tuple[bool, Optional[str]]:
    """
    Validate output path for DOCX file.
    
    Args:
        docx_path: Path to the output DOCX file
        overwrite: Whether to allow overwriting existing files
        
    Returns:
        Tuple of (is_valid, error_message)
    """
    if docx_path.suffix.lower() != ".docx":
        return False, f"Output file must have .docx extension: {docx_path}"
    
    if docx_path.exists() and not overwrite:
        return False, f"Output file already exists: {docx_path}. Use overwrite=True to replace."
    
    # Check if parent directory exists and is writable
    parent = docx_path.parent
    if not parent.exists():
        try:
            parent.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            return False, f"Cannot create output directory: {str(e)}"
    
    if not parent.is_dir():
        return False, f"Output path parent is not a directory: {parent}"
    
    return True, None


def get_file_info(pdf_path: Path) -> dict:
    """
    Get information about a PDF file.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        Dictionary with file information
    """
    info = {
        "path": str(pdf_path),
        "name": pdf_path.name,
        "size": pdf_path.stat().st_size if pdf_path.exists() else 0,
        "pages": 0,
        "encrypted": False,
        "metadata": {},
    }
    
    try:
        doc = fitz.open(str(pdf_path))
        info["pages"] = len(doc)
        info["encrypted"] = doc.is_encrypted
        info["metadata"] = doc.metadata
        doc.close()
    except Exception as e:
        logging.warning(f"Could not read PDF info: {e}")
    
    return info


def format_file_size(size_bytes: int) -> str:
    """
    Format file size in human-readable format.
    
    Args:
        size_bytes: Size in bytes
        
    Returns:
        Formatted string (e.g., "1.5 MB")
    """
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} PB"


def create_backup(file_path: Path) -> Optional[Path]:
    """
    Create a backup of a file.
    
    Args:
        file_path: Path to the file to backup
        
    Returns:
        Path to the backup file, or None if backup failed
    """
    if not file_path.exists():
        return None
    
    backup_path = file_path.with_suffix(f"{file_path.suffix}.bak")
    counter = 1
    
    # Find unique backup name
    while backup_path.exists():
        backup_path = file_path.with_suffix(f"{file_path.suffix}.bak.{counter}")
        counter += 1
    
    try:
        import shutil
        shutil.copy2(file_path, backup_path)
        return backup_path
    except Exception as e:
        logging.error(f"Failed to create backup: {e}")
        return None


def setup_logging(verbose: bool = True, log_file: Optional[Path] = None) -> logging.Logger:
    """
    Setup logging configuration.
    
    Args:
        verbose: Whether to enable verbose logging
        log_file: Optional path to log file
        
    Returns:
        Configured logger
    """
    logger = logging.getLogger("pdf_to_docx")
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    
    # Remove existing handlers
    logger.handlers.clear()
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO if verbose else logging.WARNING)
    console_format = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    console_handler.setFormatter(console_format)
    logger.addHandler(console_handler)
    
    # File handler (if specified)
    if log_file:
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_format = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        file_handler.setFormatter(file_format)
        logger.addHandler(file_handler)
    
    return logger

