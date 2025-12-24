"""
Main converter module for PDF to DOCX conversion.
"""

import logging
from pathlib import Path
from typing import Optional, Callable
from pdf2docx import Converter as PDF2DOCXConverter

from .config import ConversionConfig, DEFAULT_CONFIG
from .utils import (
    validate_pdf,
    validate_output_path,
    get_file_info,
    format_file_size,
    create_backup,
    setup_logging,
)


class PDFToDOCXConverter:
    """
    Converts PDF files to DOCX format with formatting preservation.
    
    Uses the pdf2docx library for high-quality conversion with excellent
    format preservation including tables, images, and text formatting.
    """
    
    def __init__(
        self,
        pdf_path: Path,
        docx_path: Optional[Path] = None,
        config: Optional[ConversionConfig] = None
    ):
        """
        Initialize the converter.
        
        Args:
            pdf_path: Path to the input PDF file
            docx_path: Path to the output DOCX file (optional, auto-generated if not provided)
            config: Conversion configuration (optional, uses default if not provided)
        """
        self.pdf_path = Path(pdf_path)
        self.config = config or DEFAULT_CONFIG
        
        # Validate PDF
        is_valid, error = validate_pdf(self.pdf_path)
        if not is_valid:
            raise ValueError(f"Invalid PDF file: {error}")
        
        # Set output path
        if docx_path:
            self.docx_path = Path(docx_path)
        else:
            self.docx_path = self.pdf_path.with_suffix('.docx')
        
        # Validate output path
        is_valid, error = validate_output_path(self.docx_path, self.config.overwrite)
        if not is_valid:
            raise ValueError(f"Invalid output path: {error}")
        
        # Setup logging
        self.logger = setup_logging(
            verbose=self.config.verbose,
            log_file=self.config.log_file
        )
        
        # Converter instance (created during convert)
        self.converter: Optional[PDF2DOCXConverter] = None
        
        # Progress callback
        self.progress_callback: Optional[Callable[[int, int], None]] = None
    
    def set_progress_callback(self, callback: Callable[[int, int], None]):
        """
        Set a progress callback function.
        
        Args:
            callback: Function that takes (current_page, total_pages) as arguments
        """
        self.progress_callback = callback
    
    def _log_file_info(self):
        """Log information about the PDF file."""
        info = get_file_info(self.pdf_path)
        self.logger.info(f"PDF File: {info['name']}")
        self.logger.info(f"Size: {format_file_size(info['size'])}")
        self.logger.info(f"Pages: {info['pages']}")
        self.logger.info(f"Encrypted: {info['encrypted']}")
        
        if info['encrypted']:
            self.logger.warning("PDF is encrypted. Conversion may fail if password is required.")
    
    def _create_backup_if_needed(self):
        """Create backup of output file if it exists and backup is enabled."""
        if self.docx_path.exists() and self.config.create_backup:
            backup_path = create_backup(self.docx_path)
            if backup_path:
                self.logger.info(f"Created backup: {backup_path}")
            else:
                self.logger.warning("Failed to create backup")
    
    def convert(self) -> Path:
        """
        Convert PDF to DOCX with format preservation.
        
        Returns:
            Path to the created DOCX file
            
        Raises:
            ValueError: If configuration is invalid
            RuntimeError: If conversion fails
        """
        # Validate configuration
        config_errors = self.config.validate()
        if config_errors:
            raise ValueError(f"Invalid configuration: {', '.join(config_errors)}")
        
        # Log file information
        self._log_file_info()
        
        # Create backup if needed
        self._create_backup_if_needed()
        
        self.logger.info(f"Starting conversion to: {self.docx_path}")
        self.logger.info("This may take a moment depending on PDF complexity...")
        
        try:
            # Initialize converter
            self.converter = PDF2DOCXConverter(str(self.pdf_path))
            
            # Get page range
            start = self.config.start_page
            end = self.config.end_page
            
            # Convert PDF to DOCX
            # pdf2docx handles all the complex layout preservation
            self.converter.convert(
                str(self.docx_path),
                start=start,
                end=end
            )
            
            # Verify output file was created
            if not self.docx_path.exists():
                raise RuntimeError("Conversion completed but output file was not created")
            
            output_size = self.docx_path.stat().st_size
            self.logger.info(f"Conversion complete!")
            self.logger.info(f"Output file: {self.docx_path}")
            self.logger.info(f"Output size: {format_file_size(output_size)}")
            
            return self.docx_path
            
        except Exception as e:
            self.logger.error(f"Conversion failed: {e}", exc_info=True)
            raise RuntimeError(f"Failed to convert PDF: {str(e)}") from e
        
        finally:
            if self.converter:
                self.converter.close()
                self.converter = None
    
    def convert_with_progress(self) -> Path:
        """
        Convert PDF to DOCX with progress tracking.
        
        Returns:
            Path to the created DOCX file
        """
        # Get total pages for progress tracking
        info = get_file_info(self.pdf_path)
        total_pages = info['pages']
        
        # Adjust for page range
        start = self.config.start_page
        end = self.config.end_page or total_pages
        pages_to_convert = end - start + 1
        
        if self.progress_callback:
            self.progress_callback(0, pages_to_convert)
        
        try:
            result = self.convert()
            
            if self.progress_callback:
                self.progress_callback(pages_to_convert, pages_to_convert)
            
            return result
        except Exception as e:
            if self.progress_callback:
                self.progress_callback(0, pages_to_convert)  # Reset on error
            raise
    
    def __enter__(self):
        """Context manager entry."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        if self.converter:
            self.converter.close()
            self.converter = None
        return False

