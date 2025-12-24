"""
Configuration module for PDF to DOCX conversion.
"""

from dataclasses import dataclass, field
from typing import Optional, List
from pathlib import Path


@dataclass
class ConversionConfig:
    """Configuration for PDF to DOCX conversion."""
    
    # Page range
    start_page: int = 0
    end_page: Optional[int] = None
    
    # Conversion options
    multi_processing: bool = False
    cpu_count: Optional[int] = None
    
    # Layout preservation
    preserve_layout: bool = True
    preserve_images: bool = True
    preserve_tables: bool = True
    
    # Output options
    overwrite: bool = False
    create_backup: bool = False
    
    # Advanced options
    table_settings: dict = field(default_factory=lambda: {
        "min_border_vertical": 0.5,
        "min_border_horizontal": 0.5,
        "intersection_threshold": 0.25,
        "min_words_vertical": 3,
        "min_words_horizontal": 1,
    })
    
    # Logging
    verbose: bool = True
    log_file: Optional[Path] = None
    
    def validate(self) -> List[str]:
        """
        Validate configuration and return list of errors.
        
        Returns:
            List of error messages (empty if valid)
        """
        errors = []
        
        if self.start_page < 0:
            errors.append("start_page must be non-negative")
        
        if self.end_page is not None and self.end_page < self.start_page:
            errors.append("end_page must be >= start_page")
        
        if self.cpu_count is not None and self.cpu_count < 1:
            errors.append("cpu_count must be >= 1")
        
        return errors
    
    def to_dict(self) -> dict:
        """Convert configuration to dictionary."""
        return {
            "start_page": self.start_page,
            "end_page": self.end_page,
            "multi_processing": self.multi_processing,
            "cpu_count": self.cpu_count,
            "preserve_layout": self.preserve_layout,
            "preserve_images": self.preserve_images,
            "preserve_tables": self.preserve_tables,
            "overwrite": self.overwrite,
            "create_backup": self.create_backup,
            "table_settings": self.table_settings,
            "verbose": self.verbose,
            "log_file": str(self.log_file) if self.log_file else None,
        }


# Default configuration
DEFAULT_CONFIG = ConversionConfig()

