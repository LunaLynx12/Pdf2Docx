"""
Basic usage example for PDF to DOCX converter.
"""

from pathlib import Path
from pdf_to_docx import PDFToDOCXConverter, ConversionConfig

# Example 1: Simple conversion
def simple_conversion():
    """Simple conversion with default settings."""
    converter = PDFToDOCXConverter(
        pdf_path="input.pdf",
        docx_path="output.docx"
    )
    result = converter.convert()
    print(f"Converted to: {result}")


# Example 2: Conversion with custom configuration
def custom_conversion():
    """Conversion with custom configuration."""
    config = ConversionConfig(
        start_page=0,
        end_page=10,  # Convert first 11 pages
        overwrite=True,
        create_backup=True,
        verbose=True
    )
    
    converter = PDFToDOCXConverter(
        pdf_path="input.pdf",
        docx_path="output.docx",
        config=config
    )
    result = converter.convert()
    print(f"Converted to: {result}")


# Example 3: Using context manager
def context_manager_example():
    """Using converter as a context manager."""
    with PDFToDOCXConverter("input.pdf", "output.docx") as converter:
        result = converter.convert()
        print(f"Converted to: {result}")


# Example 4: Progress tracking
def progress_tracking_example():
    """Conversion with progress callback."""
    def progress_callback(current, total):
        percentage = (current / total) * 100
        print(f"Progress: {current}/{total} pages ({percentage:.1f}%)")
    
    converter = PDFToDOCXConverter("input.pdf", "output.docx")
    converter.set_progress_callback(progress_callback)
    result = converter.convert_with_progress()
    print(f"Converted to: {result}")


if __name__ == "__main__":
    # Uncomment the example you want to run
    # simple_conversion()
    # custom_conversion()
    # context_manager_example()
    # progress_tracking_example()
    pass

