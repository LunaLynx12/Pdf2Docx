"""
Advanced usage examples for PDF to DOCX converter.
"""

from pathlib import Path
from pdf_to_docx import PDFToDOCXConverter, ConversionConfig
from pdf_to_docx.utils import validate_pdf, get_file_info, format_file_size


def batch_conversion():
    """Convert multiple PDF files."""
    pdf_files = [
        "document1.pdf",
        "document2.pdf",
        "document3.pdf",
    ]
    
    results = []
    for pdf_file in pdf_files:
        pdf_path = Path(pdf_file)
        if not pdf_path.exists():
            print(f"Skipping {pdf_file}: file not found")
            continue
        
        # Validate PDF
        is_valid, error = validate_pdf(pdf_path)
        if not is_valid:
            print(f"Skipping {pdf_file}: {error}")
            continue
        
        # Convert
        try:
            converter = PDFToDOCXConverter(
                pdf_path=pdf_path,
                docx_path=pdf_path.with_suffix('.docx')
            )
            result = converter.convert()
            results.append(result)
            print(f"✓ Converted: {pdf_file} -> {result}")
        except Exception as e:
            print(f"✗ Failed to convert {pdf_file}: {e}")
    
    print(f"\nBatch conversion complete: {len(results)}/{len(pdf_files)} files converted")


def conversion_with_validation():
    """Convert with detailed validation and information."""
    pdf_path = Path("input.pdf")
    
    # Get file information
    info = get_file_info(pdf_path)
    print(f"PDF Information:")
    print(f"  Name: {info['name']}")
    print(f"  Size: {format_file_size(info['size'])}")
    print(f"  Pages: {info['pages']}")
    print(f"  Encrypted: {info['encrypted']}")
    
    if info['encrypted']:
        print("  ⚠ Warning: PDF is encrypted")
        return
    
    # Validate PDF
    is_valid, error = validate_pdf(pdf_path)
    if not is_valid:
        print(f"Error: {error}")
        return
    
    # Convert with custom config
    config = ConversionConfig(
        start_page=0,
        end_page=None,  # All pages
        overwrite=True,
        create_backup=True,
        verbose=True
    )
    
    converter = PDFToDOCXConverter(
        pdf_path=pdf_path,
        docx_path=pdf_path.with_suffix('.docx'),
        config=config
    )
    
    try:
        result = converter.convert()
        output_size = result.stat().st_size
        print(f"\n✓ Conversion successful!")
        print(f"  Output: {result}")
        print(f"  Size: {format_file_size(output_size)}")
    except Exception as e:
        print(f"\n✗ Conversion failed: {e}")


def convert_specific_pages():
    """Convert specific page range."""
    pdf_path = Path("input.pdf")
    
    # Convert pages 5-10 (0-based indexing)
    config = ConversionConfig(
        start_page=5,
        end_page=10,
        overwrite=True
    )
    
    converter = PDFToDOCXConverter(
        pdf_path=pdf_path,
        docx_path="output_pages_5-10.docx",
        config=config
    )
    
    result = converter.convert()
    print(f"Converted pages 5-10 to: {result}")


if __name__ == "__main__":
    # Uncomment the example you want to run
    # batch_conversion()
    # conversion_with_validation()
    # convert_specific_pages()
    pass

