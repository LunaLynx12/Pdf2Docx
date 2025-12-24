"""
Command-line interface for PDF to DOCX converter.
"""

import sys
import argparse
from pathlib import Path
from typing import Optional

from .converter import PDFToDOCXConverter
from .config import ConversionConfig
from .utils import get_file_info, format_file_size


def create_parser() -> argparse.ArgumentParser:
    """Create command-line argument parser."""
    parser = argparse.ArgumentParser(
        description="Convert PDF files to DOCX format with excellent format preservation.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s document.pdf
  %(prog)s input.pdf output.docx
  %(prog)s input.pdf --start-page 0 --end-page 10
  %(prog)s input.pdf --overwrite --verbose
        """
    )
    
    parser.add_argument(
        "pdf_path",
        type=str,
        help="Path to the input PDF file"
    )
    
    parser.add_argument(
        "output",
        type=str,
        nargs="?",
        default=None,
        help="Path to the output DOCX file (optional, defaults to input filename with .docx extension)"
    )
    
    # Page range options
    parser.add_argument(
        "--start-page",
        type=int,
        default=0,
        help="Start page index (0-based, default: 0)"
    )
    
    parser.add_argument(
        "--end-page",
        type=int,
        default=None,
        help="End page index (0-based, default: last page)"
    )
    
    # Output options
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite output file if it exists"
    )
    
    parser.add_argument(
        "--backup",
        action="store_true",
        help="Create backup of existing output file"
    )
    
    # Logging options
    parser.add_argument(
        "--verbose",
        action="store_true",
        default=True,
        help="Enable verbose output (default: True)"
    )
    
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Disable verbose output"
    )
    
    parser.add_argument(
        "--log-file",
        type=str,
        default=None,
        help="Path to log file (optional)"
    )
    
    # Advanced options
    parser.add_argument(
        "--multi-processing",
        action="store_true",
        help="Enable multi-processing (experimental)"
    )
    
    parser.add_argument(
        "--cpu-count",
        type=int,
        default=None,
        help="Number of CPU cores to use for multi-processing"
    )
    
    return parser


def print_file_info(pdf_path: Path):
    """Print information about the PDF file."""
    info = get_file_info(pdf_path)
    print(f"\nPDF Information:")
    print(f"  File: {info['name']}")
    print(f"  Size: {format_file_size(info['size'])}")
    print(f"  Pages: {info['pages']}")
    print(f"  Encrypted: {info['encrypted']}")
    if info['encrypted']:
        print(f"  ⚠ Warning: PDF is encrypted")
    print()


def main(args: Optional[list] = None) -> int:
    """
    Main entry point for command-line interface.
    
    Args:
        args: Command-line arguments (defaults to sys.argv)
        
    Returns:
        Exit code (0 for success, non-zero for error)
    """
    parser = create_parser()
    parsed_args = parser.parse_args(args)
    
    # Handle quiet flag
    verbose = parsed_args.verbose and not parsed_args.quiet
    
    try:
        # Validate input file
        pdf_path = Path(parsed_args.pdf_path)
        if not pdf_path.exists():
            print(f"Error: PDF file not found: {pdf_path}", file=sys.stderr)
            return 1
        
        # Set output path
        if parsed_args.output:
            docx_path = Path(parsed_args.output)
        else:
            docx_path = pdf_path.with_suffix('.docx')
        
        # Print file info
        if verbose:
            print_file_info(pdf_path)
        
        # Create configuration
        config = ConversionConfig(
            start_page=parsed_args.start_page,
            end_page=parsed_args.end_page,
            overwrite=parsed_args.overwrite,
            create_backup=parsed_args.backup,
            verbose=verbose,
            log_file=Path(parsed_args.log_file) if parsed_args.log_file else None,
            multi_processing=parsed_args.multi_processing,
            cpu_count=parsed_args.cpu_count,
        )
        
        # Create converter and convert
        converter = PDFToDOCXConverter(
            pdf_path=pdf_path,
            docx_path=docx_path,
            config=config
        )
        
        result_path = converter.convert()
        
        if verbose:
            print(f"\n✓ Conversion successful!")
            print(f"  Output: {result_path}")
            output_size = result_path.stat().st_size
            print(f"  Size: {format_file_size(output_size)}")
        
        return 0
        
    except KeyboardInterrupt:
        print("\n\nConversion cancelled by user.", file=sys.stderr)
        return 130
    
    except Exception as e:
        print(f"\nError: {e}", file=sys.stderr)
        if verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())

