"""
Setup script for PDF to DOCX converter package.
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read README for long description
readme_file = Path(__file__).parent / "README.md"
long_description = readme_file.read_text(encoding="utf-8") if readme_file.exists() else ""

# Read version from package
version_file = Path(__file__).parent / "pdf_to_docx" / "__init__.py"
version = "1.0.0"
if version_file.exists():
    for line in version_file.read_text().splitlines():
        if line.startswith("__version__"):
            version = line.split("=")[1].strip().strip('"').strip("'")
            break

setup(
    name="pdf-to-docx",
    version=version,
    author="Your Name",
    author_email="your.email@example.com",
    description="A professional Python package for converting PDF files to DOCX format with excellent format preservation",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/LunaLynx12/Pdf2Docx",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Intended Audience :: End Users/Desktop",
        "Topic :: Office/Business",
        "Topic :: Text Processing",
        "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    install_requires=[
        "pdf2docx>=0.5.8",
    ],
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
            "mypy>=1.0.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "pdf-to-docx=pdf_to_docx.cli:main",
        ],
    },
    include_package_data=True,
    zip_safe=False,
)

