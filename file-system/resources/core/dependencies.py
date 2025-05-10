"""Module for handling dependency imports and checks."""

import os
import sys
from typing import Dict, Any, List

# PDF related imports
try:
    import PyPDF2
except ImportError:
    print("PyPDF2 not installed. To read PDF files: pip install PyPDF2")

try:
    import fitz  # PyMuPDF
    has_pymupdf = True
except ImportError:
    has_pymupdf = False
    print("PyMuPDF not installed. For better PDF image detection: pip install pymupdf")

try:
    import pdfplumber
    has_pdfplumber = True
except ImportError:
    has_pdfplumber = False
    print("pdfplumber not installed. For better PDF text extraction: pip install pdfplumber")

try:
    import tabula
    has_tabula = True
except ImportError:
    has_tabula = False
    print("Tabula-py not installed. For better PDF table extraction: pip install tabula-py")

# Office document imports
try:
    import docx  # for .docx files
except ImportError:
    print("python-docx not installed. To read Word files: pip install python-docx")

try:
    import openpyxl  # for .xlsx files
except ImportError:
    print("openpyxl not installed. To read Excel files: pip install openpyxl")

try:
    from pptx import Presentation  # for .pptx files
except ImportError:
    print("python-pptx not installed. To read PowerPoint files: pip install python-pptx")

# Image handling
try:
    from PIL import Image
    has_pil = True
except ImportError:
    print("Pillow not installed. For better image analysis: pip install Pillow")
    has_pil = False

# Ebook and RTF support
try:
    import ebooklib
    from ebooklib import epub
    from bs4 import BeautifulSoup
    has_epub_support = True
except ImportError:
    has_epub_support = False
    print("EPUB support not available. To read EPUB files: pip install ebooklib beautifulsoup4")

try:
    import striprtf.striprtf as striprtf
    has_rtf_support = True
except ImportError:
    has_rtf_support = False
    print("RTF support not available. To read RTF files: pip install striprtf")

# Export all feature flags and modules
__all__ = [
    # Feature flags
    'has_pymupdf',
    'has_pdfplumber',
    'has_tabula',
    'has_pil',
    'has_epub_support',
    'has_rtf_support',
    
    # Modules
    'PyPDF2',
    'fitz',
    'pdfplumber',
    'tabula',
    'docx',
    'openpyxl',
    'Presentation',
    'Image',
    'epub',
    'BeautifulSoup',
    'striprtf'
] 