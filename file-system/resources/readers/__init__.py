"""File readers for different file formats."""

# Import dependencies check variables
from .text_reader import read_text_file
from .pdf_reader import read_pdf_file, has_pymupdf, has_tabula, has_pdfplumber
from .office_readers import read_docx_file, read_xlsx_file, read_pptx_file, has_pil
from .data_readers import read_csv_file
from .ebook_readers import read_epub_file, has_epub_support
from .format_readers import read_rtf_file, has_rtf_support

__all__ = [
    'read_text_file',
    'read_pdf_file',
    'read_docx_file',
    'read_xlsx_file',
    'read_pptx_file',
    'read_csv_file',
    'read_epub_file',
    'read_rtf_file',
    'has_pymupdf',
    'has_tabula',
    'has_pdfplumber',
    'has_pil',
    'has_epub_support',
    'has_rtf_support',
] 