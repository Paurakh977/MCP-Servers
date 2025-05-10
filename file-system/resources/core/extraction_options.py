"""Module for handling file extraction options."""

from typing import Dict, Any
from .dependencies import (
    has_tabula, has_pdfplumber, has_pymupdf,
    has_pil, has_epub_support
)

def get_file_extraction_options(file_extension: str) -> Dict[str, Any]:
    """
    Get available extraction options for a specific file type
    
    Args:
        file_extension: File extension with leading dot (.pdf, .docx, etc.)
    
    Returns:
        Dictionary of available extraction options
    """
    # Common options for all file types
    common_options = {
        "summarize": {
            "type": "boolean", 
            "description": "Summarize large content", 
            "default": False
        },
        "max_length": {
            "type": "integer", 
            "description": "Maximum length for summarization", 
            "default": 500
        }
    }
    
    # PDF-specific options
    if file_extension == '.pdf':
        options = {
            **common_options,
            "tables": {
                "type": "boolean", 
                "description": "Extract tables from PDF", 
                "default": True, 
                "available": has_tabula or has_pdfplumber
            },
            "images": {
                "type": "boolean", 
                "description": "Extract image information", 
                "default": True, 
                "available": has_pymupdf
            },
            "metadata_only": {
                "type": "boolean", 
                "description": "Extract only document metadata", 
                "default": False
            },
        }
        return options
    
    # Office document options
    elif file_extension in ['.docx', '.xlsx', '.pptx']:
        options = {
            **common_options,
            "extract_tables": {
                "type": "boolean", 
                "description": "Extract tables from document", 
                "default": True
            },
            "extract_images": {
                "type": "boolean", 
                "description": "Extract image information", 
                "default": True, 
                "available": has_pil
            },
            "metadata_only": {
                "type": "boolean", 
                "description": "Extract only document metadata", 
                "default": False
            },
        }
        return options
    
    # EPUB options
    elif file_extension == '.epub':
        options = {
            **common_options,
            "metadata_only": {
                "type": "boolean", 
                "description": "Extract only document metadata", 
                "default": False, 
                "available": has_epub_support
            },
            "extract_toc": {
                "type": "boolean", 
                "description": "Extract table of contents", 
                "default": True, 
                "available": has_epub_support
            },
        }
        return options
    
    # CSV options
    elif file_extension == '.csv':
        options = {
            **common_options,
            "analyze_columns": {
                "type": "boolean", 
                "description": "Analyze column data types", 
                "default": True
            },
            "max_rows": {
                "type": "integer", 
                "description": "Maximum number of rows to extract", 
                "default": 2000
            },
        }
        return options
    
    # Default to common options for other file types
    return common_options 