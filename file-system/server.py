from mcp.server import Server
import mcp.types as types
import asyncio
import os, sys
from typing import Optional, Dict, Any, List
from mcp.server.stdio import stdio_server
from pydantic import AnyUrl
import urllib.parse
import mimetypes
import json
import time

# Direct imports for required libraries
try:
    import PyPDF2
except ImportError:
    print("PyPDF2 not installed. To read PDF files: pip install PyPDF2")

try:
    import docx  # for .docx files
except ImportError:
    print("python-docx not installed. To read Word files: pip install python-docx")

try:
    import fitz  # PyMuPDF
    has_pymupdf = True
except ImportError:
    has_pymupdf = False
    print("PyMuPDF not installed. For better PDF image detection: pip install pymupdf")

try:
    import openpyxl  # for .xlsx files
except ImportError:
    print("openpyxl not installed. To read Excel files: pip install openpyxl")

try:
    from pptx import Presentation  # for .pptx files
except ImportError:
    print("python-pptx not installed. To read PowerPoint files: pip install python-pptx")

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

try:
    from PIL import Image
    has_pil = True
except ImportError:
    print("Pillow not installed. For better image analysis: pip install Pillow")
    has_pil = False

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

# Import our file extraction utilities
from resources import extract_file_content, read_file
from resources.readers import (
    read_text_file, read_pdf_file, read_docx_file, read_xlsx_file,
    read_pptx_file, read_csv_file, read_epub_file, read_rtf_file,
    has_tabula, has_pdfplumber, has_pil,
    has_epub_support, has_rtf_support
)
from resources.utils.formatters import summarize_content

# Override the problematic reader functions with direct implementations
def patched_read_pdf_file(file_path: str) -> str:
    """Extract text and identify images from PDF files with layout preservation."""
    content = []
    
    # Try pdfplumber first if available
    if has_pdfplumber:
        try:
            content.append("--- Document Metadata ---")
            pdf = pdfplumber.open(file_path)
            
            # Basic document info
            content.append(f"PDF Document: {os.path.basename(file_path)}")
            content.append(f"Number of pages: {len(pdf.pages)}")
            
            # Try to extract metadata
            if hasattr(pdf, 'metadata') and pdf.metadata:
                for key, value in pdf.metadata.items():
                    if value and str(value).strip():
                        # Clean up key name
                        clean_key = key[1:] if isinstance(key, str) and key.startswith('/') else key
                        content.append(f"{clean_key}: {value}")
            
            content.append("-" * 40)
            
            # Process each page
            for i, page in enumerate(pdf.pages):
                content.append(f"--- Page {i + 1} ---")
                
                # Extract page dimensions
                width, height = page.width, page.height
                content.append(f"Page dimensions: {width:.2f} x {height:.2f} points")
                
                # Extract text with layout preservation
                page_text = page.extract_text(x_tolerance=3, y_tolerance=3)
                
                # Try to detect tables
                tables = page.extract_tables()
                if tables:
                    content.append(f"[Contains {len(tables)} table{'s' if len(tables) > 1 else ''}]")
                    
                    # Format each table
                    for t_idx, table in enumerate(tables):
                        content.append(f"\nTable {t_idx + 1}:")
                        
                        # Format the table with proper alignment
                        for row in table:
                            # Clean row data and handle None values
                            cleaned_row = [str(cell).strip() if cell is not None else "" for cell in row]
                            row_text = " | ".join(cleaned_row)
                            content.append(row_text)
                        
                        content.append("")  # Add spacing after table
                
                # Add the page text with layout preservation
                if not page_text or page_text.isspace():
                    content.append("[This page appears to be empty or contains only non-text elements]")
                else:
                    content.append(page_text)
            
            pdf.close()
            return "\n\n".join(content)
        except Exception as e:
            print(f"pdfplumber processing failed: {str(e)}. Trying PyMuPDF.")
    
    # Try PyMuPDF if available
    if has_pymupdf:
        try:
            pdf_document = fitz.open(file_path)
            
            # Extract document metadata
            content.append("--- Document Metadata ---")
            metadata = pdf_document.metadata
            for key, value in metadata.items():
                if value:
                    content.append(f"{key}: {value}")
            content.append("-" * 40)
            
            # Document summary
            content.append(f"PDF Document: {os.path.basename(file_path)}")
            content.append(f"Number of pages: {len(pdf_document)}")
            content.append(f"PDF Version: {pdf_document.pdf_version}")
            if pdf_document.is_encrypted:
                content.append("Status: Encrypted")
            
            content.append("-" * 40)
            
            # Process each page
            for page_num, page in enumerate(pdf_document):
                content.append(f"--- Page {page_num + 1} ---")
                
                # Get page text
                page_text = page.get_text()
                
                # Add the page text
                if not page_text.strip():
                    content.append("[This page appears to be empty or contains only non-text elements]")
                else:
                    content.append(page_text.strip())
            
            return "\n\n".join(content)
        except Exception as e:
            print(f"PyMuPDF processing failed: {str(e)}. Falling back to PyPDF2.")
    
    # Use PyPDF2 as final fallback
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            num_pages = len(pdf_reader.pages)
            
            # Try to get document info
            content.append("--- Document Metadata ---")
            info = pdf_reader.metadata
            if info:
                for key in info:
                    if info.get(key):
                        # Clean up key name by removing leading slash
                        clean_key = key[1:] if key.startswith('/') else key
                        content.append(f"{clean_key}: {info.get(key)}")
            
            # Add document summary
            content.append(f"PDF Document: {os.path.basename(file_path)}")
            content.append(f"Number of pages: {num_pages}")
            content.append("-" * 40)
            
            # Extract text from each page
            for page_num in range(num_pages):
                page = pdf_reader.pages[page_num]
                
                # Start page content
                content.append(f"--- Page {page_num + 1} ---")
                
                # Extract text with better layout handling
                try:
                    page_text = page.extract_text()
                except:
                    page_text = "Error extracting text from this page"
                
                # Handle empty pages
                if not page_text or page_text.isspace():
                    content.append("[This page appears to be empty or contains only non-text elements]")
                else:
                    content.append(page_text)
        
        return "\n\n".join(content)
    except Exception as e:
        return f"Error reading PDF file: {str(e)}"

def patched_read_docx_file(file_path: str) -> str:
    """Extract text and identify objects from Microsoft Word (.docx) files in sequential order."""
    try:
        doc = docx.Document(file_path)
        full_text = []
        
        # Document properties first
        try:
            full_text.append("--- Document Properties ---")
            core_props = doc.core_properties
            if hasattr(core_props, 'title') and core_props.title:
                full_text.append(f"Title: {core_props.title}")
            if hasattr(core_props, 'author') and core_props.author:
                full_text.append(f"Author: {core_props.author}")
            if hasattr(core_props, 'created') and core_props.created:
                full_text.append(f"Created: {core_props.created}")
            if hasattr(core_props, 'modified') and core_props.modified:
                full_text.append(f"Modified: {core_props.modified}")
            full_text.append("-" * 40)
        except:
            pass  # Ignore if properties can't be accessed
        
        # Document statistics
        full_text.append("--- Document Statistics ---")
        full_text.append(f"Paragraphs: {len(doc.paragraphs)}")
        full_text.append(f"Sections: {len(doc.sections)}")
        full_text.append(f"Tables: {len(doc.tables)}")
        full_text.append("-" * 40)
        
        # Main content extraction
        full_text.append("--- Document Content ---")
        
        # Get paragraphs
        for para in doc.paragraphs:
            if para.text.strip():  # Skip empty paragraphs
                full_text.append(para.text)
        
        # Process tables
        for i, table in enumerate(doc.tables):
            full_text.append(f"\n--- Table {i+1} ---")
            
            for row in table.rows:
                row_cells = []
                for cell in row.cells:
                    text = cell.text.strip()
                    row_cells.append(text if text else "")
                full_text.append(" | ".join(row_cells))
            
            full_text.append("--- End Table ---\n")
        
        return "\n".join(full_text)
    except Exception as e:
        return f"Error reading DOCX file: {str(e)}"

# Replace imported functions with our patched versions
read_pdf_file = patched_read_pdf_file
read_docx_file = patched_read_docx_file

# Patched version of extract_file_content
def patched_extract_file_content(file_path: str) -> Dict[str, Any]:
    """Extract content from a file based on its extension."""
    if not os.path.exists(file_path):
        return {"success": False, "error": f"File not found: {file_path}", "content": ""}
    
    file_extension = os.path.splitext(file_path)[1].lower()
    content = ""
    
    try:
        # Handle different file types
        if file_extension in ['.txt', '.md', '.json', '.html', '.xml', '.log', '.py', '.js', '.css', '.java', '.ini', '.conf', '.cfg']:
            content = read_text_file(file_path)
        elif file_extension == '.pdf':
            content = patched_read_pdf_file(file_path)
        elif file_extension == '.docx':
            content = patched_read_docx_file(file_path)
        elif file_extension == '.xlsx':
            content = read_xlsx_file(file_path)
        elif file_extension == '.pptx':
            content = read_pptx_file(file_path)
        elif file_extension == '.csv':
            content = read_csv_file(file_path)
        elif file_extension == '.epub':
            content = read_epub_file(file_path)
        elif file_extension == '.rtf':
            content = read_rtf_file(file_path)
        else:
            # Try to read as text file first, then fall back to binary warning
            try:
                content = read_text_file(file_path)
            except Exception:
                return {
                    "success": False, 
                    "error": f"Unsupported file type: {file_extension}", 
                    "content": f"This file type ({file_extension}) is not directly supported."
                }
        
        return {"success": True, "content": content, "file_path": file_path, "file_type": file_extension}
    
    except Exception as e:
        return {"success": False, "error": str(e), "content": ""}

# Patched version of read_file
def patched_read_file(path: str, summarize: bool = False, max_summary_length: int = 500) -> Dict[str, Any]:
    """
    Reads and returns the contents of the file at 'path' with appropriate handling per file type.
    
    Args:
        path (str): Path to the file to read
        summarize (bool): Whether to summarize very large content
        max_summary_length (int): Maximum length for summary if summarizing
        
    Returns:
        Dict with keys:
        - success (bool): Whether the read was successful
        - content (str): The file content, possibly summarized
        - file_path (str): Original file path
        - file_type (str): File extension
        - error (str, optional): Error message if success is False
    """
    result = patched_extract_file_content(path)
    
    # Summarize very large content if requested
    if summarize and result["success"] and len(result["content"]) > max_summary_length:
        result["content"] = summarize_content(result["content"], max_summary_length)
        result["summarized"] = True
    
    return result

# Override the imported functions with our patched versions
extract_file_content = patched_extract_file_content
read_file = patched_read_file

# Initialize allowed directories list - will be populated with command-line arguments
allowed_directories = [os.getcwd()]  # Default to current directory

server = Server(
    name="file-system",
    instructions="You are a file system MCP server. You can read and extract content from various file types including text files, PDFs, Office documents (Word, Excel, PowerPoint), CSV, EPUB, and RTF. The server provides advanced extraction capabilities for tables, images, and metadata. PDF documents will include information about tables, images, and annotations when available. Office documents will extract embedded images, charts, tables, and preserve formatting where possible.",
    version="0.1",
)


def check_path_security(allowed_base_paths: List[str], target_path: str) -> dict:
    """
    Checks if a target_path is allowed based on any of the allowed_base_paths.
    
    Args:
        allowed_base_paths: List of allowed root directory paths.
        target_path: Path to a file or directory to check.
    
    Returns:
        Dictionary containing detailed results of the checks.
    """
    # Initialize results with default values for path that's not allowed
    results = {
        'original_target': target_path,
        'normalized_target': None,
        'normalization_successful': False,
        'target_exists': False,
        'same_drive_check': True,
        'is_within_base': False,
        'is_allowed': False,
        'message': '',
        'allowed_base': None,  # Track which allowed base succeeded
    }

    try:
        normalized_target = os.path.realpath(os.path.normpath(target_path))
        results['normalized_target'] = normalized_target
        results['normalization_successful'] = True
    except Exception as e:
        results['message'] = f"Normalization error for target: {e}"
        return results

    if not os.path.exists(normalized_target):
        results['message'] = f"Target does not exist: {normalized_target}"
        return results
    results['target_exists'] = True
    
    # Try each allowed base path until one succeeds
    for base_path in allowed_base_paths:
        base_check = check_against_single_base(base_path, normalized_target)
        
        # If this base path works, use its results
        if base_check['is_allowed']:
            # Copy all relevant results from the successful check
            results['is_allowed'] = True
            results['is_within_base'] = True
            results['same_drive_check'] = True
            results['message'] = f"Path is allowed via {base_path}"
            results['allowed_base'] = base_path
            results['relative_path_from_base'] = base_check['relative_path_from_base']
            return results
    
    # If we get here, no allowed path matched
    results['message'] = f"Access denied: Path not within any allowed directory"
    return results


def check_against_single_base(base_path: str, normalized_target: str) -> dict:
    """Checks a normalized target path against a single base directory."""
    result = {
        'original_base': base_path,
        'normalized_base': None,
        'base_exists': False,
        'same_drive_check': True,
        'relative_path_from_base': None,
        'is_within_base': False,
        'is_allowed': False,
        'message': ''
    }
    
    try:
        normalized_base = os.path.realpath(os.path.normpath(base_path))
        result['normalized_base'] = normalized_base
    except Exception as e:
        result['message'] = f"Base normalization error: {e}"
        return result
    
    if not os.path.exists(normalized_base) or not os.path.isdir(normalized_base):
        result['message'] = f"Base path does not exist or not a directory: {normalized_base}"
        return result
    result['base_exists'] = True
    
    # Windows drive check
    if sys.platform == 'win32':
        base_drive = os.path.splitdrive(normalized_base)[0].lower()
        tgt_drive = os.path.splitdrive(normalized_target)[0].lower()
        if base_drive != tgt_drive:
            result['same_drive_check'] = False
            result['message'] = f"Different drive: {tgt_drive} vs {base_drive}"
            return result
    
    # Relative path
    try:
        rel = os.path.relpath(normalized_target, start=normalized_base)
        result['relative_path_from_base'] = rel
        if rel == os.pardir or rel.startswith(os.pardir + os.sep):
            result['message'] = f"Outside base: {rel}"
            return result
        result['is_within_base'] = True
        result['is_allowed'] = True
        result['message'] = "Path is allowed"
    except Exception as e:
        result['message'] = f"Relative path error: {e}"
    
    return result


def get_mime_type(file_path: str) -> str:
    """
    Determine the MIME type of a file based on its extension.
    
    Args:
        file_path: Path to the file
        
    Returns:
        MIME type as a string
    """
    # Initialize mimetypes
    if not mimetypes.inited:
        mimetypes.init()
    
    mime_type, _ = mimetypes.guess_type(file_path)
    
    # If mime_type is None, fall back to common types by extension
    if not mime_type:
        ext = os.path.splitext(file_path)[1].lower()
        mime_map = {
            '.txt': 'text/plain',
            '.md': 'text/markdown',
            '.pdf': 'application/pdf',
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            '.csv': 'text/csv',
            '.epub': 'application/epub+zip',
            '.rtf': 'application/rtf',
            '.json': 'application/json',
            '.html': 'text/html',
            '.xml': 'application/xml',
            '.py': 'text/x-python',
            '.js': 'text/javascript',
            '.css': 'text/css',
        }
        mime_type = mime_map.get(ext, 'application/octet-stream')
    
    return mime_type


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


def parse_query_params(query_string: str) -> Dict[str, Any]:
    """
    Parse query string into parameter dictionary, handling booleans and numbers
    
    Args:
        query_string: URL query string (excluding the '?')
        
    Returns:
        Dictionary of parameter values with proper typing
    """
    if not query_string:
        return {}
    
    params = {}
    query_parts = query_string.split('&')
    
    for part in query_parts:
        if '=' not in part:
            continue
        
        key, value = part.split('=', 1)
        key = key.strip()
        value = value.strip()
        
        # Handle boolean values
        if value.lower() in ['true', 'yes', '1']:
            params[key] = True
        elif value.lower() in ['false', 'no', '0']:
            params[key] = False
        # Handle integer values
        elif value.isdigit():
            params[key] = int(value)
        # Handle float values
        elif value.replace('.', '', 1).isdigit() and value.count('.') == 1:
            params[key] = float(value)
        # Default to string
        else:
            params[key] = value
    
    return params


@server.list_resources()
async def list_resources() -> list[types.Resource]:
    """
    Expose resources for file reading and information:
      1. file:///path - Read file content with advanced extraction
      2. file-info:///path - Get file metadata and capabilities
      3. directory:///path - List directory contents
      4. file-options:///path - Get available extraction options for a file
    """
    return [
        types.Resource(
            uri="file:///{path}",
            name="File Content",
            description="Read content from various file types with advanced extraction capabilities. PDF documents include tables, images, and metadata. Office documents include tables, charts, and embedded objects. All file types support optional summarization.",
            schema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string", 
                        "description": "Path to the file relative to allowed directories"
                    },
                    "summarize": {
                        "type": "boolean",
                        "description": "Whether to summarize large content",
                        "default": False
                    },
                    "max_length": {
                        "type": "integer",
                        "description": "Maximum length for summary",
                        "default": 500
                    },
                    "metadata_only": {
                        "type": "boolean",
                        "description": "Extract only metadata (for PDFs, Office docs, EPUB)",
                        "default": False
                    },
                    "tables": {
                        "type": "boolean",
                        "description": "Extract tables (for PDFs, Office docs)",
                        "default": True
                    },
                    "images": {
                        "type": "boolean",
                        "description": "Extract image information",
                        "default": True
                    }
                },
                "required": ["path"]
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
        types.Resource(
            uri="file-info:///{path}",
            name="File Information",
            description="Get metadata and extraction capabilities for a file",
            mimeType="application/json",
            schema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the file relative to current directory"
                    }
                },
                "required": ["path"]
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
        types.Resource(
            uri="directory:///{path}",
            name="Directory Listing",
            description="List files and subdirectories in a directory",
            mimeType="application/json",
            schema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the directory relative to current directory"
                    }
                },
                "required": ["path"]
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
        types.Resource(
            uri="file-options:///{path}",
            name="File Extraction Options",
            description="Get available extraction options for a specific file type",
            mimeType="application/json",
            schema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the file relative to current directory"
                    }
                },
                "required": ["path"]
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
    ]


@server.read_resource()
async def read_resource(uri: AnyUrl) -> str:
    """
    Handle resource URIs, returning file content, information, or directory listings.
    
    Args:
        uri: Resource URI to read
        
    Returns:
        Resource content as a string formatted appropriately for the LLM to understand
        
    Raises:
        Exception: If resource cannot be read or is invalid
    """
    s = str(uri)
    uri_path = ""
    
    # Parse file:/// URI
    if s.startswith("file:///"):
        uri_path = s[len("file:///"):]
        
        # Extract path and query parameters
        query_string = ""
        if "?" in uri_path:
            uri_path, query_string = uri_path.split("?", 1)
        
        path = uri_path
        
        # Parse query parameters
        params = parse_query_params(query_string)
        
        # Extract parameters with defaults
        summarize = params.get("summarize", False)
        max_length = params.get("max_length", 500)
        metadata_only = params.get("metadata_only", False)
        extract_tables = params.get("tables", True)
        extract_images = params.get("images", True)
        
        # Verify path security
        security_check = check_path_security(allowed_directories, path)
        if not security_check["is_allowed"]:
            raise ValueError(f"Access denied: {security_check['message']}")
        
        # Special handling for PDF files with options
        file_extension = os.path.splitext(path)[1].lower()
        
        # Process PDF with special handling if needed
        if file_extension == '.pdf' and metadata_only and has_pymupdf:
            # Custom handling for metadata-only PDF extraction
            try:
                pdf_document = fitz.open(path)
                metadata = pdf_document.metadata
                
                # Format metadata with document info
                content = {
                    "metadata": metadata,
                    "file_info": {
                        "path": path,
                        "pages": len(pdf_document),
                        "pdf_version": pdf_document.pdf_version,
                        "is_encrypted": pdf_document.is_encrypted,
                    }
                }
                return json.dumps(content, indent=2)
            except Exception as e:
                # Fall back to standard extraction if metadata-only fails
                pass
        
        # Standard file content extraction using our utilities
        result = read_file(path, summarize=summarize, max_summary_length=max_length)
        
        if not result["success"]:
            raise ValueError(f"Failed to read file: {result['error']}")
        
        # Return the content directly
        return result["content"]
    
    # Parse file-info:/// URI
    elif s.startswith("file-info:///"):
        uri_path = s[len("file-info:///"):]
        path = uri_path
        
        # Verify path security
        security_check = check_path_security(allowed_directories, path)
        if not security_check["is_allowed"]:
            raise ValueError(f"Access denied: {security_check['message']}")
        
        # Get file info
        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")
        
        # Collect file information
        file_ext = os.path.splitext(path)[1].lower()
        file_stat = os.stat(path)
        
        # Get basic file info
        file_info = {
            "path": path,
            "size": file_stat.st_size,
            "size_human": f"{file_stat.st_size / 1024:.2f} KB",
            "modified": file_stat.st_mtime,
            "created": file_stat.st_ctime,
            "file_type": file_ext,
            "mime_type": get_mime_type(path),
        }
        
        # Add feature detection based on file type
        file_info["capabilities"] = {
            "text_extraction": file_ext in ['.txt', '.md', '.json', '.html', '.xml', '.log', '.py', '.js', '.css', '.java', '.ini', '.conf', '.cfg'],
            "pdf_extraction": file_ext == '.pdf',
            "pdf_tables": file_ext == '.pdf' and (has_tabula or has_pdfplumber),
            "pdf_images": file_ext == '.pdf' and has_pymupdf,
            "docx_extraction": file_ext == '.docx',
            "xlsx_extraction": file_ext == '.xlsx',
            "pptx_extraction": file_ext == '.pptx', 
            "csv_extraction": file_ext == '.csv',
            "epub_extraction": file_ext == '.epub' and has_epub_support,
            "rtf_extraction": file_ext == '.rtf' and has_rtf_support,
            "image_handling": has_pil,
        }
        
        # Add advanced information about available features
        if file_ext == '.pdf':
            file_info["pdf_features"] = {
                "metadata_extraction": True,
                "text_extraction": True,
                "layout_preservation": has_pdfplumber or has_pymupdf,
                "image_extraction": has_pymupdf,
                "table_extraction": has_tabula or has_pdfplumber,
                "annotation_detection": has_pymupdf,
                "fallback_options": ["pdfplumber", "pymupdf", "PyPDF2"],
                "available_libraries": {
                    "pymupdf": has_pymupdf,
                    "pdfplumber": has_pdfplumber,
                    "tabula": has_tabula,
                }
            }
        elif file_ext == '.docx':
            file_info["docx_features"] = {
                "metadata_extraction": True,
                "text_extraction": True,
                "table_extraction": True,
                "image_detection": True,
                "heading_detection": True
            }
        elif file_ext == '.xlsx':
            file_info["xlsx_features"] = {
                "metadata_extraction": True,
                "sheet_extraction": True,
                "table_formatting": True,
                "image_detection": True,
                "chart_detection": True
            }
        elif file_ext == '.pptx':
            file_info["pptx_features"] = {
                "metadata_extraction": True,
                "slide_extraction": True,
                "text_extraction": True,
                "image_detection": True,
                "shape_detection": True,
                "chart_detection": True
            }
        elif file_ext == '.epub':
            file_info["epub_features"] = {
                "metadata_extraction": has_epub_support,
                "toc_extraction": has_epub_support,
                "content_extraction": has_epub_support,
                "available": has_epub_support
            }
        
        # Get available extraction options
        file_info["extraction_options"] = get_file_extraction_options(file_ext)
        
        # Return JSON string directly
        return json.dumps(file_info, indent=2)
    
    # Parse file-options:/// URI - Get available extraction options for a file
    elif s.startswith("file-options:///"):
        uri_path = s[len("file-options:///"):]
        path = uri_path
        
        # Verify path security
        security_check = check_path_security(allowed_directories, path)
        if not security_check["is_allowed"]:
            raise ValueError(f"Access denied: {security_check['message']}")
        
        # Check if file exists
        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")
        
        # Get file extension
        file_ext = os.path.splitext(path)[1].lower()
        
        # Get extraction options
        options = get_file_extraction_options(file_ext)
        
        # Return options as JSON string
        return json.dumps(options, indent=2)
    
    # Parse directory:/// URI
    elif s.startswith("directory:///"):
        uri_path = s[len("directory:///"):]
        path = uri_path
        
        # Verify path security
        security_check = check_path_security(allowed_directories, path)
        if not security_check["is_allowed"]:
            raise ValueError(f"Access denied: {security_check['message']}")
        
        # List directory contents
        if not os.path.isdir(path):
            raise NotADirectoryError(f"Directory not found: {path}")
        
        # Get directory contents with metadata
        contents = []
        for item in os.listdir(path):
            item_path = os.path.join(path, item)
            is_dir = os.path.isdir(item_path)
            
            # Get basic item info
            item_info = {
                "name": item,
                "path": item_path,
                "is_dir": is_dir,
            }
            
            # For files, add additional info
            if not is_dir:
                try:
                    file_size = os.path.getsize(item_path)
                    item_info.update({
                        "size": file_size,
                        "size_human": f"{file_size / 1024:.2f} KB",
                        "extension": os.path.splitext(item)[1].lower(),
                        "mime_type": get_mime_type(item_path),
                    })
                except:
                    # If we can't get file info, provide minimal data
                    item_info.update({
                        "size": None,
                        "extension": os.path.splitext(item)[1].lower(),
                    })
            
            contents.append(item_info)
        
        # Sort contents: directories first, then files alphabetically
        contents.sort(key=lambda x: (not x["is_dir"], x["name"].lower()))
        
        # Return directory listing as JSON string
        return json.dumps(contents, indent=2)
    
    # Invalid URI
    else:
        raise ValueError(f"Unsupported resource URI: {uri}")


@server.list_tools()
async def list_tools() -> list[types.Tool]:
    """
    Returns a list of available tools for file operations.
    """
    tools = [
        types.Tool(
            name="read_file",
            description="Read a file with advanced extraction capabilities",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the file to read"
                    },
                    "summarize": {
                        "type": "boolean",
                        "description": "Whether to summarize large content",
                        "default": False
                    },
                    "max_length": {
                        "type": "integer",
                        "description": "Maximum length for summary if summarizing",
                        "default": 500
                    },
                    "metadata_only": {
                        "type": "boolean",
                        "description": "Extract only metadata (for PDFs, Office docs)",
                        "default": False
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        ),
        types.Tool(
            name="list_directory",
            description="List files and folders in a directory",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the directory to list"
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        ),
        types.Tool(
            name="get_file_info",
            description="Get detailed information about a file including supported extraction capabilities",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the file to get information about"
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        )
    ]
    
    # Add a new tool to list allowed directories
    tools.append(
        types.Tool(
            name="list_allowed_directories",
            description="List directories that this server is allowed to access",
            inputSchema={
                "type": "object",
                "properties": {},
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        )
    )
    
    return tools


@server.call_tool()
async def call_tool(tool_name: str, arguments: dict) -> list[types.TextContent]:
    """
    Call a tool by name with arguments.
    """
    if tool_name == "read_file":
        path = arguments["path"]
        summarize = arguments.get("summarize", False)
        max_length = arguments.get("max_length", 500)
        metadata_only = arguments.get("metadata_only", False)

        # Verify path security
        security_check = check_path_security(allowed_directories, path)
        if not security_check["is_allowed"]:
            return [types.TextContent(
                type="text",
                text=f"Access denied: {security_check['message']}"
            )]

        # Special handling for PDFs with metadata_only
        file_extension = os.path.splitext(path)[1].lower()
        if file_extension == '.pdf' and metadata_only and has_pymupdf:
            try:
                pdf_document = fitz.open(path)
                metadata = pdf_document.metadata
                
                # Format metadata with document info
                content = {
                    "metadata": metadata,
                    "file_info": {
                        "path": path,
                        "pages": len(pdf_document),
                        "pdf_version": pdf_document.pdf_version,
                        "is_encrypted": pdf_document.is_encrypted,
                    }
                }
                return [types.TextContent(
                    type="text",
                    text=f"PDF Metadata:\n\n{json.dumps(content, indent=2)}"
                )]
            except Exception as e:
                # Fall back to standard extraction if metadata-only fails
                return [types.TextContent(
                    type="text",
                    text=f"Error extracting PDF metadata: {str(e)}"
                )]
        
        # Special handling for DOCX with metadata_only
        elif file_extension == '.docx' and metadata_only:
            try:
                try:
                    doc = docx.Document(path)
                    metadata = {}
                    
                    # Extract core properties
                    if hasattr(doc, 'core_properties'):
                        core_props = doc.core_properties
                        if hasattr(core_props, 'title') and core_props.title:
                            metadata["title"] = core_props.title
                        if hasattr(core_props, 'author') and core_props.author:
                            metadata["author"] = core_props.author
                        if hasattr(core_props, 'created') and core_props.created:
                            metadata["created"] = str(core_props.created)
                        if hasattr(core_props, 'modified') and core_props.modified:
                            metadata["modified"] = str(core_props.modified)
                        if hasattr(core_props, 'comments') and core_props.comments:
                            metadata["comments"] = core_props.comments
                        if hasattr(core_props, 'category') and core_props.category:
                            metadata["category"] = core_props.category
                        if hasattr(core_props, 'subject') and core_props.subject:
                            metadata["subject"] = core_props.subject
                        if hasattr(core_props, 'keywords') and core_props.keywords:
                            metadata["keywords"] = core_props.keywords
                    
                    # Document statistics
                    stats = {
                        "paragraphs": len(doc.paragraphs),
                        "sections": len(doc.sections),
                        "tables": len(doc.tables)
                    }
                    
                    # Get image count
                    image_count = 0
                    for rel in doc.part.rels.values():
                        if rel.reltype == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image':
                            image_count += 1
                    stats["images"] = image_count
                    
                    # Format the result
                    content = {
                        "metadata": metadata,
                        "document_statistics": stats,
                        "file_info": {
                            "path": path
                        }
                    }
                    
                    return [types.TextContent(
                        type="text",
                        text=f"DOCX Metadata:\n\n{json.dumps(content, indent=2)}"
                    )]
                except Exception as e:
                    return [types.TextContent(
                        type="text",
                        text=f"Error extracting DOCX metadata: {str(e)}"
                    )]
            except Exception as e:
                return [types.TextContent(
                    type="text",
                    text=f"Error: {str(e)}"
                )]

        # Standard file extraction using our utilities
        result = patched_read_file(path, summarize=summarize, max_summary_length=max_length)
        
        if not result["success"]:
            return [types.TextContent(
                type="text",
                text=f"Failed to read file: {result['error']}"
            )]
        
        mime_type = get_mime_type(path)
        
        return [types.TextContent(
            type="text",
            text=result["content"]
        )]

    elif tool_name == "list_directory":
        path = arguments["path"]
        
        # Verify path security
        security_check = check_path_security(allowed_directories, path)
        if not security_check["is_allowed"]:
            return [types.TextContent(
                type="text",
                text=f"Access denied: {security_check['message']}"
            )]
        
        # List directory contents
        if not os.path.isdir(path):
            return [types.TextContent(
                type="text",
                text=f"Directory not found: {path}"
            )]
        
        # Get directory contents with metadata
        contents = []
        for item in os.listdir(path):
            item_path = os.path.join(path, item)
            is_dir = os.path.isdir(item_path)
            
            # Get basic item info
            item_info = {
                "name": item,
                "path": item_path,
                "is_dir": is_dir,
            }
            
            # For files, add additional info
            if not is_dir:
                try:
                    file_size = os.path.getsize(item_path)
                    item_info.update({
                        "size": file_size,
                        "size_human": f"{file_size / 1024:.2f} KB",
                        "extension": os.path.splitext(item)[1].lower(),
                        "mime_type": get_mime_type(item_path),
                    })
                except:
                    # If we can't get file info, provide minimal data
                    item_info.update({
                        "size": None,
                        "extension": os.path.splitext(item)[1].lower(),
                    })
            
            contents.append(item_info)
        
        # Sort contents: directories first, then files alphabetically
        contents.sort(key=lambda x: (not x["is_dir"], x["name"].lower()))
        
        formatted_listing = f"Directory listing for {path}:\n\n"
        
        # Add directories
        formatted_listing += "Directories:\n"
        dirs = [item for item in contents if item["is_dir"]]
        if dirs:
            for dir_item in dirs:
                formatted_listing += f"  📁 {dir_item['name']}/\n"
        else:
            formatted_listing += "  (No directories)\n"
            
        # Add files
        formatted_listing += "\nFiles:\n"
        files = [item for item in contents if not item["is_dir"]]
        if files:
            for file_item in files:
                size_info = f" ({file_item.get('size_human', 'unknown size')})"
                formatted_listing += f"  📄 {file_item['name']}{size_info}\n"
        else:
            formatted_listing += "  (No files)\n"

        return [types.TextContent(
            type="text",
            text=formatted_listing
        )]
        
    elif tool_name == "get_file_info":
        path = arguments["path"]
        
        # Verify path security
        security_check = check_path_security(allowed_directories, path)
        if not security_check["is_allowed"]:
            return [types.TextContent(
                type="text",
                text=f"Access denied: {security_check['message']}"
            )]
        
        # Get file info
        if not os.path.exists(path):
            return [types.TextContent(
                type="text",
                text=f"File not found: {path}"
            )]
        
        # Collect file information
        file_ext = os.path.splitext(path)[1].lower()
        file_stat = os.stat(path)
        
        # Get basic file info
        file_info = {
            "path": path,
            "size": file_stat.st_size,
            "size_human": f"{file_stat.st_size / 1024:.2f} KB",
            "modified": file_stat.st_mtime,
            "created": file_stat.st_ctime,
            "file_type": file_ext,
            "mime_type": get_mime_type(path),
        }
        
        # Add feature detection based on file type
        file_info["capabilities"] = {
            "text_extraction": file_ext in ['.txt', '.md', '.json', '.html', '.xml', '.log', '.py', '.js', '.css', '.java', '.ini', '.conf', '.cfg'],
            "pdf_extraction": file_ext == '.pdf',
            "pdf_tables": file_ext == '.pdf' and (has_tabula or has_pdfplumber),
            "pdf_images": file_ext == '.pdf' and has_pymupdf,
            "docx_extraction": file_ext == '.docx',
            "xlsx_extraction": file_ext == '.xlsx',
            "pptx_extraction": file_ext == '.pptx', 
            "csv_extraction": file_ext == '.csv',
            "epub_extraction": file_ext == '.epub' and has_epub_support,
            "rtf_extraction": file_ext == '.rtf' and has_rtf_support,
            "image_handling": has_pil,
        }
        
        # Generate human-friendly output
        formatted_info = f"File Information for {path}:\n\n"
        formatted_info += f"File Size: {file_info['size_human']}\n"
        formatted_info += f"Type: {file_info['file_type']} ({file_info['mime_type']})\n"
        formatted_info += f"Modified: {time.ctime(file_info['modified'])}\n"
        formatted_info += f"Created: {time.ctime(file_info['created'])}\n\n"
        
        formatted_info += "Available Extraction Features:\n"
        
        for cap, available in file_info["capabilities"].items():
            if available:
                formatted_info += f"✅ {cap.replace('_', ' ').title()}\n"
            else:
                formatted_info += f"❌ {cap.replace('_', ' ').title()}\n"
        
        return [types.TextContent(
            type="text",
            text=formatted_info
        )]
    
    # New tool to list allowed directories
    elif tool_name == "list_allowed_directories":
        return [types.TextContent(
            type="text",
            text=f"Allowed directories:\n{json.dumps(allowed_directories, indent=2)}"
        )]
    
    # Unknown tool
    return [types.TextContent(
        type="text",
        text=f"Unknown tool: {tool_name}"
    )]


async def main() -> None:
    # Parse command-line arguments for allowed directories
    global allowed_directories
    if len(sys.argv) > 1:
        allowed_directories = [os.path.abspath(os.path.normpath(dir_path)) for dir_path in sys.argv[1:]]
    
    print(f"Starting MCP file-system server with allowed directories:", file=sys.stderr)
    for directory in allowed_directories:
        print(f"  - {directory}", file=sys.stderr)
    
    async with stdio_server() as streams:
        await server.run(
            read_stream=streams[0],
            write_stream=streams[1],
            initialization_options=server.create_initialization_options()
        )


if __name__ == "__main__":
    asyncio.run(main())