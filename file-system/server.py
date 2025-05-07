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

# Import our file extraction utilities
from resources import extract_file_content, read_file
from resources.readers import (
    read_text_file, read_pdf_file, read_docx_file, read_xlsx_file,
    read_pptx_file, read_csv_file, read_epub_file, read_rtf_file,
    has_pymupdf, has_tabula, has_pdfplumber, has_pil,
    has_epub_support, has_rtf_support
)
from resources.utils.formatters import summarize_content

server = Server(
    name="file-system",
    instructions="You are a file system MCP server. You can read and extract content from various file types including text files, PDFs, Office documents (Word, Excel, PowerPoint), CSV, EPUB, and RTF. The server provides advanced extraction capabilities for tables, images, and metadata. PDF documents will include information about tables, images, and annotations when available. Office documents will extract embedded images, charts, tables, and preserve formatting where possible.",
    version="0.1",
)


def check_path_security(allowed_base_path: str, target_path: str) -> dict:
    """
    Checks if a target_path is allowed based on the allowed_base_path.

    This involves:
    1. Normalizing and resolving real paths for both inputs.
    2. Checking if the normalized paths exist on the filesystem.
    3. On Windows, checking if paths are on the same drive.
    4. Checking if the target path is a child path (or the same) as the base path
       using a secure relative path comparison.

    Args:
        allowed_base_path: The root directory path that is permitted.
        
        target_path: The path to a file or directory to check.
                     

    Returns:
        A dictionary containing detailed results of the checks:
        - 'original_base': The input allowed_base_path string.
        - 'original_target': The input target_path string.
        - 'normalized_base': The normalized and real path of the base, or None if normalization failed.
        - 'normalized_target': The normalized and real path of the target, or None if normalization failed.
        - 'normalization_successful': bool, True if normalization succeeded.
        - 'base_exists': bool, True if normalized_base_path exists.
        - 'target_exists': bool, True if normalized_target_path exists.
        - 'same_drive_check': bool, True if paths are on the same drive (Windows only, always True otherwise).
        - 'relative_path_from_base': str, The path from normalized_base to normalized_target, or None if not applicable.
        - 'is_within_base': bool, True if the relative path doesn't go outside the base.
        - 'is_allowed': bool, The final decision (True if all checks pass).
        - 'message': str, A human-readable summary of the result.
    """
    results = {
        'original_base': allowed_base_path,
        'original_target': target_path,
        'normalized_base': None,
        'normalized_target': None,
        'normalization_successful': False,
        'base_exists': False,
        'target_exists': False,
        'same_drive_check': True,
        'relative_path_from_base': None,
        'is_within_base': False,
        'is_allowed': False,
        'message': ''
    }

    try:
        normalized_base = os.path.realpath(os.path.normpath(allowed_base_path))
        normalized_target = os.path.realpath(os.path.normpath(target_path))
        results['normalized_base'] = normalized_base
        results['normalized_target'] = normalized_target
        results['normalization_successful'] = True
    except Exception as e:
        results['message'] = f"Normalization error: {e}"
        return results

    if not os.path.exists(normalized_base) or not os.path.isdir(normalized_base):
        results['message'] = f"Base path does not exist or not a directory: {normalized_base}"
        return results
    results['base_exists'] = True

    if not os.path.exists(normalized_target):
        results['message'] = f"Target does not exist: {normalized_target}"
        return results
    results['target_exists'] = True

    # Windows drive check
    if sys.platform == 'win32':
        base_drive = os.path.splitdrive(normalized_base)[0].lower()
        tgt_drive = os.path.splitdrive(normalized_target)[0].lower()
        if base_drive != tgt_drive:
            results['same_drive_check'] = False
            results['message'] = f"Different drive: {tgt_drive} vs {base_drive}"
            return results

    # Relative path
    try:
        rel = os.path.relpath(normalized_target, start=normalized_base)
        results['relative_path_from_base'] = rel
        if rel == os.pardir or rel.startswith(os.pardir + os.sep):
            results['message'] = f"Outside base: {rel}"
            return results
        results['is_within_base'] = True
        results['is_allowed'] = True
        results['message'] = "Path is allowed"
    except Exception as e:
        results['message'] = f"Relative path error: {e}"
        return results

    return results


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
    base_dir = os.getcwd()  # Current working directory as base
    
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
                        "description": "Path to the file relative to current directory"
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
    base_dir = os.getcwd()
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
        security_check = check_path_security(base_dir, path)
        if not security_check["is_allowed"]:
            raise ValueError(f"Access denied: {security_check['message']}")
        
        # Special handling for PDF files with options
        file_extension = os.path.splitext(path)[1].lower()
        
        # Process PDF with special handling if needed
        if file_extension == '.pdf' and metadata_only and has_pymupdf:
            # Custom handling for metadata-only PDF extraction
            import fitz  # PyMuPDF
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
        security_check = check_path_security(base_dir, path)
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
        security_check = check_path_security(base_dir, path)
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
        security_check = check_path_security(base_dir, path)
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
    return [
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


@server.call_tool()
async def call_tool(tool_name: str, arguments: dict) -> list[types.TextContent]:
    """
    Call a tool by name with arguments.
    """
    base_dir = os.getcwd()

    if tool_name == "read_file":
        path = arguments["path"]
        summarize = arguments.get("summarize", False)
        max_length = arguments.get("max_length", 500)
        metadata_only = arguments.get("metadata_only", False)

        # Verify path security
        security_check = check_path_security(base_dir, path)
        if not security_check["is_allowed"]:
            return [types.TextContent(
                type="text",
                text=f"Access denied: {security_check['message']}"
            )]

        # Special handling for PDFs with metadata_only
        file_extension = os.path.splitext(path)[1].lower()
        if file_extension == '.pdf' and metadata_only and has_pymupdf:
            import fitz  # PyMuPDF
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

        # Standard file extraction
        result = read_file(path, summarize=summarize, max_summary_length=max_length)
        
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
        security_check = check_path_security(base_dir, path)
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
        security_check = check_path_security(base_dir, path)
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
    
    # Unknown tool
    return [types.TextContent(
        type="text",
        text=f"Unknown tool: {tool_name}"
    )]


async def main() -> None:
    async with stdio_server() as streams:
        await server.run(
            read_stream=streams[0],
            write_stream=streams[1],
            initialization_options=server.create_initialization_options()
        )


if __name__ == "__main__":
    asyncio.run(main())