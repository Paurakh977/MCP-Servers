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
from datetime import datetime
#checking some crusial imports because in the client config file the env might not have the following dependencies
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
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import range_boundaries
    has_openpyxl = True
except ImportError:
    has_openpyxl = False
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

from resources.core import (
    # Feature flags
    has_pymupdf, has_pdfplumber, has_tabula,
    has_pil, has_epub_support, has_rtf_support,
    
    # File utilities
    check_path_security,
    check_against_single_base,
    get_mime_type,
    parse_query_params,
    
    # Extraction options
    get_file_extraction_options
)

from resources import extract_file_content, read_file
from resources.readers import (
    read_text_file, read_pdf_file, read_docx_file, read_xlsx_file,
    read_pptx_file, read_csv_file, read_epub_file, read_rtf_file
)
from resources.utils.formatters import summarize_content

# Initialize allowed directories list - will be populated with command-line arguments
allowed_directories = [os.getcwd()]  # Default to current directory

server = Server(
    name="file-system",
    instructions="You are a file system MCP server. You can read and extract content from various file types including text files, PDFs, Office documents (Word, Excel, PowerPoint), CSV, EPUB, and RTF. The server provides advanced extraction capabilities for tables, images, and metadata. PDF documents will include information about tables, images, and annotations when available. Office documents will extract embedded images, charts, tables, and preserve formatting where possible.",
    version="0.1",
)

@server.list_resources()
async def list_resources() -> list[types.Resource]:
    """
    Expose resources for file reading and information:
      1. file:///path - Read file content with advanced extraction
      2. excel-info:///path - Get Excel workbook metadata and sheet information
      3. excel-sheet:///path - Read specific Excel sheet content
      4. file-info:///path - Get file metadata and capabilities
      5. directory:///path - List directory contents
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
                    }
                },
                "required": ["path"]
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
        types.Resource(
            uri="excel-info:///{path}",
            name="Excel Workbook Information",
            description="Get metadata and sheet information from Excel files",
            mimeType="application/json",
            schema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel file"
                    }
                },
                "required": ["path"]
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
        types.Resource(
            uri="excel-sheet:///{path}",
            name="Excel Sheet Content",
            description="Read content from a specific sheet in an Excel file",
            mimeType="application/json",
            schema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel file"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the sheet to read"
                    },
                    "cell_range": {
                        "type": "string",
                        "description": "Optional cell range to read (e.g. 'A1:D10')",
                        "default": None
                    }
                },
                "required": ["path", "sheet_name"]
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
    ]

@server.read_resource()
async def read_resource(uri: AnyUrl) -> str:
    """Handle resource URIs, returning file content, information, or directory listings."""
    try:
        s = str(uri)
        
        # Parse file:/// URI
        if s.startswith("file:///"):
            try:
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
                
                # Verify path security
                security_check = check_path_security(allowed_directories, path)
                if not security_check["is_allowed"]:
                    return json.dumps({
                        "error": f"Access denied: {security_check['message']}"
                    })
                
                # Standard file content extraction
                result = read_file(path, summarize=summarize, max_summary_length=max_length)
                
                if not result["success"]:
                    return json.dumps({
                        "error": result["error"]
                    })
                
                return json.dumps({
                    "content": result["content"],
                    "file_type": result["file_type"],
                    "file_path": result["file_path"]
                })
                
            except Exception as e:
                return json.dumps({
                    "error": f"Failed to process file: {str(e)}"
                })
        
        # Parse excel-info:/// URI
        elif s.startswith("excel-info:///"):
            try:
                uri_path = s[len("excel-info:///"):]
                path = uri_path
                
                # Verify path security
                security_check = check_path_security(allowed_directories, path)
                if not security_check["is_allowed"]:
                    return json.dumps({
                        "error": f"Access denied: {security_check['message']}"
                    })
                
                # Verify file exists and is Excel
                if not os.path.exists(path):
                    return json.dumps({
                        "error": f"File not found: {path}"
                    })
                
                file_extension = os.path.splitext(path)[1].lower()
                if file_extension != '.xlsx':
                    return json.dumps({
                        "error": f"Not an Excel file: {path}"
                    })
                
                # Get Excel workbook info
                try:
                    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
                    
                    # Format the output in a more readable way
                    output_lines = []
                    output_lines.append(f"Excel File: {os.path.basename(path)}")
                    output_lines.append(f"Number of Sheets: {len(workbook.worksheets)}")
                    output_lines.append("\nSheet Information:")
                    
                    for sheet in workbook.worksheets:
                        # Get sheet dimensions
                        min_col, min_row, max_col, max_row = range_boundaries(sheet.calculate_dimension())
                        
                        output_lines.append(f"\n📄 Sheet: {sheet.title}")
                        output_lines.append(f"   Dimensions: {sheet.calculate_dimension()}")
                        output_lines.append(f"   Rows: {sheet.max_row}")
                        output_lines.append(f"   Columns: {max_col - min_col + 1}")
                        
                        # Get column headers
                        columns = []
                        column_refs = []
                        for col in range(min_col, max_col + 1):
                            cell = sheet.cell(min_row, col)
                            header = cell.value if cell.value is not None else f"Column {get_column_letter(col)}"
                            columns.append(header)
                            column_refs.append(get_column_letter(col))
                        
                        # Format column information
                        output_lines.append("   Columns:")
                        for i, (col_ref, col_name) in enumerate(zip(column_refs, columns)):
                            output_lines.append(f"     {col_ref}: {col_name}")
                    
                    workbook.close()
                    return json.dumps({
                        "content": "\n".join(output_lines)
                    })
                    
                except Exception as e:
                    return json.dumps({
                        "error": f"Failed to read Excel file: {str(e)}"
                    })
            
            except Exception as e:
                return json.dumps({
                    "error": f"Failed to process Excel info: {str(e)}"
                })
        
        # Parse excel-sheet:/// URI
        elif s.startswith("excel-sheet:///"):
            try:
                uri_path = s[len("excel-sheet:///"):]
                
                # Extract path and query parameters
                query_string = ""
                if "?" in uri_path:
                    uri_path, query_string = uri_path.split("?", 1)
                
                path = uri_path
                
                # Parse query parameters
                params = parse_query_params(query_string)
                
                # Get required sheet_name and optional cell_range
                sheet_name = params.get("sheet_name")
                if not sheet_name:
                    return json.dumps({
                        "error": "sheet_name parameter is required"
                    })
                
                cell_range = params.get("cell_range")
                
                # Verify path security
                security_check = check_path_security(allowed_directories, path)
                if not security_check["is_allowed"]:
                    return json.dumps({
                        "error": f"Access denied: {security_check['message']}"
                    })
                
                # Verify file exists and is Excel
                if not os.path.exists(path):
                    return json.dumps({
                        "error": f"File not found: {path}"
                    })
                
                file_extension = os.path.splitext(path)[1].lower()
                if file_extension != '.xlsx':
                    return json.dumps({
                        "error": f"Not an Excel file: {path}"
                    })
                
                # Read the specific sheet
                try:
                    workbook = openpyxl.load_workbook(path, data_only=True)
                    
                    # Validate sheet exists
                    if sheet_name not in workbook.sheetnames:
                        workbook.close()
                        return json.dumps({
                            "error": f"Sheet '{sheet_name}' not found in workbook"
                        })
                    
                    sheet = workbook[sheet_name]
                    
                    # Parse cell range if provided
                    if cell_range:
                        try:
                            min_col, min_row, max_col, max_row = range_boundaries(cell_range)
                        except ValueError:
                            workbook.close()
                            return json.dumps({
                                "error": f"Invalid cell range format: {cell_range}"
                            })
                    else:
                        # Use full sheet range
                        min_col, min_row, max_col, max_row = range_boundaries(sheet.calculate_dimension())
                    
                    # Get column headers (first row)
                    columns = []
                    for col in range(min_col, max_col + 1):
                        cell = sheet.cell(min_row, col)
                        header = cell.value if cell.value is not None else f"Column {get_column_letter(col)}"
                        columns.append(header)
                    
                    # Collect non-empty rows
                    records = []
                    for row in range(min_row + 1, max_row + 1):
                        values = []
                        row_has_data = False
                        
                        for col in range(min_col, max_col + 1):
                            cell = sheet.cell(row, col)
                            value = cell.value
                            
                            # Convert datetime objects to ISO format
                            if isinstance(value, datetime):
                                value = value.isoformat()
                            
                            values.append(value)
                            if value is not None and value != "":
                                row_has_data = True
                        
                        if row_has_data:
                            records.append({
                                "row": row,
                                "values": values,
                                "cell_refs": [f"{get_column_letter(col)}{row}" for col in range(min_col, max_col + 1)]
                            })
                    
                    # Count charts and images
                    chart_count = len([drawing for drawing in sheet._charts])
                    image_count = len([drawing for drawing in sheet._images])
                    
                    # Prepare sheet data
                    sheet_data = {
                        "sheet_name": sheet_name,
                        "dimensions": cell_range or sheet.calculate_dimension(),
                        "non_empty_cells": len(records),
                        "charts": chart_count,
                        "images": image_count,
                        "columns": columns,
                        "column_refs": [get_column_letter(col) for col in range(min_col, max_col + 1)],
                        "records": records,
                        "metadata": {
                            "total_rows": max_row - min_row + 1,
                            "total_columns": max_col - min_col + 1,
                            "visible_rows": len(records),
                            "has_charts": chart_count > 0,
                            "has_images": image_count > 0
                        }
                    }
                    
                    workbook.close()
                    return json.dumps(sheet_data)
                    
                except Exception as e:
                    if 'workbook' in locals():
                        workbook.close()
                    return json.dumps({
                        "error": f"Failed to read Excel sheet: {str(e)}"
                    })
            
            except Exception as e:
                return json.dumps({
                    "error": f"Failed to process Excel sheet: {str(e)}"
                })
        
        # Parse file-info:/// URI
        elif s.startswith("file-info:///"):
            try:
                uri_path = s[len("file-info:///"):]
                path = uri_path
                
                # Verify path security
                security_check = check_path_security(allowed_directories, path)
                if not security_check["is_allowed"]:
                    return json.dumps({
                        "error": f"Access denied: {security_check['message']}"
                    })
                
                # Get file info
                if not os.path.exists(path):
                    return json.dumps({
                        "error": f"File not found: {path}"
                    })
                
                # Get file info as JSON
                file_info = get_file_info(path)
                return json.dumps(file_info)
                
            except Exception as e:
                return json.dumps({
                    "error": f"Failed to get file info: {str(e)}"
                })
        
        # Parse directory:/// URI
        elif s.startswith("directory:///"):
            try:
                uri_path = s[len("directory:///"):]
                path = uri_path
                
                # Verify path security
                security_check = check_path_security(allowed_directories, path)
                if not security_check["is_allowed"]:
                    return json.dumps({
                        "error": f"Access denied: {security_check['message']}"
                    })
                
                # List directory contents
                if not os.path.isdir(path):
                    return json.dumps({
                        "error": f"Directory not found: {path}"
                    })
                
                # Get directory listing as JSON
                dir_listing = get_directory_listing(path)
                return json.dumps(dir_listing)
                
            except Exception as e:
                return json.dumps({
                    "error": f"Failed to list directory: {str(e)}"
                })
        
        # Invalid URI
        else:
            return json.dumps({
                "error": f"Unsupported resource URI: {uri}"
            })
            
    except Exception as e:
        return json.dumps({
            "error": f"Invalid URI or unexpected error: {str(e)}"
        })

@server.list_tools()
async def list_tools() -> list[types.Tool]:
    """
    Returns a list of available tools for file operations.
    """
    return [
        types.Tool(
            name="read_file",
            description="Read a file with advanced extraction capabilities. You can provide either an absolute path, "
                       "relative path, or just the filename - the server will search for matching files in allowed directories. "
                       "For example: 'report.pdf', 'downloads/data.xlsx', or just 'summary' to find files containing that name.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path or name of the file to read. Can be absolute path, relative path, or filename to search for."
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
            name="get_excel_info",
            description="Get metadata and sheet information from an Excel file. You can provide either an absolute path, "
                       "relative path, or just the filename - the server will search for matching Excel files in allowed directories. "
                       "For example: 'data.xlsx', 'reports/summary.xlsx', or just 'sales' to find Excel files containing that name.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path or name of the Excel file. Can be absolute path, relative path, or filename to search for."
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        ),
        types.Tool(
            name="read_excel_sheet",
            description="Read content from a specific sheet in an Excel file. You can provide either an absolute path, "
                       "relative path, or just the filename - the server will search for matching Excel files in allowed directories. "
                       "For example: 'data.xlsx', 'reports/summary.xlsx', or just 'sales' to find Excel files containing that name.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path or name of the Excel file. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the sheet to read"
                    },
                    "cell_range": {
                        "type": "string",
                        "description": "Optional cell range to read (e.g. 'A1:D10')",
                        "default": None
                    }
                },
                "required": ["path", "sheet_name"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        ),
        types.Tool(
            name="list_directory",
            description="List files and folders in a directory. You can provide either an absolute path, relative path, "
                       "or just the directory name - the server will search for matching directories in allowed locations. "
                       "For example: 'downloads', 'documents/reports', or just 'data' to find directories containing that name.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path or name of the directory. Can be absolute path, relative path, or directory name to search for."
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
            description="Get detailed information about a file including supported extraction capabilities. You can provide either "
                       "an absolute path, relative path, or just the filename - the server will search for matching files in allowed directories. "
                       "For example: 'document.pdf', 'reports/data.xlsx', or just 'summary' to find files containing that name.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path or name of the file. Can be absolute path, relative path, or filename to search for."
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        ),
        types.Tool(
            name="list_allowed_directories",
            description="List directories that this server is allowed to access. Use this to understand where the server can search for files.",
            inputSchema={
                "type": "object",
                "properties": {},
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        )
    ]

def find_file_in_allowed_dirs(file_pattern: str, allowed_dirs: list[str]) -> Optional[str]:
    """
    Search for a file matching the pattern in allowed directories and their subdirectories.
    Returns the absolute path if found, None otherwise.
    
    Args:
        file_pattern (str): File name or pattern to search for (case-insensitive)
        allowed_dirs (list[str]): List of allowed directory paths to search in
    """
    file_pattern = file_pattern.lower()
    found_files = []
    
    for base_dir in allowed_dirs:
        for root, _, files in os.walk(base_dir):
            for filename in files:
                if file_pattern in filename.lower():
                    abs_path = os.path.abspath(os.path.join(root, filename))
                    found_files.append({
                        "path": abs_path,
                        "similarity": filename.lower().count(file_pattern) / len(filename)
                    })
    
    # Sort by similarity score (higher is better)
    found_files.sort(key=lambda x: x["similarity"], reverse=True)
    
    return found_files[0]["path"] if found_files else None

@server.call_tool()
async def call_tool(tool_name: str, arguments: dict) -> list[types.TextContent]:
    """Call a tool by name with arguments."""
    try:
        if tool_name == "read_file":
            path = arguments["path"]
            summarize = arguments.get("summarize", False)
            max_length = arguments.get("max_length", 500)
            metadata_only = arguments.get("metadata_only", False)

            # If path doesn't exist, try to find it in allowed directories
            if not os.path.exists(path):
                found_path = find_file_in_allowed_dirs(path, allowed_directories)
                if found_path:
                    path = found_path
                    print(f"Found file at: {path}", file=sys.stderr)
                else:
                    return [types.TextContent(
                        type="text",
                        text=f"Could not find file matching '{path}' in allowed directories: {allowed_directories}"
                    )]

            # Verify path security
            security_check = check_path_security(allowed_directories, path)
            if not security_check["is_allowed"]:
                return [types.TextContent(
                    type="text",
                    text=f"Access denied: {security_check['message']}"
                )]

            # Standard file content extraction
            result = read_file(path, summarize=summarize, max_summary_length=max_length)
            
            if not result["success"]:
                return [types.TextContent(
                    type="text",
                    text=f"Failed to read file: {result['error']}"
                )]
            
            return [types.TextContent(
                type="text",
                text=result["content"]
            )]
        
        elif tool_name == "get_excel_info":
            path = arguments["path"]
            
            # If path doesn't exist, try to find it in allowed directories
            if not os.path.exists(path):
                found_path = find_file_in_allowed_dirs(path, allowed_directories)
                if found_path:
                    path = found_path
                    print(f"Found Excel file at: {path}", file=sys.stderr)
                else:
                    return [types.TextContent(
                        type="text",
                        text=f"Could not find Excel file matching '{path}' in allowed directories: {allowed_directories}"
                    )]
            
            # Verify path security
            security_check = check_path_security(allowed_directories, path)
            if not security_check["is_allowed"]:
                return [types.TextContent(
                    type="text",
                    text=f"Access denied: {security_check['message']}"
                )]
            
            # Verify file exists and is Excel
            if not os.path.exists(path):
                return [types.TextContent(
                    type="text",
                    text=f"File not found: {path}"
                )]
            
            file_extension = os.path.splitext(path)[1].lower()
            if file_extension != '.xlsx':
                return [types.TextContent(
                    type="text",
                    text=f"Not an Excel file: {path}"
                )]
            
            # Get Excel workbook info
            try:
                workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
                
                # Format the output in a more readable way
                output_lines = []
                output_lines.append(f"Excel File: {os.path.basename(path)}")
                output_lines.append(f"Number of Sheets: {len(workbook.worksheets)}")
                output_lines.append("\nSheet Information:")
                
                for sheet in workbook.worksheets:
                    # Get sheet dimensions
                    min_col, min_row, max_col, max_row = range_boundaries(sheet.calculate_dimension())
                    
                    output_lines.append(f"\n📄 Sheet: {sheet.title}")
                    output_lines.append(f"   Dimensions: {sheet.calculate_dimension()}")
                    output_lines.append(f"   Rows: {sheet.max_row}")
                    output_lines.append(f"   Columns: {max_col - min_col + 1}")
                    
                    # Get column headers
                    columns = []
                    column_refs = []
                    for col in range(min_col, max_col + 1):
                        cell = sheet.cell(min_row, col)
                        header = cell.value if cell.value is not None else f"Column {get_column_letter(col)}"
                        columns.append(header)
                        column_refs.append(get_column_letter(col))
                    
                    # Format column information
                    output_lines.append("   Columns:")
                    for i, (col_ref, col_name) in enumerate(zip(column_refs, columns)):
                        output_lines.append(f"     {col_ref}: {col_name}")
                
                workbook.close()
                return [types.TextContent(
                    type="text",
                    text="\n".join(output_lines)
                )]
                
            except Exception as e:
                return [types.TextContent(
                    type="text",
                    text=f"Failed to read Excel file: {str(e)}"
                )]
        
        elif tool_name == "read_excel_sheet":
            path = arguments["path"]
            sheet_name = arguments["sheet_name"]
            cell_range = arguments.get("cell_range")
            
            # If path doesn't exist, try to find it in allowed directories
            if not os.path.exists(path):
                found_path = find_file_in_allowed_dirs(path, allowed_directories)
                if found_path:
                    path = found_path
                    print(f"Found Excel file at: {path}", file=sys.stderr)
                else:
                    return [types.TextContent(
                        type="text",
                        text=f"Could not find Excel file matching '{path}' in allowed directories: {allowed_directories}"
                    )]
            
            # Verify file exists and is Excel
            if not os.path.exists(path):
                return [types.TextContent(
                    type="text",
                    text=f"File not found: {path}"
                )]
            
            file_extension = os.path.splitext(path)[1].lower()
            if file_extension != '.xlsx':
                return [types.TextContent(
                    type="text",
                    text=f"Not an Excel file: {path}"
                )]
            
            # Read the specific sheet
            try:
                workbook = openpyxl.load_workbook(path, data_only=True)
                
                # Validate sheet exists
                if sheet_name not in workbook.sheetnames:
                    workbook.close()
                    return [types.TextContent(
                        type="text",
                        text=f"Sheet '{sheet_name}' not found in workbook"
                    )]
                
                sheet = workbook[sheet_name]
                
                # Parse cell range if provided
                if cell_range:
                    try:
                        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
                    except ValueError:
                        workbook.close()
                        return [types.TextContent(
                            type="text",
                            text=f"Invalid cell range format: {cell_range}"
                        )]
                else:
                    # Use full sheet range
                    min_col, min_row, max_col, max_row = range_boundaries(sheet.calculate_dimension())
                
                # Get column headers (first row)
                columns = []
                for col in range(min_col, max_col + 1):
                    cell = sheet.cell(min_row, col)
                    header = cell.value if cell.value is not None else f"Column {get_column_letter(col)}"
                    columns.append(header)
                
                # Collect non-empty rows
                records = []
                for row in range(min_row + 1, max_row + 1):
                    values = []
                    row_has_data = False
                    
                    for col in range(min_col, max_col + 1):
                        cell = sheet.cell(row, col)
                        value = cell.value
                        
                        # Convert datetime objects to ISO format
                        if isinstance(value, datetime):
                            value = value.isoformat()
                        
                        values.append(value)
                        if value is not None and value != "":
                            row_has_data = True
                    
                    if row_has_data:
                        records.append({
                            "row": row,
                            "values": values,
                            "cell_refs": [f"{get_column_letter(col)}{row}" for col in range(min_col, max_col + 1)]
                        })
                
                # Count charts and images
                chart_count = len([drawing for drawing in sheet._charts])
                image_count = len([drawing for drawing in sheet._images])
                
                # Prepare sheet data
                sheet_data = {
                    "sheet_name": sheet_name,
                    "dimensions": cell_range or sheet.calculate_dimension(),
                    "non_empty_cells": len(records),
                    "charts": chart_count,
                    "images": image_count,
                    "columns": columns,
                    "column_refs": [get_column_letter(col) for col in range(min_col, max_col + 1)],
                    "records": records,
                    "metadata": {
                        "total_rows": max_row - min_row + 1,
                        "total_columns": max_col - min_col + 1,
                        "visible_rows": len(records),
                        "has_charts": chart_count > 0,
                        "has_images": image_count > 0
                    }
                }
                
                workbook.close()
                return [types.TextContent(
                    type="text",
                    text=json.dumps(sheet_data, indent=2)
                )]
                
            except Exception as e:
                if 'workbook' in locals():
                    workbook.close()
                return [types.TextContent(
                    type="text",
                    text=f"Failed to read Excel sheet: {str(e)}"
                )]
        
        elif tool_name == "list_directory":
            path = arguments["path"]
            
            # If path doesn't exist, try to find it in allowed directories
            if not os.path.exists(path):
                try:
                    # Try to find a directory matching the name
                    for base_dir in allowed_directories:
                        for root, dirs, _ in os.walk(base_dir):
                            if path.lower() in [d.lower() for d in dirs]:
                                path = os.path.join(root, next(d for d in dirs if path.lower() in d.lower()))
                                print(f"Found directory at: {path}", file=sys.stderr)
                                break
                        if os.path.exists(path):
                            break
                except Exception as e:
                    print(f"Error while searching for directory: {e}", file=sys.stderr)
            
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
            
            # If path doesn't exist, try to find it in allowed directories
            if not os.path.exists(path):
                found_path = find_file_in_allowed_dirs(path, allowed_directories)
                if found_path:
                    path = found_path
                    print(f"Found file at: {path}", file=sys.stderr)
                else:
                    return [types.TextContent(
                        type="text",
                        text=f"Could not find file matching '{path}' in allowed directories: {allowed_directories}"
                    )]
            
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
            
            # Get file info as JSON
            file_info = get_file_info(path)
            
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
        
    except Exception as e:
        return [types.TextContent(
            type="text",
            text=f"Error: {str(e)}"
        )]

def get_file_info(path: str) -> dict:
    """Get detailed file information and capabilities."""
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
    
    return file_info

def get_directory_listing(path: str) -> dict:
    """Get directory contents with metadata."""
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
    
    return {
        "path": path,
        "contents": contents
    }

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