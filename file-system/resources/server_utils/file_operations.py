"""File system operations for the MCP server."""

import os
import time
import mimetypes
from typing import Optional, Dict, Any, List

# Import feature flags from core
from resources.core import (
    has_pymupdf, has_pdfplumber, has_tabula,
    has_pil, has_epub_support, has_rtf_support,
    get_mime_type, get_file_extraction_options
)

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