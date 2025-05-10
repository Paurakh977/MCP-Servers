"""File content extraction module."""
import os
import json
from typing import Dict, Any

from .readers import (
    read_text_file, read_pdf_file, read_docx_file, read_xlsx_file,
    read_pptx_file, read_csv_file, read_epub_file, read_rtf_file
)
from .utils.formatters import summarize_content

def extract_file_content(file_path: str, sheet_name: str = None, cell_range: str = None) -> Dict[str, Any]:
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
            content = read_pdf_file(file_path)
        elif file_extension == '.docx':
            content = read_docx_file(file_path)
        elif file_extension == '.xlsx':
            content = read_xlsx_file(file_path, sheet_name=sheet_name, cell_range=cell_range)
            # Parse JSON string back to dict for consistent return
            try:
                content = json.loads(content)
            except json.JSONDecodeError:
                return {"success": False, "error": "Failed to parse Excel file output", "content": content}
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

def read_file(path: str, summarize: bool = False, max_summary_length: int = 500,
              sheet_name: str = None, cell_range: str = None) -> Dict[str, Any]:
    """
    Reads and returns the contents of the file at 'path' with appropriate handling per file type.
    
    Args:
        path (str): Path to the file to read
        summarize (bool): Whether to summarize very large content
        max_summary_length (int): Maximum length for summary if summarizing
        sheet_name (str, optional): For Excel files, specific sheet to read
        cell_range (str, optional): For Excel files, cell range to read (e.g. 'A1:D10')
        
    Returns:
        Dict with keys:
        - success (bool): Whether the read was successful
        - content (str/dict): The file content, possibly summarized. For Excel files, this is a dict
        - file_path (str): Original file path
        - file_type (str): File extension
        - error (str, optional): Error message if success is False
    """
    result = extract_file_content(path, sheet_name=sheet_name, cell_range=cell_range)
    
    # Handle summarization for text content only (not for JSON/dict content from Excel)
    if summarize and result["success"] and isinstance(result["content"], str) and len(result["content"]) > max_summary_length:
        result["content"] = summarize_content(result["content"], max_summary_length)
        result["summarized"] = True
    
    return result 