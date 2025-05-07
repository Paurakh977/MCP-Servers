"""File content extraction module."""
import os
from typing import Dict, Any

from .readers import (
    read_text_file, read_pdf_file, read_docx_file, read_xlsx_file,
    read_pptx_file, read_csv_file, read_epub_file, read_rtf_file
)
from .utils.formatters import summarize_content

def extract_file_content(file_path: str) -> Dict[str, Any]:
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

def read_file(path: str, summarize: bool = False, max_summary_length: int = 500) -> Dict[str, Any]:
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
    result = extract_file_content(path)
    
    # Summarize very large content if requested
    if summarize and result["success"] and len(result["content"]) > max_summary_length:
        result["content"] = summarize_content(result["content"], max_summary_length)
        result["summarized"] = True
    
    return result 