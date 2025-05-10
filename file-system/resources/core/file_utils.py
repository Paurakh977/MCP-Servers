"""Module for file-related utility functions."""

import os
import sys
import mimetypes
from typing import Dict, Any, List

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