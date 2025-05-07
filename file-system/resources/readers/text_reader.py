"""Text file reader module."""

def read_text_file(file_path: str) -> str:
    """Read content from text-based files with proper encoding detection."""
    encodings = ['utf-8', 'latin-1', 'windows-1252', 'ascii']
    
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                return f.read()
        except UnicodeDecodeError:
            continue

    # If all encodings fail, try binary mode as last resort
    try:
        with open(file_path, 'rb') as f:
            return f.read().decode('utf-8', errors='replace')
    except Exception as e:
        return f"Error: Could not decode file with available encodings: {str(e)}" 