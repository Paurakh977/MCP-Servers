"""Format readers module for RTF and other specialized formats."""
import os

# For RTF documents
try:
    import striprtf.striprtf as striprtf
    has_rtf_support = True
except ImportError:
    has_rtf_support = False
    print("RTF support not available. To read RTF files: pip install striprtf")

def read_rtf_file(file_path: str) -> str:
    """Extract text content from RTF files."""
    if not has_rtf_support:
        return "RTF support not available. Install required package with: pip install striprtf"
    
    try:
        # Read RTF content
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            rtf_text = file.read()
        
        # Convert RTF to plain text
        plain_text = striprtf.rtf_to_text(rtf_text)
        
        # Format output
        content = []
        content.append(f"--- RTF Document: {os.path.basename(file_path)} ---")
        content.append(f"Size: {os.path.getsize(file_path) / 1024:.2f} KB")
        content.append("-" * 40)
        content.append(plain_text)
        
        return "\n".join(content)
    except Exception as e:
        return f"Error reading RTF file: {str(e)}" 