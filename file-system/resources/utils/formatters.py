"""Output formatting utilities."""
import json
from typing import Dict, Any

def summarize_content(content: str, max_length: int = 500) -> str:
    """Create a brief summary of the content if it's too long."""
    if len(content) <= max_length:
        return content
    
    # Simple summarization: first few characters and last few characters
    first_part = content[:max_length // 2]
    last_part = content[-(max_length // 2):]
    return f"{first_part}\n\n... [Content truncated, total length: {len(content)} characters] ...\n\n{last_part}"

def print_output(result: Dict[str, Any], output_format: str = "text") -> None:
    """Print the output in the specified format."""
    if output_format == "json":
        print(json.dumps(result, indent=2))
    else:
        if result["success"]:
            print(f"File: {result['file_path']} ({result['file_type']})")
            print("-" * 40)
            print(result["content"])
        else:
            print(f"Error: {result['error']}") 