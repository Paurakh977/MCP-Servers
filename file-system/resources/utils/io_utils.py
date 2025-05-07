"""File I/O utility functions."""
import os

def save_to_file(content: str, original_path: str) -> str:
    """Save content to a text file and return the path."""
    # Create output filename based on original
    base_name = os.path.splitext(os.path.basename(original_path))[0]
    output_path = f"{base_name}_extracted.txt"
    
    # Save content to file
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return output_path
    except Exception as e:
        print(f"Error saving to file: {str(e)}")
        return None 