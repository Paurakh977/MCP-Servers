# MCP File System Server

A Model Context Protocol (MCP) server that provides file system resources and tools for reading different file formats with enhanced capabilities.

## Features

- Access to file system resources through MCP
- Rich text extraction from various file formats:
  - PDF documents (with image detection and layout preservation)
  - Microsoft Office files (Word, Excel, PowerPoint)
  - CSV and data files
  - EPUB ebooks
  - RTF documents
  - Plain text files
- File content summarization
- Directory listing

## Installation

1. Install the required dependencies:

```bash
pip install "mcp[cli]" PyPDF2
```

2. For enhanced functionality, install optional dependencies:

```bash
# For enhanced PDF support
pip install pymupdf pdfplumber tabula-py

# For Office files support
pip install python-docx openpyxl python-pptx Pillow

# For ebook support
pip install ebooklib beautifulsoup4

# For RTF support
pip install striprtf
```

## Usage

### Running the MCP Server

```bash
python main.py
```

This will start the MCP server on the default port.

### Integrating with Claude Desktop

To install this server in Claude Desktop:

```bash
mcp install main.py
```

### MCP Resources

The server exposes the following resources:

- `file://{file_path}` - Get the full content of a file
- `file-summary://{file_path}` - Get a summarized version of a file's content
- `directory://{dir_path}` - List all files in a directory

### MCP Tools

The server provides the following tools:

- `read_file(file_path, summarize, max_length)` - Read a file with optional summarization
- `save_file_content(content, file_path)` - Save text content to a file
- `search_files(directory, pattern, recursive)` - Search for files matching a pattern

## Example Queries for Claude

Once your MCP server is connected to Claude, you can use queries like:

- "Show me the content of my report.pdf file"
- "Summarize the Excel spreadsheet data.xlsx"
- "What files are in my documents folder?"
- "Search for all Python files in the project"

## Module Structure

```
resources/
├── __init__.py          # Main package exports
├── cli.py               # Command-line interface
├── extractor.py         # Main extraction API
├── readers/
│   ├── __init__.py      # Reader exports
│   ├── data_readers.py  # CSV file readers
│   ├── ebook_readers.py # EPUB file readers
│   ├── format_readers.py # RTF file readers
│   ├── office_readers.py # Word, Excel, PowerPoint readers
│   ├── pdf_reader.py    # PDF document readers
│   └── text_reader.py   # Text file readers
└── utils/
    ├── __init__.py      # Utilities exports
    ├── formatters.py    # Output formatting utils
    └── io_utils.py      # File I/O operations
```

## Standalone File Reading

The package can also be used standalone without MCP:

```python
from resources import read_file, print_output

# Read a file
result = read_file("document.pdf")

# Print the result
print_output(result)
```

## Command-Line Usage

The module also provides a command-line interface:

```bash
python -m resources.cli document.pdf --summarize --max-length 1000
```

Or use interactive mode:

```bash
python -m resources.cli
```
