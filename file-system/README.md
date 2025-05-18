# File System MCP Server

A powerful MCP (Managed Control Protocol) server for file system operations with advanced capabilities for reading and extracting content from various file types.

## Features

- Read content from various file types with advanced extraction:
  - PDF documents (including tables, images, and metadata)
  - Office documents (Word, Excel, PowerPoint)
  - Text files, CSV, EPUB, RTF, and more
- Extract metadata and capabilities information for files
- List directory contents with detailed information
- Secure path validation and access control

## Installation

### Prerequisites

- Python 3.7+
- For PDF table extraction: Java Runtime Environment (JRE) is required for tabula-py

### Install dependencies (Recommended method using uv)

```bash
# Clone the repository
git clone https://github.com/Paurakh977/MCP-Servers.git
cd file-system

# Create a virtual environment
uv venv

# Activate the virtual environment
# On Linux/macOS:
source .venv/bin/activate
# On Windows:
.venv\Scripts\activate

# Install dependencies using uv sync (requirements.txt already provided)
uv sync
```

### Alternative installation using pip

```bash
# Clone the repository
git clone https://github.com/Paurakh977/MCP-Servers.git
cd file-system

# Create a virtual environment
python -m venv .venv

# Activate the virtual environment
# On Linux/macOS:
source .venv/bin/activate
# On Windows:
.venv\Scripts\activate

# Install dependencies
pip install .

```

## Running the Server

To run the server, you need to specify one or more allowed directories that the server can access:

```bash
python server.py /path/to/allowed/directory1 /path/to/allowed/directory2
```

The server will only allow access to files within these specified directories for security reasons.

## Integrating with Claude Desktop

To configure Claude desktop to use this MCP server, add the following to your Claude config.json file:

### Config File Location
- Windows: `%APPDATA%\Claude\config.json` (typically `C:\Users\username\AppData\Roaming\Claude\config.json`)
- macOS: `~/Library/Application Support/Claude/config.json`
- Linux: `~/.config/Claude/config.json`

If the file doesn't exist, create it with the following content:

```json
{
  "mcpServers": {
    "file-system": {
      "command": "uv",
      "args": [
        "run",
        "--with",
        "mcp[cli]",
        "--with",
        "PyPDF2",
        "--with",
        "python-docx",
        "--with",
        "openpyxl",
        "--with",
        "python-pptx",
        "--with",
        "pillow",
        "--with",
        "pymupdf",
        "--with",
        "pdfplumber",
        "--with",
        "tabula-py",
        "--with",
        "ebooklib",
        "--with",
        "beautifulsoup4",
        "--with",
        "striprtf",
        "PATH_TO_SERVER_PY",
        "ALLOWED_DIR_1",
        "ALLOWED_DIR_2",
        "ALLOWED_DIR_3"
      ]
    }
  }
}
```

Replace:
- `PATH_TO_SERVER_PY` with the absolute path to the server.py file
- `ALLOWED_DIR_1`, `ALLOWED_DIR_2`, etc. with the absolute paths to directories you want to give Claude access to

For example:

```json
{
  "mcpServers": {
    "file-system": {
      "command": "uv",
      "args": [
        "run",
        "--with",
        "mcp[cli]",
        "--with",
        "PyPDF2",
        "--with",
        "python-docx",
        "--with",
        "openpyxl",
        "--with",
        "python-pptx",
        "--with",
        "pillow",
        "--with",
        "pymupdf",
        "--with",
        "pdfplumber",
        "--with",
        "tabula-py",
        "--with",
        "ebooklib",
        "--with",
        "beautifulsoup4",
        "--with",
        "striprtf",
        "C:\\Users\\username\\file-system\\server.py",
        "C:\\Users\\username\\Documents",
        "C:\\Users\\username\\Downloads"
      ]
    }
  }
}
```

After updating the config file, restart Claude Desktop for the changes to take effect.

## Usage in Claude

Once the server is configured, Claude can access your files and directories. You can ask Claude to:

1. Read file content:
   - "Read the PDF document at C:\\Users\\username\\Documents\\report.pdf"
   - "Summarize the content of my Word document at C:\\Users\\username\\Documents\\document.docx"

2. List directory contents:
   - "List all files in my Downloads folder"
   - "Show me what documents I have in C:\\Users\\username\\Documents"

3. Get file information:
   - "What information can you extract from this Excel file: C:\\Users\\username\\Documents\\spreadsheet.xlsx?"
   - "Tell me about this PDF file at C:\\Users\\username\\Documents\\document.pdf"

## Supported File Types

- Text files (.txt, .md, .json, .html, .xml, .log, .py, .js, .css, .java, .ini, .conf, .cfg)
- PDF documents (.pdf)
- Microsoft Word documents (.docx)
- Microsoft Excel spreadsheets (.xlsx)
- Microsoft PowerPoint presentations (.pptx)
- CSV files (.csv)
- EPUB e-books (.epub)
- Rich Text Format (.rtf)

## Security

- The server implements strict path validation to prevent access outside allowed directories
- Access is limited to directories specified when starting the server
- The MCP server ensures all paths are properly normalized and verified before access

## Limitations

- Some extraction features require additional libraries:
  - PDF image detection requires PyMuPDF
  - PDF table extraction requires tabula-py or pdfplumber
  - EPUB support requires ebooklib and beautifulsoup4
  - RTF support requires striprtf
- File types not explicitly supported will be treated as plain text files when possible

```json
resources/
├── core/                     # Core functionality and utilities
│   ├── __init__.py          # Exports core features and flags
│   ├── dependencies.py      # Manages all external package dependencies
│   ├── extraction_options.py # File extraction configuration options
│   └── file_utils.py        # Security and file handling utilities
│
├── readers/                  # File type-specific readers
│   ├── __init__.py          # Exports all reader functions
│   ├── pdf_reader.py        # PDF file extraction with layout preservation
│   ├── office_readers.py    # Word, Excel, PowerPoint file readers
│   ├── format_readers.py    # Various format readers
│   ├── ebook_readers.py     # EPUB and ebook format readers
│   ├── data_readers.py      # CSV and data file readers
│   └── text_reader.py       # Plain text file reader
│
├── utils/                    # Utility functions
│   └── formatters.py        # Text formatting and summarization
│
├── __init__.py              # Main package exports
└── extractor.py             # High-level file extraction interface
```

test 
"Create formulas to calculate the average, maximum, and minimum scores in the student sheet"
"Apply conditional formatting to highlight students with scores above 90"
"Create a formula that concatenates student names with their grades and sections"
"Calculate the percentage of students in each grade level"
"Apply VLOOKUP to connect student and teacher data based on subject names"