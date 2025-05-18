"""Resource definitions for the MCP file-system server."""

import mcp.types as types

def get_resource_definitions() -> list[types.Resource]:
    """
    Get a list of all resource definitions for the MCP server.
    
    Returns:
        List of Resource objects that define the capabilities of the server.
    """
    return [
        types.Resource(
            uri="file:///{path}",
            name="File Content",
            description="Read content from various file types with advanced extraction capabilities. PDF documents include tables, images, and metadata. Office documents include tables, charts, and embedded objects. All file types support optional summarization.",
            schema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string", 
                        "description": "Path to the file relative to allowed directories"
                    },
                    "summarize": {
                        "type": "boolean",
                        "description": "Whether to summarize large content",
                        "default": False
                    },
                    "max_length": {
                        "type": "integer",
                        "description": "Maximum length for summary",
                        "default": 500
                    },
                    "metadata_only": {
                        "type": "boolean",
                        "description": "Extract only metadata (for PDFs, Office docs, EPUB)",
                        "default": False
                    }
                },
                "required": ["path"]
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
        types.Resource(
            uri="excel-info:///{path}",
            name="Excel Workbook Information",
            description="Get metadata and sheet information from Excel files",
            mimeType="application/json",
            schema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel file"
                    }
                },
                "required": ["path"]
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
        types.Resource(
            uri="excel-sheet:///{path}",
            name="Excel Sheet Content",
            description="Read content from a specific sheet in an Excel file",
            mimeType="application/json",
            schema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel file"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the sheet to read"
                    },
                    "cell_range": {
                        "type": "string",
                        "description": "Optional cell range to read (e.g. 'A1:D10')",
                        "default": None
                    }
                },
                "required": ["path", "sheet_name"]
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
        types.Resource(
            uri="file-info:///{path}",
            name="File Information",
            description="Get metadata and extraction capabilities for a file",
            mimeType="application/json",
            schema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the file relative to current directory"
                    }
                },
                "required": ["path"]
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
        types.Resource(
            uri="directory:///{path}",
            name="Directory Listing",
            description="List files and subdirectories in a directory",
            mimeType="application/json",
            schema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the directory relative to current directory"
                    }
                },
                "required": ["path"]
            },
            idempotentHint=True,
            readOnlyHint=True,
        ),
    ] 