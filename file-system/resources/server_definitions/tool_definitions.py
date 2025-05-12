"""Tool definitions for the MCP file-system server."""

import mcp.types as types

def get_tool_definitions() -> list[types.Tool]:
    """
    Get a list of all tool definitions for the MCP server.
    
    Returns:
        List of Tool objects that define the capabilities of the server.
    """
    return [
        types.Tool(
            name="read_file",
            description="Read a file with advanced extraction capabilities. You can provide either an absolute path, "
                       "relative path, or just the filename - the server will search for matching files in allowed directories. "
                       "For example: 'report.pdf', 'downloads/data.xlsx', or just 'summary' to find files containing that name.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path or name of the file to read. Can be absolute path, relative path, or filename to search for."
                    },
                    "summarize": {
                        "type": "boolean",
                        "description": "Whether to summarize large content",
                        "default": False
                    },
                    "max_length": {
                        "type": "integer",
                        "description": "Maximum length for summary if summarizing",
                        "default": 500
                    },
                    "metadata_only": {
                        "type": "boolean",
                        "description": "Extract only metadata (for PDFs, Office docs)",
                        "default": False
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        ),
        types.Tool(
            name="get_excel_info",
            description="Get metadata and sheet information from an Excel file. You can provide either an absolute path, "
                       "relative path, or just the filename - the server will search for matching Excel files in allowed directories. "
                       "For example: 'data.xlsx', 'reports/summary.xlsx', or just 'sales' to find Excel files containing that name.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path or name of the Excel file. Can be absolute path, relative path, or filename to search for."
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        ),
        types.Tool(
            name="read_excel_sheet",
            description="Read content from a specific sheet in an Excel file. You can provide either an absolute path, "
                       "relative path, or just the filename - the server will search for matching Excel files in allowed directories. "
                       "For example: 'data.xlsx', 'reports/summary.xlsx', or just 'sales' to find Excel files containing that name.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path or name of the Excel file. Can be absolute path, relative path, or filename to search for."
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
                "required": ["path", "sheet_name"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        ),
        types.Tool(
            name="list_directory",
            description="List files and folders in a directory. You can provide either an absolute path, relative path, "
                       "or just the directory name - the server will search for matching directories in allowed locations. "
                       "For example: 'downloads', 'documents/reports', or just 'data' to find directories containing that name.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path or name of the directory. Can be absolute path, relative path, or directory name to search for."
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        ),
        types.Tool(
            name="get_file_info",
            description="Get detailed information about a file including supported extraction capabilities. You can provide either "
                       "an absolute path, relative path, or just the filename - the server will search for matching files in allowed directories. "
                       "For example: 'document.pdf', 'reports/data.xlsx', or just 'summary' to find files containing that name.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path or name of the file. Can be absolute path, relative path, or filename to search for."
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        ),
        types.Tool(
            name="list_allowed_directories",
            description="List directories that this server is allowed to access. Use this to understand where the server can search for files.",
            inputSchema={
                "type": "object",
                "properties": {},
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        )
    ] 