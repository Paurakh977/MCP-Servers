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
        ),
        # Excel Workbook Operations
        types.Tool(
            name="create_excel_workbook",
            description="Create a new Excel workbook. You can provide an absolute path or relative path within an allowed directory.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path where to create the Excel workbook. Can be absolute or relative to an allowed directory."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name for the initial worksheet (default: 'Sheet1')",
                        "default": "Sheet1"
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        types.Tool(
            name="get_workbook_metadata",
            description="Get detailed metadata about an Excel workbook including sheets, ranges and dimensions.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "include_ranges": {
                        "type": "boolean",
                        "description": "Whether to include detailed range information for each sheet",
                        "default": False
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        ),
        # Excel Worksheet Operations
        types.Tool(
            name="create_worksheet",
            description="Create a new worksheet in an existing Excel workbook.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name for the new worksheet"
                    }
                },
                "required": ["path", "sheet_name"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        types.Tool(
            name="copy_worksheet",
            description="Copy a worksheet within the same Excel workbook.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "source_sheet": {
                        "type": "string",
                        "description": "Name of the sheet to copy"
                    },
                    "target_sheet": {
                        "type": "string",
                        "description": "Name for the new copied sheet"
                    }
                },
                "required": ["path", "source_sheet", "target_sheet"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        types.Tool(
            name="delete_worksheet",
            description="Delete a worksheet from an Excel workbook.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the sheet to delete"
                    }
                },
                "required": ["path", "sheet_name"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        types.Tool(
            name="rename_worksheet",
            description="Rename a worksheet in an Excel workbook.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "old_name": {
                        "type": "string",
                        "description": "Current name of the sheet"
                    },
                    "new_name": {
                        "type": "string",
                        "description": "New name for the sheet"
                    }
                },
                "required": ["path", "old_name", "new_name"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        # Excel Range Operations
        types.Tool(
            name="copy_excel_range",
            description="Copy a range of cells to another location in Excel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Source worksheet name"
                    },
                    "source_range": {
                        "type": "string",
                        "description": "Source range to copy (e.g. 'A1:D10' or just 'A1' for a single cell)"
                    },
                    "target_start": {
                        "type": "string",
                        "description": "Starting cell for paste location (e.g. 'E1')"
                    },
                    "target_sheet": {
                        "type": "string",
                        "description": "Optional target worksheet name if different from source",
                        "default": None
                    }
                },
                "required": ["path", "sheet_name", "source_range", "target_start"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        types.Tool(
            name="delete_excel_range",
            description="Delete a range of cells and shift remaining cells in Excel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Worksheet name"
                    },
                    "start_cell": {
                        "type": "string",
                        "description": "Starting cell of range (e.g. 'A1')"
                    },
                    "end_cell": {
                        "type": "string",
                        "description": "Ending cell of range (e.g. 'D10')",
                        "default": None
                    },
                    "shift_direction": {
                        "type": "string",
                        "description": "Direction to shift cells ('up' or 'left')",
                        "default": "up"
                    }
                },
                "required": ["path", "sheet_name", "start_cell"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        types.Tool(
            name="merge_excel_cells",
            description="Merge a range of cells in Excel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Worksheet name"
                    },
                    "start_cell": {
                        "type": "string",
                        "description": "Starting cell of range (e.g. 'A1')"
                    },
                    "end_cell": {
                        "type": "string",
                        "description": "Ending cell of range (e.g. 'D10')"
                    }
                },
                "required": ["path", "sheet_name", "start_cell", "end_cell"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        types.Tool(
            name="unmerge_excel_cells",
            description="Unmerge a previously merged range of cells in Excel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Worksheet name"
                    },
                    "start_cell": {
                        "type": "string",
                        "description": "Starting cell of merged range (e.g. 'A1')"
                    },
                    "end_cell": {
                        "type": "string",
                        "description": "Ending cell of merged range (e.g. 'D10')"
                    }
                },
                "required": ["path", "sheet_name", "start_cell", "end_cell"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        # Excel Data Operations
        types.Tool(
            name="write_excel_data",
            description="Write data to an Excel worksheet.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Worksheet name"
                    },
                    "data": {
                        "type": "array",
                        "description": "Array of arrays (rows and columns) containing data to write",
                        "items": {
                            "type": "array",
                            "items": {
                                "type": ["string", "number", "boolean", "null"]
                            }
                        }
                    },
                    "start_cell": {
                        "type": "string",
                        "description": "Starting cell for data insertion (e.g. 'A1')",
                        "default": "A1"
                    },
                    "headers": {
                        "type": "boolean",
                        "description": "Whether the first row of data contains headers",
                        "default": True
                    },
                    "auto_adjust_width": {
                        "type": "boolean",
                        "description": "Whether to automatically adjust column widths based on content",
                        "default": False
                    }
                },
                "required": ["path", "sheet_name", "data"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        types.Tool(
            name="format_excel_range",
            description="Apply formatting to a range of cells in Excel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Worksheet name"
                    },
                    "start_cell": {
                        "type": "string",
                        "description": "Starting cell of range (e.g. 'A1')"
                    },
                    "end_cell": {
                        "type": "string",
                        "description": "Ending cell of range (e.g. 'D10')",
                        "default": None
                    },
                    "bold": {
                        "type": "boolean",
                        "description": "Apply bold formatting",
                        "default": False
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "Apply italic formatting",
                        "default": False
                    },
                    "font_size": {
                        "type": "integer",
                        "description": "Set font size",
                        "default": None
                    },
                    "font_color": {
                        "type": "string",
                        "description": "Set font color (hex code e.g. 'FF0000' for red)",
                        "default": None
                    },
                    "bg_color": {
                        "type": "string",
                        "description": "Set background color (hex code e.g. 'FFFF00' for yellow)",
                        "default": None
                    },
                    "alignment": {
                        "type": "string",
                        "description": "Set text alignment ('left', 'center', 'right')",
                        "default": None
                    },
                    "wrap_text": {
                        "type": "boolean",
                        "description": "Enable text wrapping",
                        "default": False
                    },
                    "border_style": {
                        "type": "string",
                        "description": "Add borders ('thin', 'medium', 'thick', 'dashed', 'dotted', 'double')",
                        "default": None
                    },
                    "auto_adjust_width": {
                        "type": "boolean",
                        "description": "Automatically adjust column widths based on content",
                        "default": False
                    }
                },
                "required": ["path", "sheet_name", "start_cell"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        types.Tool(
            name="adjust_column_widths",
            description="Adjust column widths in an Excel worksheet based on content or custom specifications.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Worksheet name"
                    },
                    "column_range": {
                        "type": "string",
                        "description": "Range of columns to adjust (e.g. 'A:D' or just 'A' for a single column)",
                        "default": None
                    },
                    "auto_fit": {
                        "type": "boolean",
                        "description": "Whether to automatically fit column widths to content",
                        "default": True
                    },
                    "custom_widths": {
                        "type": "object", 
                        "description": "Optional dictionary mapping column letters to widths (e.g. {'A': 15, 'B': 20})",
                        "default": None,
                        "additionalProperties": {
                            "type": "integer"
                        }
                    }
                },
                "required": ["path", "sheet_name"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        )
    ] 
