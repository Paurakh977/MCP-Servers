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
        ),
        types.Tool(
            name="apply_excel_formula",
            description="Apply any Excel formula to a specific cell with robust error handling and support for advanced features. Handles modern array formulas (XLOOKUP, FILTER, UNIQUE), cross-worksheet references, external workbook references, and protects against common errors like division by zero. Automatically performs syntax validation, ensures proper formatting, and provides detailed feedback.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the worksheet where the formula should be applied."
                    },
                    "cell": {
                        "type": "string",
                        "description": "Cell reference where formula should be applied (e.g., 'A1', 'B5')."
                    },
                    "formula": {
                        "type": "string", 
                        "description": "Excel formula to apply (with or without leading '='). Can include references to other sheets, workbooks, and use any Excel function."
                    },
                    "protect_from_errors": {
                        "type": "boolean",
                        "description": "Whether to automatically protect against errors by wrapping risky formulas in IFERROR(). Defaults to true.",
                        "default": True
                    },
                    "handle_arrays": {
                        "type": "boolean",
                        "description": "Whether to properly handle modern array formulas (XLOOKUP, FILTER, etc.)",
                        "default": True
                    },
                    "clear_spill_range": {
                        "type": "boolean",
                        "description": "Whether to automatically clear cells below/right to ensure array formulas can properly spill their results",
                        "default": True
                    },
                    "spill_rows": {
                        "type": "integer",
                        "description": "For array formulas like UNIQUE, how many rows to clear below for potential results. Default is 200, increase for larger datasets.",
                        "default": 200
                    }
                },
                "required": ["path", "sheet_name", "cell", "formula"]
            }
        ),
        types.Tool(
            name="apply_excel_formula_range",
            description="Apply formulas to a range of cells with powerful templating, performance optimization for large datasets, and support for modern array functions. Use {row} and {col} placeholders in formulas that will be replaced with the current row number and column letter. Handles chunked processing for large ranges, modern array formulas, and provides detailed feedback on any issues. Ideal for calculating multiple cells with a pattern.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the worksheet where the formula should be applied."
                    },
                    "start_cell": {
                        "type": "string",
                        "description": "Top-left cell of the range (e.g., 'A1', 'B5')."
                    },
                    "end_cell": {
                        "type": "string",
                        "description": "Bottom-right cell of the range (e.g., 'A10', 'D20')."
                    },
                    "formula_template": {
                        "type": "string",
                        "description": "Formula template with {row} and/or {col} placeholders that will be replaced with the current row number and column letter. Example: '=SUM(A{row}:E{row})' or '=AVERAGE({col}1:{col}10)'."
                    },
                    "protect_from_errors": {
                        "type": "boolean",
                        "description": "Whether to automatically protect against errors by wrapping risky formulas in IFERROR(). Defaults to true.",
                        "default": True
                    },
                    "dynamic_calculation": {
                        "type": "boolean",
                        "description": "Whether to use dynamic array functions for better performance",
                        "default": True
                    },
                    "clear_spill_range": {
                        "type": "boolean",
                        "description": "Whether to automatically clear cells below/right to ensure array formulas can properly spill their results",
                        "default": True
                    },
                    "chunk_size": {
                        "type": "integer",
                        "description": "For large ranges, how many cells to process at once before saving. Helps with memory usage for large workbooks. Defaults to 1000.",
                        "default": 1000
                    }
                },
                "required": ["path", "sheet_name", "start_cell", "end_cell", "formula_template"]
            }
        ),
        types.Tool(
            name="delete_excel_workbook",
            description="Delete an entire Excel workbook file from the file system. This permanently removes the file and cannot be undone.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook to delete. Can be absolute path, relative path, or filename to search for."
                    }
                },
                "required": ["path"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        types.Tool(
            name="add_excel_column",
            description="Add a new column to an Excel worksheet without rewriting the entire sheet. This is the preferred way to add a column to existing data, as it preserves all existing content and formatting. You can specify where to insert the column and add data values for the new column.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the worksheet to modify"
                    },
                    "column_name": {
                        "type": "string",
                        "description": "Name for the new column (header text)"
                    },
                    "column_position": {
                        "type": "string",
                        "description": "Optional column letter where to insert (e.g., 'C' to insert as third column). If not provided, adds to the end of existing data.",
                        "default": None
                    },
                    "data": {
                        "type": "array",
                        "description": "Optional list of values for the column. Length should match the data rows in the sheet.",
                        "items": {
                            "type": ["string", "number", "boolean", "null"]
                        },
                        "default": None
                    },
                    "header_style": {
                        "type": "object",
                        "description": "Optional styling for the header cell (e.g., {'bold': true, 'bg_color': 'FFFF00', 'font_size': 12, 'alignment': 'center'})",
                        "default": None,
                        "properties": {
                            "bold": {"type": "boolean"},
                            "bg_color": {"type": "string"},
                            "font_size": {"type": "integer"},
                            "alignment": {"type": "string"}
                        }
                    }
                },
                "required": ["path", "sheet_name", "column_name"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        types.Tool(
            name="add_data_validation",
            description="Add data validation rules to Excel cells, supporting dropdown lists, number ranges, date validation, text length limits, and custom formulas. Creates interactive dropdown menus, limits input to valid values, and can display custom error messages.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the worksheet where validation should be applied."
                    },
                    "cell_range": {
                        "type": "string",
                        "description": "Range of cells to apply validation to (e.g., 'A1:A10', 'B2:D15')."
                    },
                    "validation_type": {
                        "type": "string",
                        "description": "Type of validation to apply: 'list' (dropdown), 'decimal', 'date', 'textLength', or 'custom'.",
                        "enum": ["list", "decimal", "date", "textLength", "custom"]
                    },
                    "validation_criteria": {
                        "type": "object",
                        "description": "Criteria for validation, depends on validation_type. For 'list': {\"source\": [\"Option1\", \"Option2\"]} or {\"source\": \"=Sheet2!A1:A10\"}. For 'decimal': {\"operator\": \"between\", \"minimum\": 1, \"maximum\": 100}. For 'custom': {\"formula\": \"=AND(A1>0,A1<100)\"}."
                    },
                    "error_message": {
                        "type": "string",
                        "description": "Optional custom error message to display when validation fails."
                    }
                },
                "required": ["path", "sheet_name", "cell_range", "validation_type", "validation_criteria"]
            }
        ),
        types.Tool(
            name="apply_conditional_formatting",
            description="Apply sophisticated conditional formatting to Excel cells based on simple or complex conditions. Supports compound conditions with AND/OR operators that can combine different condition types (numeric, text, blank) in a single rule. This powerful tool efficiently highlights important data patterns without needing to manually filter and format cells.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the worksheet where conditional formatting should be applied."
                    },
                    "cell_range": {
                        "type": "string",
                        "description": "Range of cells to check and potentially format (e.g., 'A1:D10', 'B2:B20')."
                    },
                    "condition": {
                        "type": "string",
                        "description": "Condition for formatting with full support for complex expressions:\n- Numeric: \">90\", \"<=50\", \"=100\", \"<>0\" (not equal to zero)\n- Text: \"='Yes'\", \"<>'No'\", \"CONTAINS('text')\", \"STARTS_WITH('A')\", \"REGEX('pattern')\"\n- Date: \">DATE(2023,1,1)\", \"<=TODAY()\"\n- Blank: \"=ISBLANK()\", \"<>ISBLANK()\", \"ISBLANK()\", \"NOTBLANK()\"\n- Column references: \">{F}\" (compare with column F), \"<=2*{C}\" (compare with twice value in column C)\n- Named columns: \">{profit}\" (when using compare_columns parameter)\n- Compound examples:\n  * \">50 AND <=70\" (values between 50-70)\n  * \"<20 OR >80\" (values below 20 or above 80)\n  * \"CONTAINS('Pass') OR =ISBLANK()\" (cells containing 'Pass' or empty cells)\n  * \"<>ISBLANK() AND STARTS_WITH('Q')\" (non-empty cells starting with 'Q')\n  * \">{F} AND <100\" (greater than value in column F but less than 100)"
                    },
                    "condition_column": {
                        "type": "string",
                        "description": "Optional column letter that should be evaluated for the condition (e.g., 'D' for Units column). When specified, only this column is checked against the condition.",
                        "default": None
                    },
                    "format_entire_row": {
                        "type": "boolean",
                        "description": "When true, formats the entire row if the condition is met. Useful for highlighting entire rows based on a value in a specific column.",
                        "default": False
                    },
                    "columns_to_format": {
                        "type": "array",
                        "description": "Optional list of specific column letters to format when condition is met (e.g., ['A', 'C', 'D']). Use this to format only specific columns in matching rows. Takes precedence over format_entire_row.",
                        "items": {
                            "type": "string"
                        },
                        "default": None
                    },
                    "handle_formulas": {
                        "type": "boolean",
                        "description": "Whether to evaluate formulas for condition checking. When true, formula cells will be evaluated using their calculated values.",
                        "default": True
                    },
                    "outside_range_columns": {
                        "type": "array",
                        "description": "Optional list of columns outside the main range to format when condition is met (e.g., ['G', 'H']). Useful for formatting columns that aren't part of the condition range.",
                        "items": {
                            "type": "string"
                        },
                        "default": None
                    },
                    "compare_columns": {
                        "type": "object",
                        "description": "Optional mapping of column names to actual column letters for use in conditions. Example: {\"profit\": \"F\", \"sales\": \"C\"} lets you use {profit} and {sales} in conditions.",
                        "default": None
                    },
                    "date_format": {
                        "type": "string",
                        "description": "Optional date format string for parsing date values in conditions (e.g., \"%Y-%m-%d\" for YYYY-MM-DD format).",
                        "default": None
                    },
                    "icon_set": {
                        "type": "string",
                        "description": "Optional icon set to apply ('3arrows', '3trafficlights', '3symbols', '3stars', etc.). Note: This is a simplified implementation and may have limitations.",
                        "default": None
                    },
                    "bold": {
                        "type": "boolean",
                        "description": "Whether to apply bold formatting to matching cells.",
                        "default": False
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "Whether to apply italic formatting to matching cells.",
                        "default": False
                    },
                    "font_size": {
                        "type": "integer",
                        "description": "Font size to apply to matching cells.",
                        "default": None
                    },
                    "font_color": {
                        "type": "string",
                        "description": "Font color (hex code e.g., 'FF0000' or '#FF0000' for red) for matching cells.",
                        "default": None
                    },
                    "bg_color": {
                        "type": "string",
                        "description": "Background color (hex code e.g., 'FFFF00' or '#FFFF00' for yellow) for matching cells.",
                        "default": None
                    },
                    "alignment": {
                        "type": "string",
                        "description": "Text alignment ('left', 'center', 'right') for matching cells.",
                        "default": None
                    },
                    "wrap_text": {
                        "type": "boolean",
                        "description": "Whether to enable text wrapping for matching cells.",
                        "default": False
                    },
                    "border_style": {
                        "type": "string",
                        "description": "Border style ('thin', 'medium', 'thick', 'dashed', 'dotted', 'double') for matching cells.",
                        "default": None
                    }
                },
                "required": ["path", "sheet_name", "cell_range", "condition"],
                "additionalProperties": False
            },
            idempotentHint=False,
            readOnlyHint=False
        ),
        # --- Excel Table Tools ---
        types.Tool(
            name="create_excel_table",
            description="Creates a formatted Excel table from a specified data range in a sheet. This structures the data for easier sorting, filtering, and use in PivotTables.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {
                        "type": "string",
                        "description": "Path to the Excel workbook."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Name of the worksheet containing the data."
                    },
                    "data_range": {
                        "type": "string",
                        "description": "The range of data to be converted into a table (e.g., 'A1:D100'). Headers should be included in this range."
                    },
                    "table_name": {
                        "type": "string",
                        "description": "A unique name for the new table."
                    },
                    "table_style": {
                        "type": "string",
                        "description": "Optional: The style to apply to the table (e.g., 'TableStyleMedium9', 'TableStyleLight1'). Defaults to 'TableStyleMedium9'.",
                        "default": "TableStyleMedium9"
                    }
                },
                "required": ["path", "sheet_name", "data_range", "table_name"],
                "additionalProperties": False
            }
        ),
        types.Tool(
            name="sort_excel_table",
            description="Sorts an existing Excel table by a specified column and order (ascending/descending).",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet containing the table."},
                    "table_name": {"type": "string", "description": "Name of the table to sort."},
                    "sort_column_name": {"type": "string", "description": "The header name of the column to sort by."},
                    "sort_order": {
                        "type": "string",
                        "description": "Sort order: 'ascending' or 'descending'.",
                        "enum": ["ascending", "descending"],
                        "default": "ascending"
                    }
                },
                "required": ["path", "sheet_name", "table_name", "sort_column_name"],
                "additionalProperties": False
            }
        ),
        types.Tool(
            name="filter_excel_table",
            description="Filters an Excel table based on criteria for a specific column. Supports various operators.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet containing the table."},
                    "table_name": {"type": "string", "description": "Name of the table to filter."},
                    "column_name": {"type": "string", "description": "The header name of the column to filter."},
                    "criteria1": {
                        "type": "string",
                        "description": "Primary filter criteria (e.g., a value, a condition like '>100')."
                    },
                    "operator": {
                        "type": "string",
                        "description": "Filter operator (e.g., 'equals', 'contains', 'beginswith', 'endswith', 'greaterthan', 'lessthan', 'between').",
                        "enum": ["equals", "contains", "beginswith", "endswith", "greaterthan", "lessthan", "between", "notequals", "doesnotcontain"],
                        "default": "equals"
                    },
                    "criteria2": {
                        "type": "string",
                        "description": "Secondary filter criteria, used for 'between' operator.",
                        "default": None
                    }
                },
                "required": ["path", "sheet_name", "table_name", "column_name", "criteria1"],
                "additionalProperties": False
            }
        ),
        # --- PivotTable Tools ---
        types.Tool(
            name="create_pivot_table",
            description="Creates a new PivotTable from a source data range, allowing specification of rows, columns, values (with summary functions like Sum, Count, Average), and filters. This is a powerful tool for summarizing and analyzing large datasets.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "source_sheet_name": {"type": "string", "description": "Name of the sheet containing the source data for the PivotTable."},
                    "source_data_range": {"type": "string", "description": "Range of the source data (e.g., 'A1:G100', or a table name)."},
                    "target_sheet_name": {"type": "string", "description": "Name of the sheet where the PivotTable will be created."},
                    "target_cell_address": {"type": "string", "description": "Top-left cell for the PivotTable (e.g., 'A3')."},
                    "pivot_table_name": {"type": "string", "description": "A unique name for the new PivotTable."},
                    "row_fields": {
                        "type": "array", "items": {"type": "string"},
                        "description": "List of field names (column headers from source data) to use as row fields.",
                        "default": None
                    },
                    "column_fields": {
                        "type": "array", "items": {"type": "string"},
                        "description": "List of field names to use as column fields.",
                        "default": None
                    },
                    "value_fields": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "field": {"type": "string", "description": "Source field name for the value calculation."},
                                "function": {
                                    "type": "string",
                                    "description": "Summary function (e.g., 'Sum', 'Count', 'Average', 'Max', 'Min', 'Product', 'CountNumbers', 'StdDev', 'StdDevp', 'Var', 'Varp'). Defaults to 'Sum'.",
                                    "default": "Sum",
                                    "enum": ["Sum", "Count", "Average", "Max", "Min", "Product", "CountNumbers", "StdDev", "StdDevp", "Var", "Varp"]
                                },
                                "custom_name": {"type": "string", "description": "Optional custom name for the value field in the PivotTable (e.g., 'Total Sales').", "default": None}
                            },
                            "required": ["field"]
                        },
                        "description": "List of fields to use for values, including their summary function and optional custom name.",
                        "default": None
                    },
                    "filter_fields": {
                        "type": "array", "items": {"type": "string"},
                        "description": "List of field names to use as report filters (page fields).",
                        "default": None
                    },
                    "pivot_style": {
                        "type": "string",
                        "description": "Style for the PivotTable (e.g., 'PivotStyleMedium9', 'PivotStyleLight16'). Defaults to 'PivotStyleMedium9'.",
                        "default": "PivotStyleMedium9"
                    }
                },
                "required": ["path", "source_sheet_name", "source_data_range", "target_sheet_name", "target_cell_address", "pivot_table_name"],
            }
        ),
        types.Tool(
            name="modify_pivot_table_fields",
            description="Modifies the field layout of an existing PivotTable. Allows adding or removing fields from row, column, value, and filter areas.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Sheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable to modify."},
                    "add_row_fields": {"type": "array", "items": {"type": "string"}, "default": None, "description": "List of field names to add to rows."},
                    "add_column_fields": {"type": "array", "items": {"type": "string"}, "default": None, "description": "List of field names to add to columns."},
                    "add_value_fields": {
                        "type": "array", "items": {
                            "type": "object",
                            "properties": {
                                "field": {"type": "string"}, "function": {"type": "string", "default": "Sum"}, "custom_name": {"type": "string", "default": None}
                            }, "required": ["field"]},
                        "default": None, "description": "List of value fields to add (see create_pivot_table for structure)."
                    },
                    "add_filter_fields": {"type": "array", "items": {"type": "string"}, "default": None, "description": "List of field names to add to filters."},
                    "remove_fields": {"type": "array", "items": {"type": "string"}, "default": None, "description": "List of field names to remove from any area of the PivotTable."}
                },
                "required": ["path", "sheet_name", "pivot_table_name"]
            }
        ),
        types.Tool(
            name="sort_pivot_table_field",
            description="Sorts items within a PivotTable field (rows or columns) based on their labels (A-Z) or by the values of a data field (e.g., sort Products by Sum of Sales).",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Sheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable."},
                    "field_name": {"type": "string", "description": "The name of the PivotField (in rows or columns) whose items are to be sorted."},
                    "sort_on_field": {"type": "string", "description": "The caption of the DataField (value field, e.g., 'Sum of Sales') to sort by, or the 'field_name' itself to sort by labels."},
                    "sort_order": {
                        "type": "string", "description": "Sort order: 'ascending' or 'descending'.",
                        "enum": ["ascending", "descending"], "default": "ascending"
                    },
                    "sort_type": {
                        "type": "string", "description": "Sort type: 'data' (sort by values in 'sort_on_field') or 'label' (sort 'field_name' items alphabetically).",
                        "enum": ["data", "label"], "default": "data"
                    }
                },
                "required": ["path", "sheet_name", "pivot_table_name", "field_name", "sort_on_field"]
            }
        ),
        types.Tool(
            name="filter_pivot_table_items",
            description="Applies filters to a PivotTable field by specifying which items should be visible or hidden. Useful for focusing on specific categories or values within the PivotTable.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Sheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable."},
                    "field_name": {"type": "string", "description": "The name of the PivotField to filter."},
                    "visible_items": {
                        "type": "array", "items": {"type": "string"}, "default": None,
                        "description": "A list of item names to make visible. If provided, all other items in this field will be hidden. Takes precedence over hidden_items if both are provided."
                    },
                    "hidden_items": {
                        "type": "array", "items": {"type": "string"}, "default": None,
                        "description": "A list of item names to hide. Applied if visible_items is not provided."
                    }
                },
                "required": ["path", "sheet_name", "pivot_table_name", "field_name"]
            }
        ),

        types.Tool(
            name="refresh_pivot_table",
            description="Refreshes a specified PivotTable to update its data from the source. This is essential if the underlying data for the PivotTable has changed.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Sheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable to refresh."}
                },
                "required": ["path", "sheet_name", "pivot_table_name"]
            }
        ),
        # --- BEGIN ADVANCED PIVOTTABLE TOOL DEFINITIONS ---
        types.Tool(
            name="add_pivot_table_calculated_field",
            description="Adds a calculated field to an existing PivotTable (e.g., Profit = Revenue - Cost).",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable."},
                    "field_name": {"type": "string", "description": "Name for the new calculated field."},
                    "formula": {"type": "string", "description": "Formula for the calculated field, starting with '=' (e.g., '=Revenue-Cost', \"='Field Name With Space' * 0.1\")."}
                },
                "required": ["path", "sheet_name", "pivot_table_name", "field_name", "formula"]
            }
        ),
        types.Tool(
            name="add_pivot_table_calculated_item",
            description="Adds a calculated item to a PivotField within a PivotTable (e.g., a 'North America' item in 'Region' field, summing 'USA' and 'Canada').",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable."},
                    "base_field_name": {"type": "string", "description": "The name of the PivotField where the calculated item will be added."},
                    "item_name": {"type": "string", "description": "Name for the new calculated item."},
                    "formula": {"type": "string", "description": "Formula for the calculated item (e.g., \"='USA' + 'Canada'\"). Item names with spaces must be in single quotes."}
                },
                "required": ["path", "sheet_name", "pivot_table_name", "base_field_name", "item_name", "formula"]
            }
        ),
        types.Tool(
            name="create_pivot_table_slicer",
            description="Creates a slicer for a specified field in a PivotTable, allowing for interactive filtering. The slicer is placed on a specified sheet.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the sheet where the slicer will be placed."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable the slicer will connect to."},
                    "slicer_field_name": {"type": "string", "description": "The name of the field from the PivotTable to be used for the slicer."},
                    "slicer_name": {"type": "string", "description": "Optional: A unique name for the slicer object. If not provided, a default name will be generated.", "default": None},
                    "top": {"type": "number", "description": "Optional: Position of the slicer from the top edge of the sheet (in points).", "default": None},
                    "left": {"type": "number", "description": "Optional: Position of the slicer from the left edge of the sheet (in points).", "default": None},
                    "width": {"type": "number", "description": "Optional: Width of the slicer (in points).", "default": None},
                    "height": {"type": "number", "description": "Optional: Height of the slicer (in points).", "default": None}
                },
                "required": ["path", "sheet_name", "pivot_table_name", "slicer_field_name"]
            }
        ),
        types.Tool(
            name="modify_pivot_table_slicer",
            description="Modifies properties of an existing PivotTable slicer, such as selected items, style, caption, position, and size.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the sheet where the slicer is located."},
                    "slicer_name": {"type": "string", "description": "The name of the slicer to modify."},
                    "selected_items": {
                        "type": "array", "items": {"type": "string"}, "default": None,
                        "description": "Optional: List of item names to select in the slicer. If None, current selection is unchanged. If an empty list, all items are deselected."
                    },
                    "slicer_style": {"type": "string", "description": "Optional: New style for the slicer (e.g., 'SlicerStyleLight1', 'SlicerStyleDark2').", "default": None},
                    "caption": {"type": "string", "description": "Optional: New caption for the slicer header.", "default": None},
                    "top": {"type": "number", "description": "Optional: New position from the top edge (in points).", "default": None},
                    "left": {"type": "number", "description": "Optional: New position from the left edge (in points).", "default": None},
                    "width": {"type": "number", "description": "Optional: New width of the slicer (in points).", "default": None},
                    "height": {"type": "number", "description": "Optional: New height of the slicer (in points).", "default": None},
                    "number_of_columns": {"type": "integer", "description": "Optional: Number of columns to display items in the slicer.", "default": None}
                },
                "required": ["path", "sheet_name", "slicer_name"]
            }
        ),
        types.Tool(
            name="set_pivot_table_layout",
            description="Changes the report layout of a PivotTable (Compact, Outline, or Tabular) and optionally repeats item labels.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable."},
                    "layout_type": {
                        "type": "string", "description": "The desired report layout.",
                        "enum": ["compact", "outline", "tabular"]
                    },
                    "repeat_all_item_labels": {
                        "type": "boolean", "default": None,
                        "description": "Optional: For Outline and Tabular layouts, set to True to repeat item labels for all row fields. Set to False to not repeat."
                    }
                },
                "required": ["path", "sheet_name", "pivot_table_name", "layout_type"]
            }
        ),
        types.Tool(
            name="configure_pivot_table_totals",
            description="Configures grand totals (for rows and/or columns) and subtotals for specific fields in a PivotTable.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable."},
                    "grand_totals_for_rows": {"type": "boolean", "default": None, "description": "Set to True to show grand totals for rows, False to hide. Null to leave unchanged."},
                    "grand_totals_for_columns": {"type": "boolean", "default": None, "description": "Set to True to show grand totals for columns, False to hide. Null to leave unchanged."},
                    "subtotals_settings": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "field_name": {"type": "string", "description": "Name of the row or column field to configure subtotals for."},
                                "show": {"type": "boolean", "description": "True to show subtotals for this field, False to hide."}
                            },
                            "required": ["field_name", "show"]
                        },
                        "default": None,
                        "description": "Optional: List of settings to configure subtotals for specific fields."
                    }
                },
                "required": ["path", "sheet_name", "pivot_table_name"]
            }
        ),
        types.Tool(
            name="format_pivot_table_part",
            description="Applies specific formatting (font, color, alignment) to different parts of a PivotTable like data body, headers, etc.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable."},
                    "part_to_format": {
                        "type": "string",
                        "description": "The part of the PivotTable to format.",
                        "enum": ["data_body_range", "row_header_range", "column_header_range", "page_field_range"] # "grand_total_range" is more complex
                    },
                    "font_bold": {"type": "boolean", "default": None},
                    "font_italic": {"type": "boolean", "default": None},
                    "font_size": {"type": "integer", "default": None},
                    "font_color_rgb": {
                        "type": "array", "items": {"type": "integer", "minimum": 0, "maximum": 255}, "minItems": 3, "maxItems": 3,
                        "description": "Font color as an RGB tuple (e.g., [255, 0, 0] for red).", "default": None
                    },
                    "bg_color_rgb": {
                        "type": "array", "items": {"type": "integer", "minimum": 0, "maximum": 255}, "minItems": 3, "maxItems": 3,
                        "description": "Background color as an RGB tuple (e.g., [255, 255, 0] for yellow).", "default": None
                    },
                    "horizontal_alignment": {
                        "type": "string", "default": None,
                        "description": "Horizontal text alignment.",
                        "enum": ["left", "center", "right", "general"]
                    }
                },
                "required": ["path", "sheet_name", "pivot_table_name", "part_to_format"]
            }
        ),
        types.Tool(
            name="change_pivot_table_data_source",
            description="Changes the source data range for an existing PivotTable and refreshes it.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable."},
                    "new_source_data": {"type": "string", "description": "The new data source range (e.g., 'Sheet1!A1:H200' or 'MyNewTable')."}
                },
                "required": ["path", "sheet_name", "pivot_table_name", "new_source_data"]
            }
        ),
        # New tool definitions
        types.Tool(
            name="set_pivot_table_value_field_calculation",
            description="Configures how values are displayed in a PivotTable, such as percentage of total, difference from, running total, ranking, etc.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable."},
                    "value_field_caption": {"type": "string", "description": "The caption of the value field to modify (e.g., 'Sum of Sales')."},
                    "calculation_type": {
                        "type": "string", 
                        "description": "The type of calculation to apply ('normal', '% of total', 'difference_from', 'running_total', etc).",
                        "enum": [
                            "normal", "% of total", "% of row", "% of column", 
                            "difference_from", "% difference_from", 
                            "running_total", "% of running_total", 
                            "rank", "index"
                        ]
                    },
                    "base_field": {
                        "type": "string", 
                        "description": "For relative calculations (e.g., difference_from), the field to calculate relative to.",
                        "default": None
                    },
                    "base_item": {
                        "type": "string", 
                        "description": "For relative calculations (e.g., difference_from), the specific item to calculate relative to.",
                        "default": None
                    }
                },
                "required": ["path", "sheet_name", "pivot_table_name", "value_field_caption", "calculation_type"]
            }
        ),
        types.Tool(
            name="group_pivot_field_items",
            description="Groups items in a PivotTable field based on date ranges, numeric ranges, or manual selection.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable."},
                    "field_name": {"type": "string", "description": "Name of the field to group."},
                    "group_type": {
                        "type": "string", 
                        "description": "Type of grouping to perform.",
                        "enum": ["date", "numeric", "selection"],
                        "default": "date"
                    },
                    "start_value": {
                        "type": "number", 
                        "description": "For numeric grouping, the start value of the range.",
                        "default": None
                    },
                    "end_value": {
                        "type": "number", 
                        "description": "For numeric grouping, the end value of the range.",
                        "default": None
                    },
                    "interval": {
                        "type": "number", 
                        "description": "For numeric grouping, the interval size.",
                        "default": None
                    },
                    "date_parts": {
                        "type": "object", 
                        "description": "For date grouping, which date parts to include (e.g., {'years': true, 'quarters': true, 'months': true}).",
                        "default": None
                    },
                    "selected_items": {
                        "type": "array", 
                        "items": {"type": "string"},
                        "description": "For selection grouping, the items to group together.",
                        "default": None
                    }
                },
                "required": ["path", "sheet_name", "pivot_table_name", "field_name", "group_type"]
            }
        ),
        types.Tool(
            name="ungroup_pivot_field_items",
            description="Removes grouping from a PivotTable field.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable."},
                    "field_name": {"type": "string", "description": "Name of the field to ungroup."},
                    "group_name": {
                        "type": "string", 
                        "description": "Optional name of a specific group to ungroup. If not provided, ungroups all.",
                        "default": None
                    }
                },
                "required": ["path", "sheet_name", "pivot_table_name", "field_name"]
            }
        ),
        types.Tool(
            name="apply_pivot_table_conditional_formatting",
            description="Applies conditional formatting to specific parts of a PivotTable based on various criteria.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet containing the PivotTable."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable."},
                    "formatting_scope": {
                        "type": "string", 
                        "description": "Which part of the PivotTable to format.",
                        "enum": ["data_field", "field_items", "grand_totals", "subtotals"],
                        "default": "data_field"
                    },
                    "field_name": {
                        "type": "string", 
                        "description": "For data_field: the value field caption (e.g. 'Sum of Sales'). For field_items: the field to format items for (e.g. 'Region')."
                    },
                    "condition_type": {
                        "type": "string", 
                        "description": "The type of condition to apply for the formatting.",
                        "enum": ["top_bottom", "greater_than", "less_than", "between", "equal_to", "contains", "date_occurring"],
                        "default": "top_bottom"
                    },
                    "condition_parameters": {
                        "type": "object", 
                        "description": "Parameters specific to the condition_type (e.g., {'rank': 5, 'type': 'top', 'percent': True} for top 5%)."
                    },
                    "format_settings": {
                        "type": "object", 
                        "description": "Formatting to apply, such as {'bold': True, 'bg_color_rgb': [255,0,0]} for bold text with red background."
                    },
                    "specific_items": {
                        "type": "array", 
                        "items": {"type": "string"},
                        "description": "For field_items scope, optionally limit formatting to these specific items.",
                        "default": None
                    }
                },
                "required": ["path", "sheet_name", "pivot_table_name", "formatting_scope", "field_name", "condition_type"]
            }
        ),
        types.Tool(
            name="create_timeline_slicer",
            description="Creates a timeline slicer for date fields in a PivotTable, providing a specialized date-filtering interface. Timeline slicers offer intuitive filtering by days, months, quarters, and years.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet where the timeline will be placed."},
                    "pivot_table_name": {"type": "string", "description": "Name of the PivotTable to connect the timeline to."},
                    "date_field_name": {"type": "string", "description": "Name of the date field from the PivotTable to use for the timeline."},
                    "timeline_name": {
                        "type": "string", 
                        "description": "Optional custom name for the timeline. If not provided, a default name will be generated.",
                        "default": None
                    },
                    "top": {
                        "type": "number", 
                        "description": "Position of the timeline from the top edge of the sheet (in points).",
                        "default": None
                    },
                    "left": {
                        "type": "number", 
                        "description": "Position of the timeline from the left edge of the sheet (in points).",
                        "default": None
                    },
                    "width": {
                        "type": "number", 
                        "description": "Width of the timeline (in points).",
                        "default": None
                    },
                    "height": {
                        "type": "number", 
                        "description": "Height of the timeline (in points).",
                        "default": None
                    },
                    "time_level": {
                        "type": "string", 
                        "description": "Default timeline level to display.",
                        "enum": ["days", "months", "quarters", "years"],
                        "default": "months"
                    }
                },
                "required": ["path", "sheet_name", "pivot_table_name", "date_field_name"]
            }
        ),
        types.Tool(
            name="connect_slicer_to_pivot_tables",
            description="Connects a single slicer or timeline to multiple PivotTables, allowing synchronized filtering across them, which is perfect for interactive dashboards.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "sheet_name": {"type": "string", "description": "Name of the worksheet where the slicer is located."},
                    "slicer_name": {"type": "string", "description": "Name of the slicer or timeline to connect."},
                    "pivot_table_names": {
                        "type": "array", 
                        "items": {"type": "string"},
                        "description": "List of PivotTable names to connect to the slicer."
                    }
                },
                "required": ["path", "sheet_name", "slicer_name", "pivot_table_names"]
            }
        ),
        types.Tool(
            name="setup_power_pivot_data_model",
            description="Sets up a Power Pivot data model by importing external data sources and establishing relationships between tables in the Excel workbook.",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "data_sources": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "source_type": {
                                    "type": "string",
                                    "description": "Type of data source.",
                                    "enum": ["excel", "csv", "database", "worksheet_table"]
                                },
                                "location": {"type": "string", "description": "Path or connection string for the data source."},
                                "target_table_name": {"type": "string", "description": "Name for the table in the data model."},
                                "properties": {
                                    "type": "object",
                                    "description": "Source-specific properties like sheet_name, table_name, connection_string, query, etc."
                                }
                            },
                            "required": ["source_type", "location", "target_table_name"]
                        },
                        "description": "List of data sources to add to the model."
                    },
                    "relationships": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "from_table": {"type": "string", "description": "Parent table name."},
                                "from_column": {"type": "string", "description": "Column in the parent table."},
                                "to_table": {"type": "string", "description": "Child table name."},
                                "to_column": {"type": "string", "description": "Column in the child table."},
                                "active": {"type": "boolean", "description": "Whether this is an active relationship.", "default": True}
                            },
                            "required": ["from_table", "from_column", "to_table", "to_column"]
                        },
                        "description": "List of relationships to define between tables."
                    }
                },
                "required": ["path", "data_sources"]
            }
        ),
        types.Tool(
            name="create_power_pivot_measure",
            description="Creates a new calculated measure in the Power Pivot data model using DAX (Data Analysis Expressions).",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Path to the Excel workbook."},
                    "measure_name": {"type": "string", "description": "Name for the new measure."},
                    "dax_formula": {"type": "string", "description": "DAX formula that defines the measure's calculation."},
                    "table_name": {"type": "string", "description": "Name of the table where the measure will be displayed."},
                    "display_folder": {
                        "type": "string", 
                        "description": "Optional folder path to organize measures in the field list.",
                        "default": None
                    },
                    "format_string": {
                        "type": "string", 
                        "description": "Optional format string for the measure (e.g., '$#,##0.00' for currency).",
                        "default": None
                    }
                },
                "required": ["path", "measure_name", "dax_formula", "table_name"]
            }
        ),
        
        # --- PIVOTCHART TOOL DEFINITIONS ---
        
        types.Tool(
            name="create_pivot_chart",
            description="Create pivot charts from various data sources with full automation. This is your go-to tool for creating "
                       "professional pivot charts from either existing pivot tables, data ranges, or by creating new pivot tables. "
                       "Supports all major chart types (column, bar, line, pie, combo) with automatic styling and formatting. "
                       "Perfect for transforming raw data into insightful visualizations instantly.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_path": {
                        "type": "string",
                        "description": "Path to the Excel workbook. Can be absolute path, relative path, or filename to search for."
                    },
                    "source_type": {
                        "type": "string",
                        "enum": ["pivot_table", "data_range", "new_pivot"],
                        "description": "Data source type: 'pivot_table' for existing pivot table, 'data_range' for direct chart from range, 'new_pivot' to create pivot table first then chart"
                    },
                    "data_range": {
                        "type": "string",
                        "description": "Excel range (e.g., 'A1:D100') - required for 'data_range' and 'new_pivot' source types"
                    },
                    "pivot_table_name": {
                        "type": "string",
                        "description": "Name of existing pivot table - required for 'pivot_table' source type"
                    },
                    "chart_type": {
                        "type": "string",
                        "enum": ["COLUMN", "BAR", "LINE", "PIE", "AREA", "DOUGHNUT", "COMBO", "COLUMN_STACKED", "BAR_STACKED", "LINE_MARKERS", "SCATTER"],
                        "default": "COLUMN",
                        "description": "Type of chart to create. Choose based on your data visualization needs."
                    },
                    "chart_title": {
                        "type": "string",
                        "description": "Title for the chart. If not provided, chart will have no title."
                    },
                    "position": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "minItems": 2,
                        "maxItems": 2,
                        "default": [100, 100],
                        "description": "Chart position as [x, y] coordinates in pixels from top-left corner"
                    },
                    "pivot_config": {
                        "type": "object",
                        "properties": {
                            "row_fields": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Fields to use as row categories in pivot table"
                            },
                            "column_fields": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Fields to use as column categories in pivot table"
                            },
                            "value_fields": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Fields to use as values/measures in pivot table"
                            },
                            "destination": {
                                "type": "string",
                                "default": "H1",
                                "description": "Where to place the pivot table (e.g., 'H1' or 'Sheet2!A1')"
                            }
                        },
                        "description": "Pivot table configuration - required for 'new_pivot' source type"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Specific worksheet name to work with. If not provided, uses active sheet."
                    }
                },
                "required": ["workbook_path", "source_type"],
                "additionalProperties": False
            }
        ),
        
        types.Tool(
            name="manage_chart_elements",
            description="Comprehensive chart formatting and element management tool. Control every aspect of your chart's appearance "
                       "including titles, legends, data labels, gridlines, and axis formatting. This tool handles all visual "
                       "customization needs to make your charts presentation-ready with professional styling.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_path": {
                        "type": "string",
                        "description": "Path to the Excel workbook containing the chart"
                    },
                    "chart_name": {
                        "type": "string",
                        "description": "Name of the chart to modify. Use 'list_charts' tool first if you don't know the name."
                    },
                    "title_config": {
                        "type": "object",
                        "properties": {
                            "show_title": {"type": "boolean", "default": True},
                            "title_text": {"type": "string", "description": "Chart title text"}
                        },
                        "description": "Chart title configuration"
                    },
                    "axis_config": {
                        "type": "object",
                        "properties": {
                            "x_axis_title": {"type": "string", "description": "X-axis title"},
                            "y_axis_title": {"type": "string", "description": "Y-axis title"},
                            "show_x_title": {"type": "boolean", "default": True},
                            "show_y_title": {"type": "boolean", "default": True}
                        },
                        "description": "Axis titles configuration"
                    },
                    "legend_config": {
                        "type": "object",
                        "properties": {
                            "show_legend": {"type": "boolean", "default": True},
                            "position": {
                                "type": "string",
                                "enum": ["RIGHT", "LEFT", "TOP", "BOTTOM", "CORNER"],
                                "default": "RIGHT",
                                "description": "Legend position"
                            }
                        },
                        "description": "Legend configuration"
                    },
                    "data_labels": {
                        "type": "object",
                        "properties": {
                            "show_labels": {"type": "boolean", "default": False},
                            "series_index": {"type": "integer", "default": 1, "description": "Which series to apply labels to (1-based)"}
                        },
                        "description": "Data labels configuration"
                    },
                    "gridlines": {
                        "type": "object",
                        "properties": {
                            "x_major": {"type": "boolean", "description": "Show X-axis major gridlines"},
                            "x_minor": {"type": "boolean", "description": "Show X-axis minor gridlines"},
                            "y_major": {"type": "boolean", "description": "Show Y-axis major gridlines"},
                            "y_minor": {"type": "boolean", "description": "Show Y-axis minor gridlines"}
                        },
                        "description": "Gridlines configuration"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Specific worksheet name if chart is on particular sheet"
                    }
                },
                "required": ["workbook_path", "chart_name"],
                "additionalProperties": False
            }
        ),
        
        types.Tool(
            name="apply_chart_styling",
            description="Apply professional styling and layouts to charts instantly. Transform basic charts into polished, "
                       "corporate-ready visualizations using Excel's built-in styles and layouts. Also supports dynamic "
                       "chart type changes for different data presentation needs. Perfect for creating consistent, branded charts.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_path": {
                        "type": "string",
                        "description": "Path to the Excel workbook containing the chart"
                    },
                    "chart_name": {
                        "type": "string",
                        "description": "Name of the chart to style"
                    },
                    "style_id": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 48,
                        "description": "Excel chart style ID (1-48). Different IDs provide different color schemes and formatting."
                    },
                    "layout_id": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 11,
                        "description": "Excel chart layout ID (1-11). Controls overall chart element arrangement and design."
                    },
                    "new_chart_type": {
                        "type": "string",
                        "enum": ["COLUMN", "BAR", "LINE", "PIE", "AREA", "DOUGHNUT", "COMBO", "COLUMN_STACKED", "BAR_STACKED", "LINE_MARKERS", "SCATTER"],
                        "description": "Change chart type dynamically. Useful for exploring different data visualizations."
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Specific worksheet name if chart is on particular sheet"
                    }
                },
                "required": ["workbook_path", "chart_name"],
                "additionalProperties": False
            }
        ),
        
        types.Tool(
            name="manage_pivot_fields",
            description="Advanced pivot table field management and calculated field creation. Dynamically modify pivot table "
                       "structure by changing field orientations, adding/removing fields, creating calculated fields for "
                       "complex analysis (like profit margins, ratios, growth rates), and configuring summary functions. "
                       "Essential for creating sophisticated business intelligence reports.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_path": {
                        "type": "string",
                        "description": "Path to the Excel workbook containing the pivot table"
                    },
                    "pivot_table_name": {
                        "type": "string",
                        "description": "Name of the pivot table to modify"
                    },
                    "field_operations": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "field_name": {"type": "string", "description": "Name of the field to modify"},
                                "orientation": {
                                    "type": "string",
                                    "enum": ["ROW", "COLUMN", "DATA", "PAGE", "HIDDEN"],
                                    "description": "Where to place this field in the pivot table"
                                },
                                "summary_function": {
                                    "type": "string",
                                    "enum": ["SUM", "COUNT", "AVERAGE", "MAX", "MIN", "PRODUCT"],
                                    "description": "Summary function for data fields"
                                }
                            },
                            "required": ["field_name", "orientation"]
                        },
                        "description": "List of field operations to perform"
                    },
                    "calculated_fields": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "field_name": {"type": "string", "description": "Name for the calculated field"},
                                "formula": {"type": "string", "description": "Excel formula for calculation (e.g., '=Sales-Costs' for profit)"}
                            },
                            "required": ["field_name", "formula"]
                        },
                        "description": "Calculated fields to create for advanced analysis"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Specific worksheet name if pivot table is on particular sheet"
                    }
                },
                "required": ["workbook_path", "pivot_table_name"],
                "additionalProperties": False
            }
        ),
        
        types.Tool(
            name="create_combo_chart",
            description="Create sophisticated combination charts that display multiple data series with different chart types "
                       "on primary and secondary axes. Perfect for comparing different types of metrics (e.g., sales volumes vs. "
                       "profit percentages, actual vs. targets). Enables complex data storytelling in a single visualization.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_path": {
                        "type": "string",
                        "description": "Path to the Excel workbook"
                    },
                    "data_range": {
                        "type": "string",
                        "description": "Excel range containing the data (e.g., 'A1:E12')"
                    },
                    "chart_title": {
                        "type": "string",
                        "description": "Title for the combination chart"
                    },
                    "primary_series": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Names of data series to display with primary chart type (left axis)"
                    },
                    "secondary_series": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Names of data series to display with secondary chart type (right axis)"
                    },
                    "primary_type": {
                        "type": "string",
                        "enum": ["COLUMN", "BAR", "LINE", "AREA"],
                        "default": "COLUMN",
                        "description": "Chart type for primary series"
                    },
                    "secondary_type": {
                        "type": "string",
                        "enum": ["LINE", "COLUMN", "BAR", "AREA"],
                        "default": "LINE",
                        "description": "Chart type for secondary series"
                    },
                    "position": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "minItems": 2,
                        "maxItems": 2,
                        "default": [100, 100],
                        "description": "Chart position as [x, y] coordinates"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Specific worksheet name to create chart on"
                    }
                },
                "required": ["workbook_path", "data_range", "primary_series", "secondary_series"],
                "additionalProperties": False
            }
        ),
        
        types.Tool(
            name="add_chart_filters",
            description="Add interactive slicers and filters to pivot charts for dynamic data exploration. Create user-friendly "
                       "filter controls that allow end-users to slice and dice data without technical knowledge. Perfect for "
                       "building interactive dashboards and self-service analytics solutions.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_path": {
                        "type": "string",
                        "description": "Path to the Excel workbook"
                    },
                    "pivot_table_name": {
                        "type": "string",
                        "description": "Name of the pivot table to add slicers to"
                    },
                    "slicer_fields": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "field_name": {"type": "string", "description": "Field to create slicer for"},
                                "position": {
                                    "type": "array",
                                    "items": {"type": "integer"},
                                    "minItems": 2,
                                    "maxItems": 2,
                                    "default": [500, 100],
                                    "description": "Slicer position as [x, y] coordinates"
                                }
                            },
                            "required": ["field_name"]
                        },
                        "description": "List of fields to create slicers for"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Specific worksheet name"
                    }
                },
                "required": ["workbook_path", "pivot_table_name", "slicer_fields"],
                "additionalProperties": False
            }
        ),
        
        types.Tool(
            name="refresh_and_update",
            description="Refresh pivot tables, charts, and data connections to ensure all visualizations reflect the latest data. "
                       "Also supports dynamic data source switching for charts. Essential for maintaining accurate, up-to-date "
                       "reports and dashboards in live data environments.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_path": {
                        "type": "string",
                        "description": "Path to the Excel workbook"
                    },
                    "operation": {
                        "type": "string",
                        "enum": ["refresh_all", "refresh_pivot", "update_chart_source"],
                        "description": "Type of refresh/update operation to perform"
                    },
                    "pivot_table_name": {
                        "type": "string",
                        "description": "Specific pivot table name to refresh (for 'refresh_pivot' operation)"
                    },
                    "chart_name": {
                        "type": "string",
                        "description": "Chart name for data source update (for 'update_chart_source' operation)"
                    },
                    "new_data_range": {
                        "type": "string",
                        "description": "New data range for chart (for 'update_chart_source' operation)"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Specific worksheet name"
                    }
                },
                "required": ["workbook_path", "operation"],
                "additionalProperties": False
            }
        ),
        
        types.Tool(
            name="export_and_distribute",
            description="Export charts and create multi-chart dashboards for distribution. Save charts as images (PNG, JPG, PDF) "
                       "for presentations and reports, or create comprehensive dashboard layouts with multiple related charts. "
                       "Streamlines the process of sharing insights and creating professional reports.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_path": {
                        "type": "string",
                        "description": "Path to the Excel workbook"
                    },
                    "operation": {
                        "type": "string",
                        "enum": ["export_chart", "create_dashboard"],
                        "description": "Export single chart or create multi-chart dashboard"
                    },
                    "chart_name": {
                        "type": "string",
                        "description": "Name of chart to export (for 'export_chart' operation)"
                    },
                    "export_path": {
                        "type": "string",
                        "description": "File path for export (for 'export_chart' operation)"
                    },
                    "file_format": {
                        "type": "string",
                        "enum": ["PNG", "JPG", "JPEG", "GIF", "PDF"],
                        "default": "PNG",
                        "description": "Export file format"
                    },
                    "dashboard_config": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "type": {
                                    "type": "string",
                                    "enum": ["pivot", "range"],
                                    "description": "Chart creation type"
                                },
                                "params": {
                                    "type": "object",
                                    "description": "Parameters for chart creation"
                                }
                            },
                            "required": ["type", "params"]
                        },
                        "description": "Configuration for multiple charts in dashboard (for 'create_dashboard' operation)"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Specific worksheet name"
                    }
                },
                "required": ["workbook_path", "operation"],
                "additionalProperties": False
            }
        ),
        
        types.Tool(
            name="get_chart_info",
            description="Retrieve comprehensive information about charts, pivot tables, and data sources in Excel workbooks. "
                       "Perfect for discovery and analysis of existing Excel files, understanding data structures, and getting "
                       "chart specifications. Essential for troubleshooting and planning chart modifications.",
            inputSchema={
                "type": "object",
                "properties": {
                    "workbook_path": {
                        "type": "string",
                        "description": "Path to the Excel workbook"
                    },
                    "info_type": {
                        "type": "string",
                        "enum": ["list_charts", "list_pivot_tables", "chart_details", "workbook_overview"],
                        "description": "Type of information to retrieve"
                    },
                    "chart_name": {
                        "type": "string",
                        "description": "Specific chart name for detailed information (for 'chart_details' type)"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Specific worksheet name to focus on"
                    }
                },
                "required": ["workbook_path", "info_type"],
                "additionalProperties": False
            },
            idempotentHint=True,
            readOnlyHint=True
        )
    ] 


