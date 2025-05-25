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
        # Add the add_excel_column tool definition
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
        # Update the conditional formatting tool
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
        )
    ] 
