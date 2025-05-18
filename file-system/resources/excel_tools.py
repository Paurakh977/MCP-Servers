"""Excel workbook and worksheet manipulation tools for the MCP server.

This module provides functions for creating, reading, and manipulating Excel workbooks,
worksheets, ranges, and formatting. It builds upon openpyxl for Excel operations.
"""

import os
import logging
import json
from pathlib import Path
from datetime import datetime
from typing import Any, Dict, List, Optional, Union, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Border, PatternFill, Side, Alignment
from openpyxl.styles.colors import Color

# Set up logging
logger = logging.getLogger(__name__)

# Exceptions
class ExcelError(Exception):
    """Base exception for Excel operations"""
    pass

class WorkbookError(ExcelError):
    """Error related to workbook operations"""
    pass

class SheetError(ExcelError):
    """Error related to sheet operations"""
    pass

class RangeError(ExcelError):
    """Error related to range operations"""
    pass

class ValidationError(ExcelError):
    """Error related to validation"""
    pass

# Helper Functions
def parse_cell_reference(cell_ref: str) -> Tuple[int, int]:
    """Parse a cell reference (e.g. 'A1') into row and column indices"""
    if not cell_ref or not isinstance(cell_ref, str):
        raise ValidationError(f"Invalid cell reference: {cell_ref}")
        
    # Extract column letters and row numbers
    col_str = ''.join(c for c in cell_ref if c.isalpha())
    row_str = ''.join(c for c in cell_ref if c.isdigit())
    
    if not col_str or not row_str:
        raise ValidationError(f"Invalid cell reference format: {cell_ref}")
        
    try:
        col_idx = column_index_from_string(col_str)
        row_idx = int(row_str)
        return row_idx, col_idx
    except (ValueError, KeyError) as e:
        raise ValidationError(f"Invalid cell reference: {cell_ref} - {str(e)}")

def parse_cell_range(start_cell: str, end_cell: Optional[str] = None) -> Tuple[int, int, Optional[int], Optional[int]]:
    """Parse a cell range into row and column indices"""
    start_row, start_col = parse_cell_reference(start_cell)
    
    if end_cell:
        end_row, end_col = parse_cell_reference(end_cell)
        return start_row, start_col, end_row, end_col
    else:
        return start_row, start_col, None, None

def format_range_string(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    """Format a range string from row and column indices"""
    return f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"

def find_file_in_allowed_dirs(filename: str, allowed_dirs: List[str]) -> Optional[str]:
    """Find a file in the allowed directories"""
    for base_dir in allowed_dirs:
        for root, _, files in os.walk(base_dir):
            matching_files = [f for f in files if filename.lower() in f.lower()]
            if matching_files:
                return os.path.join(root, matching_files[0])
    return None

# Workbook Operations
def create_excel_workbook(filepath: str, sheet_name: str = "Sheet1") -> Dict[str, Any]:
    """Create a new Excel workbook with optional custom sheet name"""
    try:
        wb = Workbook()
        # Rename default sheet
        if "Sheet" in wb.sheetnames:
            sheet = wb["Sheet"]
            sheet.title = sheet_name
        else:
            wb.create_sheet(sheet_name)

        path = Path(filepath)
        path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(str(path))
        return {
            "success": True,
            "message": f"Created workbook: {filepath}",
            "active_sheet": sheet_name,
        }
    except Exception as e:
        logger.error(f"Failed to create workbook: {e}")
        return {
            "success": False,
            "error": f"Failed to create workbook: {str(e)}"
        }

def get_workbook_metadata(filepath: str, include_ranges: bool = False) -> Dict[str, Any]:
    """Get detailed metadata about an Excel workbook"""
    try:
        path = Path(filepath)
        if not path.exists():
            return {
                "success": False,
                "error": f"File not found: {filepath}"
            }
            
        wb = load_workbook(filepath, read_only=True)
        
        sheets_info = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            sheet_info = {
                "name": sheet_name,
                "rows": ws.max_row,
                "columns": ws.max_column,
            }
            
            if include_ranges and ws.max_row > 0 and ws.max_column > 0:
                sheet_info["used_range"] = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
                
                # Try to get column headers
                try:
                    headers = []
                    for col in range(1, ws.max_column + 1):
                        cell = ws.cell(row=1, column=col)
                        header = cell.value if cell.value is not None else f"Column {get_column_letter(col)}"
                        headers.append(str(header))
                    sheet_info["headers"] = headers
                except:
                    pass  # Skip headers if there's an error
            
            sheets_info.append(sheet_info)
        
        info = {
            "success": True,
            "filename": path.name,
            "file_size": path.stat().st_size,
            "file_size_human": f"{path.stat().st_size / 1024:.2f} KB",
            "modified": path.stat().st_mtime,
            "modified_date": datetime.fromtimestamp(path.stat().st_mtime).isoformat(),
            "sheets_count": len(wb.sheetnames),
            "sheets": sheets_info
        }
        
        wb.close()
        return info
        
    except Exception as e:
        logger.error(f"Failed to get workbook metadata: {e}")
        return {
            "success": False,
            "error": f"Failed to get workbook metadata: {str(e)}"
        }

# Worksheet Operations
def create_worksheet(filepath: str, sheet_name: str) -> Dict[str, Any]:
    """Create a new worksheet in an Excel workbook"""
    try:
        wb = load_workbook(filepath)

        # Check if sheet already exists
        if sheet_name in wb.sheetnames:
            return {
                "success": False,
                "error": f"Sheet '{sheet_name}' already exists"
            }

        # Create new sheet
        wb.create_sheet(sheet_name)
        wb.save(filepath)
        wb.close()
        return {
            "success": True,
            "message": f"Sheet '{sheet_name}' created successfully"
        }
    except Exception as e:
        logger.error(f"Failed to create worksheet: {e}")
        return {
            "success": False,
            "error": f"Failed to create worksheet: {str(e)}"
        }

def copy_worksheet(filepath: str, source_sheet: str, target_sheet: str) -> Dict[str, Any]:
    """Copy a worksheet within the same workbook"""
    try:
        wb = load_workbook(filepath)
        
        if source_sheet not in wb.sheetnames:
            return {
                "success": False,
                "error": f"Source sheet '{source_sheet}' not found"
            }
            
        if target_sheet in wb.sheetnames:
            return {
                "success": False,
                "error": f"Target sheet '{target_sheet}' already exists"
            }
            
        source = wb[source_sheet]
        target = wb.copy_worksheet(source)
        target.title = target_sheet
        
        wb.save(filepath)
        wb.close()
        return {
            "success": True,
            "message": f"Sheet '{source_sheet}' copied to '{target_sheet}'"
        }
    except Exception as e:
        logger.error(f"Failed to copy worksheet: {e}")
        return {
            "success": False,
            "error": f"Failed to copy worksheet: {str(e)}"
        }

def delete_worksheet(filepath: str, sheet_name: str) -> Dict[str, Any]:
    """Delete a worksheet from an Excel workbook"""
    try:
        wb = load_workbook(filepath)
        
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "error": f"Sheet '{sheet_name}' not found"
            }
            
        if len(wb.sheetnames) == 1:
            return {
                "success": False,
                "error": "Cannot delete the only sheet in workbook"
            }
            
        del wb[sheet_name]
        wb.save(filepath)
        wb.close()
        return {
            "success": True,
            "message": f"Sheet '{sheet_name}' deleted successfully"
        }
    except Exception as e:
        logger.error(f"Failed to delete worksheet: {e}")
        return {
            "success": False,
            "error": f"Failed to delete worksheet: {str(e)}"
        }

def rename_worksheet(filepath: str, old_name: str, new_name: str) -> Dict[str, Any]:
    """Rename a worksheet in an Excel workbook"""
    try:
        wb = load_workbook(filepath)
        
        if old_name not in wb.sheetnames:
            return {
                "success": False,
                "error": f"Sheet '{old_name}' not found"
            }
            
        if new_name in wb.sheetnames:
            return {
                "success": False,
                "error": f"Sheet '{new_name}' already exists"
            }
            
        sheet = wb[old_name]
        sheet.title = new_name
        wb.save(filepath)
        wb.close()
        return {
            "success": True,
            "message": f"Sheet renamed from '{old_name}' to '{new_name}'"
        }
    except Exception as e:
        logger.error(f"Failed to rename worksheet: {e}")
        return {
            "success": False,
            "error": f"Failed to rename worksheet: {str(e)}"
        }

# Range Operations
def copy_excel_range(
    filepath: str, 
    sheet_name: str, 
    source_range: str, 
    target_start: str, 
    target_sheet: Optional[str] = None
) -> Dict[str, Any]:
    """Copy a range of cells to another location in Excel"""
    try:
        wb = load_workbook(filepath)
        
        # Validate source sheet
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "error": f"Source sheet '{sheet_name}' not found"
            }
        
        source_ws = wb[sheet_name]
        
        # Validate target sheet
        if target_sheet and target_sheet not in wb.sheetnames:
            return {
                "success": False,
                "error": f"Target sheet '{target_sheet}' not found"
            }
        
        target_ws = wb[target_sheet] if target_sheet else source_ws
        
        # Parse source range
        try:
            if ':' in source_range:
                source_start, source_end = source_range.split(':')
                src_start_row, src_start_col, src_end_row, src_end_col = parse_cell_range(
                    source_start, source_end
                )
            else:
                src_start_row, src_start_col, src_end_row, src_end_col = parse_cell_range(
                    source_range
                )
                src_end_row = src_start_row
                src_end_col = src_start_col
        except ValidationError as e:
            return {
                "success": False,
                "error": f"Invalid source range: {str(e)}"
            }
            
        # Parse target start
        try:
            tgt_start_row, tgt_start_col, _, _ = parse_cell_range(target_start)
        except ValidationError as e:
            return {
                "success": False,
                "error": f"Invalid target cell: {str(e)}"
            }
            
        # Copy range
        for i, row in enumerate(range(src_start_row, src_end_row + 1)):
            for j, col in enumerate(range(src_start_col, src_end_col + 1)):
                try:
                    source_cell = source_ws.cell(row=row, column=col)
                    target_cell = target_ws.cell(row=tgt_start_row + i, column=tgt_start_col + j)
                    
                    # Copy value
                    target_cell.value = source_cell.value
                    
                    # Copy style if possible
                    if hasattr(source_cell, '_style') and source_cell._style:
                        target_cell._style = source_cell._style.copy()
                except Exception as e:
                    logger.warning(f"Could not copy cell at row {row}, column {col}: {e}")
                    continue
                    
        # Save workbook
        wb.save(filepath)
        wb.close()
        
        # Prepare success message
        source_range_str = source_range if ':' in source_range else source_range
        target_info = f" to {target_sheet}" if target_sheet else ""
        
        return {
            "success": True,
            "message": f"Range {source_range_str} copied from {sheet_name} to {target_start}{target_info}"
        }
    except Exception as e:
        logger.error(f"Failed to copy range: {e}")
        return {
            "success": False,
            "error": f"Failed to copy range: {str(e)}"
        }

def delete_excel_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: Optional[str] = None,
    shift_direction: str = "up"
) -> Dict[str, Any]:
    """Delete a range of cells and shift remaining cells in Excel"""
    try:
        wb = load_workbook(filepath)
        
        # Validate sheet
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "error": f"Sheet '{sheet_name}' not found"
            }
            
        worksheet = wb[sheet_name]
        
        # Validate range
        try:
            start_row, start_col, end_row, end_col = parse_cell_range(start_cell, end_cell)
            
            if end_row is None:
                end_row = start_row
                end_col = start_col
                
            if end_row > worksheet.max_row:
                return {
                    "success": False,
                    "error": f"End row {end_row} out of bounds (1-{worksheet.max_row})"
                }
                
            if end_col > worksheet.max_column:
                return {
                    "success": False,
                    "error": f"End column {end_col} out of bounds (1-{worksheet.max_column})"
                }
        except ValidationError as e:
            return {
                "success": False,
                "error": f"Invalid range: {str(e)}"
            }
            
        # Validate shift direction
        if shift_direction not in ["up", "left"]:
            return {
                "success": False,
                "error": f"Invalid shift direction: {shift_direction}. Must be 'up' or 'left'"
            }
            
        # Format range string for the message
        range_string = format_range_string(start_row, start_col, end_row, end_col)
        
        # Clear range contents first
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = worksheet.cell(row=row, column=col)
                cell.value = None
                
        # Shift cells if requested
        if shift_direction == "up":
            # Calculate how many rows to delete
            num_rows = end_row - start_row + 1
            worksheet.delete_rows(start_row, num_rows)
        elif shift_direction == "left":
            # Calculate how many columns to delete
            num_cols = end_col - start_col + 1
            worksheet.delete_cols(start_col, num_cols)
            
        # Save workbook
        wb.save(filepath)
        wb.close()
        
        action = "deleted and cells shifted" if shift_direction else "cleared"
        return {
            "success": True,
            "message": f"Range {range_string} {action} in sheet '{sheet_name}'"
        }
    except Exception as e:
        logger.error(f"Failed to delete range: {e}")
        return {
            "success": False,
            "error": f"Failed to delete range: {str(e)}"
        }

def merge_excel_cells(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str
) -> Dict[str, Any]:
    """Merge a range of cells in Excel"""
    try:
        wb = load_workbook(filepath)
        
        # Validate sheet
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "error": f"Sheet '{sheet_name}' not found"
            }
            
        worksheet = wb[sheet_name]
        
        # Validate range
        try:
            start_row, start_col, end_row, end_col = parse_cell_range(start_cell, end_cell)
            
            if end_row is None or end_col is None:
                return {
                    "success": False,
                    "error": "Both start and end cells must be specified for merging"
                }
                
            if end_row > worksheet.max_row:
                return {
                    "success": False,
                    "error": f"End row {end_row} out of bounds (1-{worksheet.max_row})"
                }
                
            if end_col > worksheet.max_column:
                return {
                    "success": False,
                    "error": f"End column {end_col} out of bounds (1-{worksheet.max_column})"
                }
                
        except ValidationError as e:
            return {
                "success": False,
                "error": f"Invalid range: {str(e)}"
            }
            
        # Format range string
        range_string = format_range_string(start_row, start_col, end_row, end_col)
        
        # Check if already merged
        for merged_range in worksheet.merged_cells.ranges:
            if str(merged_range) == range_string:
                return {
                    "success": False,
                    "error": f"Range {range_string} is already merged"
                }
            
        # Merge cells
        worksheet.merge_cells(range_string)
        
        # Save workbook
        wb.save(filepath)
        wb.close()
        
        return {
            "success": True,
            "message": f"Range {range_string} merged in sheet '{sheet_name}'"
        }
    except Exception as e:
        logger.error(f"Failed to merge cells: {e}")
        return {
            "success": False,
            "error": f"Failed to merge cells: {str(e)}"
        }

def unmerge_excel_cells(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str
) -> Dict[str, Any]:
    """Unmerge a previously merged range of cells in Excel"""
    try:
        wb = load_workbook(filepath)
        
        # Validate sheet
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "error": f"Sheet '{sheet_name}' not found"
            }
            
        worksheet = wb[sheet_name]
        
        # Validate range
        try:
            start_row, start_col, end_row, end_col = parse_cell_range(start_cell, end_cell)
            
            if end_row is None or end_col is None:
                return {
                    "success": False,
                    "error": "Both start and end cells must be specified for unmerging"
                }
                
        except ValidationError as e:
            return {
                "success": False,
                "error": f"Invalid range: {str(e)}"
            }
            
        # Format range string
        range_string = format_range_string(start_row, start_col, end_row, end_col)
        
        # Check if the range is actually merged
        merged_ranges = worksheet.merged_cells.ranges
        target_range = range_string.upper()
        
        if not any(str(merged_range).upper() == target_range for merged_range in merged_ranges):
            return {
                "success": False,
                "error": f"Range {range_string} is not merged"
            }
            
        # Unmerge cells
        worksheet.unmerge_cells(range_string)
        
        # Save workbook
        wb.save(filepath)
        wb.close()
        
        return {
            "success": True,
            "message": f"Range {range_string} unmerged in sheet '{sheet_name}'"
        }
    except Exception as e:
        logger.error(f"Failed to unmerge cells: {e}")
        return {
            "success": False,
            "error": f"Failed to unmerge cells: {str(e)}"
        }

# Data Operations
def write_excel_data(
    filepath: str,
    sheet_name: str,
    data: List[List[Any]],
    start_cell: str = "A1",
    headers: bool = True,
    auto_adjust_width: bool = False
) -> Dict[str, Any]:
    """Write data to an Excel worksheet"""
    try:
        # Validate input data
        if not data or not isinstance(data, list):
            return {
                "success": False,
                "error": "Data must be a non-empty list"
            }
            
        # Load workbook
        wb = load_workbook(filepath)
        
        # Validate sheet
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "error": f"Sheet '{sheet_name}' not found"
            }
            
        worksheet = wb[sheet_name]
        
        # Parse start cell
        try:
            start_row, start_col, _, _ = parse_cell_range(start_cell)
        except ValidationError as e:
            return {
                "success": False,
                "error": f"Invalid start cell: {str(e)}"
            }
            
        # Write data
        rows_written = 0
        max_col_width = {}  # Track max width for each column
        
        for row_idx, row_data in enumerate(data):
            if not isinstance(row_data, list):
                # Try to convert to list if possible (e.g., for dictionaries)
                if hasattr(row_data, 'values'):
                    row_data = list(row_data.values())
                else:
                    row_data = [row_data]  # Single value
                    
            for col_idx, value in enumerate(row_data):
                cell = worksheet.cell(
                    row=start_row + row_idx,
                    column=start_col + col_idx
                )
                cell.value = value
                
                # Add bold formatting for headers
                if headers and row_idx == 0:
                    cell.font = Font(bold=True)
                
                # Track max column width if auto-adjusting
                if auto_adjust_width:
                    # Calculate the display length of the value
                    str_value = str(value) if value is not None else ""
                    display_length = len(str_value)
                    
                    # Get column letter for this column
                    col_letter = get_column_letter(start_col + col_idx)
                    
                    # Update max width if needed
                    if col_letter not in max_col_width or display_length > max_col_width[col_letter]:
                        max_col_width[col_letter] = display_length
                    
            rows_written += 1
        
        # Adjust column widths if requested
        if auto_adjust_width and max_col_width:
            for col_letter, width in max_col_width.items():
                # Add padding and set a minimum width
                adjusted_width = min(max(width + 2, 8), 75)  # Min 8, max 75 characters
                worksheet.column_dimensions[col_letter].width = adjusted_width
        
        # Save workbook
        wb.save(filepath)
        wb.close()
        
        return {
            "success": True,
            "message": f"Data written to sheet '{sheet_name}' starting at {start_cell}",
            "rows_written": rows_written,
            "columns_written": len(data[0]) if data and data[0] else 0
        }
    except Exception as e:
        logger.error(f"Failed to write data: {e}")
        return {
            "success": False,
            "error": f"Failed to write data: {str(e)}"
        }

def format_excel_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: Optional[str] = None,
    bold: bool = False,
    italic: bool = False,
    font_size: Optional[int] = None,
    font_color: Optional[str] = None,
    bg_color: Optional[str] = None,
    alignment: Optional[str] = None,
    wrap_text: bool = False,
    border_style: Optional[str] = None,
    auto_adjust_width: bool = False
) -> Dict[str, Any]:
    """Apply formatting to a range of cells in Excel"""
    try:
        # Load workbook
        wb = load_workbook(filepath)
        
        # Validate sheet
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "error": f"Sheet '{sheet_name}' not found"
            }
            
        worksheet = wb[sheet_name]
        
        # Parse cell range
        try:
            start_row, start_col, end_row, end_col = parse_cell_range(start_cell, end_cell)
            
            if end_row is None:
                end_row = start_row
                end_col = start_col
                
        except ValidationError as e:
            return {
                "success": False,
                "error": f"Invalid range: {str(e)}"
            }
            
        # Set up font
        font_args = {}
        if bold:
            font_args["bold"] = True
        if italic:
            font_args["italic"] = True
        if font_size:
            font_args["size"] = font_size
        if font_color:
            # Handle color format (with or without #)
            if font_color.startswith("#"):
                font_color = font_color[1:]
            font_args["color"] = font_color
            
        # Set up fill
        fill = None
        if bg_color:
            # Handle color format (with or without #)
            if bg_color.startswith("#"):
                bg_color = bg_color[1:]
            fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
            
        # Set up alignment
        align = None
        if alignment or wrap_text:
            align_horz = None
            if alignment:
                align_horz = {
                    "left": "left",
                    "center": "center",
                    "right": "right"
                }.get(alignment.lower())
                
            align = Alignment(horizontal=align_horz, wrap_text=wrap_text)
            
        # Set up border
        border = None
        if border_style:
            border_styles = {
                "thin": Side(style="thin"),
                "medium": Side(style="medium"),
                "thick": Side(style="thick"),
                "dashed": Side(style="dashed"),
                "dotted": Side(style="dotted"),
                "double": Side(style="double")
            }
            if border_style.lower() in border_styles:
                side = border_styles[border_style.lower()]
                border = Border(left=side, right=side, top=side, bottom=side)
        
        # Track max column widths if auto-adjusting
        max_col_width = {} if auto_adjust_width else None
        
        # Apply formatting to range
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = worksheet.cell(row=row, column=col)
                
                # Apply font if any font properties set
                if font_args:
                    # Start with existing font and update properties
                    new_font = Font(
                        name=cell.font.name,
                        bold=cell.font.bold,
                        italic=cell.font.italic,
                        size=cell.font.size,
                        color=cell.font.color
                    )
                    # Update with specified properties
                    for key, value in font_args.items():
                        setattr(new_font, key, value)
                    cell.font = new_font
                
                # Apply fill if specified
                if fill:
                    cell.fill = fill
                    
                # Apply alignment if specified
                if align:
                    cell.alignment = align
                    
                # Apply border if specified
                if border:
                    cell.border = border
                
                # Track max column width if auto-adjusting
                if auto_adjust_width and cell.value is not None:
                    # Calculate display length
                    str_value = str(cell.value)
                    display_length = len(str_value)
                    
                    # Get column letter
                    col_letter = get_column_letter(col)
                    
                    # Update max width if needed
                    if col_letter not in max_col_width or display_length > max_col_width[col_letter]:
                        max_col_width[col_letter] = display_length
        
        # Adjust column widths if requested
        if auto_adjust_width and max_col_width:
            for col_letter, width in max_col_width.items():
                # Add padding based on content (more padding for wider content)
                padding = 2 if width < 20 else 4
                adjusted_width = min(width + padding, 75)  # Max 75 characters
                worksheet.column_dimensions[col_letter].width = adjusted_width
        
        # Save workbook
        wb.save(filepath)
        wb.close()
        
        # Format range string for the message
        range_string = format_range_string(start_row, start_col, end_row, end_col)
        
        return {
            "success": True,
            "message": f"Formatting applied to range {range_string} in sheet '{sheet_name}'"
        }
    except Exception as e:
        logger.error(f"Failed to apply formatting: {e}")
        return {
            "success": False,
            "error": f"Failed to apply formatting: {str(e)}"
        }

def adjust_column_widths(
    filepath: str,
    sheet_name: str,
    column_range: Optional[str] = None,
    auto_fit: bool = True,
    custom_widths: Optional[Dict[str, int]] = None
) -> Dict[str, Any]:
    """Adjust column widths in an Excel worksheet
    
    Args:
        filepath: Path to the Excel workbook
        sheet_name: Name of the worksheet to modify
        column_range: Optional range of columns to adjust (e.g. 'A:D')
        auto_fit: Whether to automatically fit column widths to content
        custom_widths: Optional dictionary mapping column letters to widths
    
    Returns:
        Dictionary with operation result
    """
    try:
        # Load workbook (can't use read_only mode when modifying)
        wb = load_workbook(filepath)
        
        # Validate sheet
        if sheet_name not in wb.sheetnames:
            return {
                "success": False,
                "error": f"Sheet '{sheet_name}' not found"
            }
            
        worksheet = wb[sheet_name]
        
        # Parse column range if provided
        start_col, end_col = 1, worksheet.max_column
        if column_range:
            try:
                if ':' in column_range:
                    start_col_str, end_col_str = column_range.split(':')
                    start_col = column_index_from_string(start_col_str)
                    end_col = column_index_from_string(end_col_str)
                else:
                    # Single column
                    start_col = end_col = column_index_from_string(column_range)
            except Exception as e:
                return {
                    "success": False,
                    "error": f"Invalid column range format: {str(e)}"
                }
        
        adjusted_columns = []
        
        # Apply custom widths if provided
        if custom_widths:
            for col_letter, width in custom_widths.items():
                try:
                    worksheet.column_dimensions[col_letter].width = width
                    adjusted_columns.append(col_letter)
                except Exception as e:
                    logger.warning(f"Failed to adjust column {col_letter}: {e}")
        
        # Auto-fit columns based on content
        if auto_fit:
            # For each column in range, find max content width
            for col_idx in range(start_col, end_col + 1):
                col_letter = get_column_letter(col_idx)
                max_length = 0
                
                # Check each cell in the column
                for row_idx in range(1, worksheet.max_row + 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    if cell.value:
                        # Get string representation and its length
                        cell_text = str(cell.value)
                        
                        # Adjust for bold text (takes more space)
                        if hasattr(cell, 'font') and cell.font and cell.font.bold:
                            length = len(cell_text) * 1.2  # Add 20% for bold
                        else:
                            length = len(cell_text)
                            
                        max_length = max(max_length, length)
                
                # Set column width (with some padding) if content exists
                if max_length > 0:
                    # Add padding based on content length
                    padding = 2 if max_length < 20 else 4
                    
                    # Set width (with minimum and maximum limits)
                    width = min(max(max_length + padding, 8), 75)  # Min 8, max 75 characters
                    worksheet.column_dimensions[col_letter].width = width
                    adjusted_columns.append(col_letter)
        
        # Save workbook
        wb.save(filepath)
        wb.close()
        
        return {
            "success": True,
            "message": f"Adjusted {len(adjusted_columns)} column widths in sheet '{sheet_name}'",
            "adjusted_columns": adjusted_columns
        }
    except Exception as e:
        logger.error(f"Failed to adjust column widths: {e}")
        return {
            "success": False,
            "error": f"Failed to adjust column widths: {str(e)}"
        } 