"""Office document reader module for Word, Excel and PowerPoint files."""
import os

# For image detection and properties
try:
    from PIL import Image
    has_pil = True
except ImportError:
    print("Pillow not installed. For better image analysis: pip install Pillow")
    has_pil = False

# For Microsoft Office documents
try:
    import docx  # for .docx files
    from docx.oxml.ns import qn  # for accessing embedded objects
    from docx.oxml import OxmlElement
except ImportError:
    print("python-docx not installed. To read Word files: pip install python-docx")

try:
    import openpyxl  # for .xlsx files
    from openpyxl.drawing.image import Image as XlsxImage
except ImportError:
    print("openpyxl not installed. To read Excel files: pip install openpyxl")

try:
    from pptx import Presentation  # for .pptx files
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    print("python-pptx not installed. To read PowerPoint files: pip install python-pptx")

# Function definitions will go here (added in separate edit)
def read_docx_file(file_path: str) -> str:
    """Extract text and identify objects from Microsoft Word (.docx) files in sequential order."""
    try:
        doc = docx.Document(file_path)
        full_text = []
        
        # Document properties first
        try:
            full_text.append("--- Document Properties ---")
            core_props = doc.core_properties
            if hasattr(core_props, 'title') and core_props.title:
                full_text.append(f"Title: {core_props.title}")
            if hasattr(core_props, 'author') and core_props.author:
                full_text.append(f"Author: {core_props.author}")
            if hasattr(core_props, 'created') and core_props.created:
                full_text.append(f"Created: {core_props.created}")
            if hasattr(core_props, 'modified') and core_props.modified:
                full_text.append(f"Modified: {core_props.modified}")
            full_text.append("-" * 40)
        except:
            pass  # Ignore if properties can't be accessed
        
        # Process paragraphs
        full_text.append("--- Document Content ---")
        for para in doc.paragraphs:
            if para.text.strip():  # Skip empty paragraphs
                full_text.append(para.text)
        
        # Process tables
        if doc.tables:
            full_text.append("\n--- Tables ---")
            for i, table in enumerate(doc.tables):
                full_text.append(f"\nTable {i+1}:")
                for row in table.rows:
                    row_text = []
                    for cell in row.cells:
                        # Join all paragraphs in the cell
                        cell_text = ' '.join([p.text for p in cell.paragraphs if p.text.strip()])
                        row_text.append(cell_text if cell_text else '')
                    full_text.append(" | ".join(row_text))
        
        return "\n".join(full_text)
    except Exception as e:
        return f"Error reading DOCX file: {str(e)}"

def read_xlsx_file(file_path: str) -> str:
    """Extract data from Microsoft Excel (.xlsx) files."""
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        content = []
        
        # Document properties
        content.append("--- Excel Document Properties ---")
        content.append(f"Filename: {os.path.basename(file_path)}")
        content.append(f"Number of sheets: {len(workbook.sheetnames)}")
        content.append(f"Sheet names: {', '.join(workbook.sheetnames)}")
        content.append("-" * 40)
        
        # Process each sheet
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            content.append(f"\n--- Sheet: {sheet_name} ---")
            content.append(f"Dimensions: {sheet.dimensions}")
            
            # Count non-empty cells
            non_empty = 0
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        non_empty += 1
            
            content.append(f"Non-empty cells: {non_empty}")
            
            # Check for charts, images, etc.
            if sheet._charts:
                content.append(f"Charts: {len(sheet._charts)}")
            
            if sheet._images:
                content.append(f"Images: {len(sheet._images)}")
            
            # Table data
            content.append("\n--- Data ---")
            
            # Get maximum row and column indices
            max_row = sheet.max_row
            max_col = sheet.max_column
            
            # Avoid excessive output for very large sheets
            display_rows = min(max_row, 100)  # Display up to 100 rows
            display_cols = min(max_col, 20)   # Display up to 20 columns
            
            # Add truncation notice if needed
            if display_rows < max_row or display_cols < max_col:
                content.append(f"Note: Displaying {display_rows}x{display_cols} of {max_row}x{max_col} cells")
            
            # Calculate column widths
            col_widths = []
            for col_idx in range(1, display_cols + 1):
                max_width = 0
                for row_idx in range(1, display_rows + 1):
                    cell = sheet.cell(row=row_idx, column=col_idx)
                    if cell.value is not None:
                        max_width = max(max_width, len(str(cell.value)))
                col_widths.append(max(max_width, 3))  # Minimum 3 chars width
            
            # Generate header row (column labels)
            header_row = []
            for col_idx in range(1, display_cols + 1):
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                header_row.append(col_letter.ljust(col_widths[col_idx-1]))
            content.append("| " + " | ".join(header_row) + " |")
            
            # Generate separator
            separator = []
            for width in col_widths:
                separator.append("-" * width)
            content.append("| " + " | ".join(separator) + " |")
            
            # Generate data rows
            for row_idx in range(1, display_rows + 1):
                row_data = []
                for col_idx in range(1, display_cols + 1):
                    cell = sheet.cell(row=row_idx, column=col_idx)
                    value = str(cell.value) if cell.value is not None else ""
                    # Truncate long values
                    if len(value) > 30:
                        value = value[:27] + "..."
                    row_data.append(value.ljust(col_widths[col_idx-1]))
                content.append("| " + " | ".join(row_data) + " |")
            
            # If there are more rows/columns, indicate truncation
            if max_row > display_rows:
                content.append(f"... {max_row - display_rows} more rows ...")
            if max_col > display_cols:
                content.append(f"... {max_col - display_cols} more columns ...")
        
        return "\n".join(content)
    except Exception as e:
        return f"Error reading Excel file: {str(e)}"

def read_pptx_file(file_path: str) -> str:
    """Extract content from Microsoft PowerPoint (.pptx) files."""
    try:
        presentation = Presentation(file_path)
        content = []
        
        # Document properties
        content.append("--- PowerPoint Document Properties ---")
        content.append(f"Filename: {os.path.basename(file_path)}")
        content.append(f"Number of slides: {len(presentation.slides)}")
        
        # Try to get core properties
        try:
            core_props = presentation.core_properties
            if hasattr(core_props, 'title') and core_props.title:
                content.append(f"Title: {core_props.title}")
            if hasattr(core_props, 'author') and core_props.author:
                content.append(f"Author: {core_props.author}")
            if hasattr(core_props, 'created') and core_props.created:
                content.append(f"Created: {core_props.created}")
            if hasattr(core_props, 'modified') and core_props.modified:
                content.append(f"Modified: {core_props.modified}")
        except:
            pass  # Ignore if properties can't be accessed
        
        content.append("-" * 40)
        
        # Process each slide
        for i, slide in enumerate(presentation.slides):
            content.append(f"\n--- Slide {i+1} ---")
            
            # Slide title
            if slide.shapes.title and slide.shapes.title.text:
                content.append(f"Title: {slide.shapes.title.text}")
            
            # Count different types of shapes
            shape_counts = {}
            for shape in slide.shapes:
                shape_type = str(shape.shape_type).replace('MSO_SHAPE_TYPE.', '')
                if shape_type not in shape_counts:
                    shape_counts[shape_type] = 0
                shape_counts[shape_type] += 1
            
            # Add shape count information
            content.append("Elements:")
            for shape_type, count in shape_counts.items():
                content.append(f"  {shape_type}: {count}")
            
            # Extract text from shapes
            texts = []
            for shape in slide.shapes:
                if hasattr(shape, 'text') and shape.text:
                    # Skip the title as we've already included it
                    if shape != slide.shapes.title:
                        texts.append(shape.text)
                
                # Check if this is a table
                if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                    table = shape.table
                    content.append("\nTable content:")
                    for row in table.rows:
                        row_texts = []
                        for cell in row.cells:
                            if cell.text:
                                row_texts.append(cell.text.replace('\n', ' '))
                            else:
                                row_texts.append("")
                        content.append("  " + " | ".join(row_texts))
            
            # Add extracted text
            if texts:
                content.append("\nText content:")
                for text in texts:
                    # Format multi-line text with indentation
                    for line in text.split('\n'):
                        if line.strip():
                            content.append(f"  {line}")
        
        return "\n".join(content)
    except Exception as e:
        return f"Error reading PowerPoint file: {str(e)}" 