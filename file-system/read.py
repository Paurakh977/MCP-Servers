import sys
import os
import csv
import io
import json
import time
from typing import Dict, Any, Union, List, Optional

# For PDF processing
try:
    import PyPDF2
except ImportError:
    print("PyPDF2 not installed. To read PDF files: pip install PyPDF2")

# For enhanced PDF processing with better image detection
try:
    import fitz  # PyMuPDF
    has_pymupdf = True
except ImportError:
    has_pymupdf = False
    print("PyMuPDF not installed. For better PDF image detection: pip install pymupdf")

# For EPUB e-books
try:
    import ebooklib
    from ebooklib import epub
    from bs4 import BeautifulSoup
    has_epub_support = True
except ImportError:
    has_epub_support = False
    print("EPUB support not available. To read EPUB files: pip install ebooklib beautifulsoup4")

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

# For image detection and properties
try:
    from PIL import Image
    has_pil = True
except ImportError:
    print("Pillow not installed. For better image analysis: pip install Pillow")
    has_pil = False

# For RTF documents
try:
    import striprtf.striprtf as striprtf
    has_rtf_support = True
except ImportError:
    has_rtf_support = False
    print("RTF support not available. To read RTF files: pip install striprtf")

def read_text_file(file_path: str) -> str:
    """Read content from text-based files with proper encoding detection."""
    encodings = ['utf-8', 'latin-1', 'windows-1252', 'ascii']
    
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                return f.read()
        except UnicodeDecodeError:
            continue

    # If all encodings fail, try binary mode as last resort
    try:
        with open(file_path, 'rb') as f:
            return f.read().decode('utf-8', errors='replace')
    except Exception as e:
        return f"Error: Could not decode file with available encodings: {str(e)}"

def read_pdf_file(file_path: str) -> str:
    """Extract text and identify images from PDF files."""
    # Try PyMuPDF first if available (better image handling)
    if has_pymupdf:
        try:
            content = []
            pdf_document = fitz.open(file_path)
            
            # Extract document metadata
            content.append("--- Document Metadata ---")
            metadata = pdf_document.metadata
            for key, value in metadata.items():
                if value:
                    content.append(f"{key}: {value}")
            content.append("-" * 40)
            
            # Document summary
            content.append(f"PDF Document: {os.path.basename(file_path)}")
            content.append(f"Number of pages: {len(pdf_document)}")
            content.append(f"PDF Version: {pdf_document.pdf_version}")
            if pdf_document.is_encrypted:
                content.append("Status: Encrypted")
            
            content.append("-" * 40)
            
            # Process each page
            for page_num, page in enumerate(pdf_document):
                content.append(f"--- Page {page_num + 1} ---")
                
                # Get text
                page_text = page.get_text()
                
                # Check for images
                image_list = page.get_images()
                if image_list:
                    content.append(f"[Contains {len(image_list)} image{'s' if len(image_list) > 1 else ''}]")
                    
                    # Try to get more details about images
                    for img_index, img in enumerate(image_list):
                        xref = img[0]  # image reference number
                        try:
                            base_image = pdf_document.extract_image(xref)
                            if base_image:
                                img_ext = base_image["ext"]     # image file extension
                                img_width = base_image.get("width", "unknown")
                                img_height = base_image.get("height", "unknown")
                                content.append(f"  Image {img_index+1}: {img_width}x{img_height} {img_ext}")
                        except:
                            # If detailed extraction fails, just note the presence
                            pass
                
                # Check for links
                links = page.get_links()
                if links:
                    content.append(f"[Contains {len(links)} link{'s' if len(links) > 1 else ''}]")
                
                # Check for annotations
                annots = page.annots()
                if annots:
                    content.append(f"[Contains {len(annots)} annotation{'s' if len(annots) > 1 else ''}]")
                
                # Handle empty pages
                if not page_text or page_text.isspace():
                    content.append("[This page appears to be empty or contains only non-text elements]")
                else:
                    content.append(page_text)
            
            return "\n\n".join(content)
        except Exception as e:
            # Fall back to PyPDF2 if PyMuPDF fails
            print(f"PyMuPDF processing failed: {str(e)}. Falling back to PyPDF2.")
    
    # Use PyPDF2 as fallback or if PyMuPDF is not available
    try:
        content = []
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            num_pages = len(pdf_reader.pages)
            
            # Try to get document info
            try:
                info = pdf_reader.metadata
                if info:
                    content.append("--- Document Metadata ---")
                    for key in info:
                        if info.get(key):
                            # Clean up key name by removing leading slash
                            clean_key = key[1:] if key.startswith('/') else key
                            content.append(f"{clean_key}: {info.get(key)}")
                    if content:  # Add separator if we added metadata
                        content.append("-" * 40)
            except:
                pass  # Ignore if metadata can't be accessed
            
            # Add document summary
            content.append(f"PDF Document: {os.path.basename(file_path)}")
            content.append(f"Number of pages: {num_pages}")
            content.append("-" * 40)
            
            # Extract text and detect resources from each page
            for page_num in range(num_pages):
                page = pdf_reader.pages[page_num]
                page_text = page.extract_text()
                
                # Start page content
                content.append(f"--- Page {page_num + 1} ---")
                
                # Try to get page resources to detect images and other objects
                try:
                    if '/Resources' in page:
                        resources = page['/Resources']
                        
                        # Check for XObject resources (images, forms, etc.)
                        if '/XObject' in resources:
                            xobjects = resources['/XObject']
                            if isinstance(xobjects, dict):
                                image_count = 0
                                form_count = 0
                                
                                for obj_name, obj_ref in xobjects.items():
                                    if '/Subtype' in obj_ref and obj_ref['/Subtype'] == '/Image':
                                        image_count += 1
                                    elif '/Subtype' in obj_ref and obj_ref['/Subtype'] == '/Form':
                                        form_count += 1
                                
                                if image_count > 0:
                                    content.append(f"[Contains {image_count} image{'s' if image_count > 1 else ''}]")
                                if form_count > 0:
                                    content.append(f"[Contains {form_count} form object{'s' if form_count > 1 else ''}]")
                except:
                    # If resource detection fails, just continue
                    pass
                
                # Handle empty pages
                if not page_text or page_text.isspace():
                    content.append("[This page appears to be empty or contains only non-text elements]")
                else:
                    content.append(page_text)
        
        return "\n\n".join(content)
    except Exception as e:
        return f"Error reading PDF file: {str(e)}"

def read_docx_file(file_path: str) -> str:
    """Extract text and identify objects from Microsoft Word (.docx) files."""
    try:
        doc = docx.Document(file_path)
        full_text = []
        
        # Document properties (enhanced)
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
            if hasattr(core_props, 'comments') and core_props.comments:
                full_text.append(f"Comments: {core_props.comments}")
            if hasattr(core_props, 'category') and core_props.category:
                full_text.append(f"Category: {core_props.category}")
            if hasattr(core_props, 'subject') and core_props.subject:
                full_text.append(f"Subject: {core_props.subject}")
            if hasattr(core_props, 'keywords') and core_props.keywords:
                full_text.append(f"Keywords: {core_props.keywords}")
            if full_text:  # Add a separator if we added properties
                full_text.append("-" * 40)
        except:
            pass  # Ignore if properties can't be accessed
        
        # Document statistics
        full_text.append("--- Document Statistics ---")
        full_text.append(f"Paragraphs: {len(doc.paragraphs)}")
        full_text.append(f"Sections: {len(doc.sections)}")
        full_text.append(f"Tables: {len(doc.tables)}")
        
        # Detect images and other embedded objects
        try:
            # Scan for embedded objects by looking at the underlying XML
            image_count = 0
            chart_count = 0
            shape_count = 0
            
            for rel in doc.part.rels.values():
                if rel.reltype == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image':
                    image_count += 1
                elif rel.reltype == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart':
                    chart_count += 1
                elif rel.reltype == 'http://schemas.microsoft.com/office/2007/relationships/shape':
                    shape_count += 1
            
            if image_count > 0:
                full_text.append(f"Images: {image_count}")
            if chart_count > 0:
                full_text.append(f"Charts: {chart_count}")
            if shape_count > 0:
                full_text.append(f"Shapes: {shape_count}")
            
            full_text.append("-" * 40)
        except:
            pass  # Skip if object detection fails
        
        # Track heading levels for structure
        current_heading = 0
        
        # Extract text from paragraphs with style information
        full_text.append("--- Document Content ---")
        for para in doc.paragraphs:
            if not para.text.strip():
                continue  # Skip empty paragraphs
                
            # Check if it's a heading and track heading level
            if para.style.name.startswith('Heading'):
                try:
                    heading_level = int(para.style.name.replace('Heading', ''))
                    current_heading = heading_level
                    # Format headings with appropriate markdown-style markers
                    full_text.append(f"{'#' * heading_level} {para.text}")
                except ValueError:
                    # If heading level can't be determined, just add the text
                    full_text.append(para.text)
            else:
                # For regular paragraphs
                text = para.text.strip()
                if text:
                    # Check for images or other objects within this paragraph
                    has_objects = False
                    for run in para.runs:
                        if run.element.findall('.//'+qn('w:drawing')) or run.element.findall('.//'+qn('w:pict')):
                            has_objects = True
                            break
                    
                    if has_objects:
                        full_text.append(f"{text} [Contains embedded object(s)]")
                    else:
                        full_text.append(text)
        
        # Add a separator before tables
        if doc.tables:
            full_text.append("\n" + "-" * 40 + "\nTABLES:\n" + "-" * 40)
            
        # Extract text from tables with better formatting
        for i, table in enumerate(doc.tables):
            full_text.append(f"\nTable {i+1}:")
            
            # Get all cell text first to determine column widths
            table_data = []
            max_widths = []
            
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    # Check cell for embedded objects
                    has_objects = False
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            if run.element.findall('.//'+qn('w:drawing')) or run.element.findall('.//'+qn('w:pict')):
                                has_objects = True
                                break
                    
                    cell_text = cell.text.strip().replace('\n', ' ')
                    if has_objects:
                        cell_text += " [Contains embedded object(s)]"
                    row_data.append(cell_text)
                
                # Update max column widths
                while len(max_widths) < len(row_data):
                    max_widths.append(0)
                
                for i, cell_text in enumerate(row_data):
                    max_widths[i] = max(max_widths[i], len(cell_text))
                
                table_data.append(row_data)
            
            # Format table with proper alignment
            for row_data in table_data:
                formatted_row = " | ".join([cell_text for cell_text in row_data])
                full_text.append(formatted_row)
        
        return "\n".join(full_text)
    except Exception as e:
        return f"Error reading DOCX file: {str(e)}"

def read_xlsx_file(file_path: str) -> str:
    """Extract data and identify objects from Microsoft Excel (.xlsx) files."""
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True)  # Use read_only=False to access images
        result = []
        
        # Workbook properties
        result.append("--- Workbook Properties ---")
        result.append(f"Filename: {os.path.basename(file_path)}")
        result.append(f"Number of sheets: {len(workbook.sheetnames)}")
        
        # Try to get document properties
        try:
            props = workbook.properties
            if props.title:
                result.append(f"Title: {props.title}")
            if props.subject:
                result.append(f"Subject: {props.subject}")
            if props.creator:
                result.append(f"Creator: {props.creator}")
            if props.created:
                result.append(f"Created: {props.created}")
        except:
            pass  # Skip if properties aren't accessible
            
        result.append("-" * 40)
        
        # Process each worksheet
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            result.append(f"--- Sheet: {sheet_name} ---")
            
            # Get sheet dimensions and properties
            dim = sheet.calculate_dimension()
            result.append(f"Dimension: {dim}")
            
            # Find charts
            chart_count = 0
            for chart_ref in sheet._charts:
                chart_count += 1
            
            if chart_count > 0:
                result.append(f"Charts: {chart_count}")
            
            # Find images
            image_count = 0
            try:
                for image in sheet._images:
                    image_count += 1
            except:
                pass  # Skip if images can't be accessed
                
            if image_count > 0:
                result.append(f"Images: {image_count}")
                
            # Find merged cells
            if sheet.merged_cells:
                result.append(f"Merged cell ranges: {len(sheet.merged_cells.ranges)}")
            
            # Find conditional formatting
            if hasattr(sheet, 'conditional_formatting') and sheet.conditional_formatting:
                result.append(f"Conditional formatting rules: {len(sheet.conditional_formatting._cf_rules)}")
            
            result.append("-" * 40)
            
            # Track non-empty cells to avoid excessive empty data
            rows_processed = 0
            empty_row_count = 0
            max_empty_rows = 10  # Stop after this many consecutive empty rows
            
            # Process rows
            for row_cells in sheet.iter_rows(values_only=True):
                # Check if row is empty
                if all(cell is None or str(cell).strip() == "" for cell in row_cells):
                    empty_row_count += 1
                    if empty_row_count > max_empty_rows and rows_processed > 10:
                        # If we've seen many rows and hit many empty ones, assume end of data
                        break
                else:
                    empty_row_count = 0  # Reset empty row counter
                
                # Format row data
                row_data = []
                for cell_value in row_cells:
                    if cell_value is None:
                        cell_value = ""
                    row_data.append(str(cell_value))
                
                result.append("\t".join(row_data))
                rows_processed += 1
                
                # Limit to 1000 rows for performance
                if rows_processed >= 1000:
                    result.append("... (truncated due to large size)")
                    break
        
        return "\n".join(result)
    except Exception as e:
        return f"Error reading XLSX file: {str(e)}"

def read_pptx_file(file_path: str) -> str:
    """Extract text and identify objects from Microsoft PowerPoint (.pptx) files."""
    try:
        prs = Presentation(file_path)
        full_text = []
        
        # Document properties
        full_text.append("--- Presentation Properties ---")
        full_text.append(f"Filename: {os.path.basename(file_path)}")
        full_text.append(f"Number of slides: {len(prs.slides)}")
        
        # Try to get core properties
        try:
            core_props = prs.core_properties
            if hasattr(core_props, 'title') and core_props.title:
                full_text.append(f"Title: {core_props.title}")
            if hasattr(core_props, 'author') and core_props.author:
                full_text.append(f"Author: {core_props.author}")
            if hasattr(core_props, 'subject') and core_props.subject:
                full_text.append(f"Subject: {core_props.subject}")
            if hasattr(core_props, 'keywords') and core_props.keywords:
                full_text.append(f"Keywords: {core_props.keywords}")
        except:
            pass  # Skip if properties can't be accessed
        
        full_text.append("-" * 40)
        
        # Process each slide
        for i, slide in enumerate(prs.slides):
            slide_content = [f"--- Slide {i+1} ---"]
            
            # Extract text from slide title
            if slide.shapes.title:
                slide_content.append(f"Title: {slide.shapes.title.text}")
            
            # Shape and object summary
            shape_counts = {
                'text_boxes': 0,
                'pictures': 0,
                'charts': 0,
                'tables': 0,
                'diagrams': 0,
                'videos': 0,
                'other_shapes': 0
            }
            
            # Extract text and count objects
            text_shapes = []
            
            # Process all shapes
            for shape in slide.shapes:
                # Count by shape type
                if shape.has_text_frame:
                    if shape != slide.shapes.title and shape.text.strip():
                        shape_counts['text_boxes'] += 1
                        text_shapes.append(shape.text.strip())
                elif shape.has_table:
                    shape_counts['tables'] += 1
                elif shape.has_chart:
                    shape_counts['charts'] += 1
                elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    shape_counts['pictures'] += 1
                elif shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
                    shape_counts['videos'] += 1
                elif shape.shape_type == MSO_SHAPE_TYPE.DIAGRAM:
                    shape_counts['diagrams'] += 1
                else:
                    shape_counts['other_shapes'] += 1
            
            # Add shape summary
            shape_summary = []
            for shape_type, count in shape_counts.items():
                if count > 0:
                    shape_summary.append(f"{count} {shape_type.replace('_', ' ')}")
            
            if shape_summary:
                slide_content.append("Objects: " + ", ".join(shape_summary))
            
            # Add text content
            if text_shapes:
                slide_content.append("Text Content:")
                for text in text_shapes:
                    slide_content.append(f"  • {text}")
            
            # Process tables if present
            tables_processed = 0
            for shape in slide.shapes:
                if shape.has_table:
                    tables_processed += 1
                    table = shape.table
                    slide_content.append(f"\nTable {tables_processed}:")
                    
                    # Process each row in the table
                    for r, row in enumerate(table.rows):
                        row_text = []
                        for c, cell in enumerate(row.cells):
                            if cell.text:
                                row_text.append(cell.text.strip().replace('\n', ' '))
                            else:
                                row_text.append("")
                        
                        slide_content.append(" | ".join(row_text))
            
            # Check if slide has minimal content
            if len(slide_content) <= 2:  # Only slide number and maybe title
                slide_content.append("[This slide appears to have no text content or contains only non-text elements]")
            
            full_text.append("\n".join(slide_content))
        
        return "\n\n".join(full_text)
    except Exception as e:
        return f"Error reading PPTX file: {str(e)}"

def read_csv_file(file_path: str) -> str:
    """Read CSV files with enhanced delimiter detection and rich formatting."""
    try:
        # Try multiple encodings
        encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'windows-1252']
        
        for encoding in encodings:
            try:
                # Try to detect delimiter and file structure
                with open(file_path, 'r', newline='', encoding=encoding) as csvfile:
                    # Read a sample to analyze
                    sample = csvfile.read(8192)  # Read larger sample for better detection
                    csvfile.seek(0)  # Reset to beginning of file
                    
                    # Analyze file structure
                    result = []
                    result.append(f"--- CSV File Analysis: {os.path.basename(file_path)} ---")
                    
                    # Count lines in the file
                    line_count = sample.count('\n')
                    if csvfile.read() != '':  # If there's more content after the sample
                        line_count = line_count + 1 + '...'
                    csvfile.seek(0)  # Reset again
                    
                    result.append(f"Estimated number of rows: {line_count if isinstance(line_count, int) else '1000+'}")
                    
                    # Try to detect the dialect
                    try:
                        dialect = csv.Sniffer().sniff(sample)
                        delimiter = dialect.delimiter
                        result.append(f"Detected delimiter: '{delimiter}'")
                        result.append(f"Quote character: '{dialect.quotechar}'")
                        
                        # Check if file has a header
                        has_header = csv.Sniffer().has_header(sample)
                        result.append(f"Has header row: {has_header}")
                    except:
                        # If sniffing fails, try common delimiters
                        for delimiter in [',', ';', '\t', '|']:
                            if delimiter in sample:
                                result.append(f"Using delimiter: '{delimiter}'")
                                break
                        has_header = True  # Assume header by default
                    
                    # Add separator
                    result.append("-" * 40)
                    
                    # Read and format CSV data
                    reader = csv.reader(csvfile, delimiter=delimiter)
                    rows = []
                    
                    # Read header if present
                    header = next(reader) if has_header else None
                    if header:
                        result.append(f"Column count: {len(header)}")
                        result.append(f"Headers: {', '.join(header)}")
                        result.append("-" * 40)
                        rows.append("\t".join(header))
                    
                    # Process data rows
                    row_count = 0
                    max_rows = 2000  # Reasonable limit for large files
                    column_stats = {}  # Track stats about each column
                    
                    for row in reader:
                        rows.append("\t".join(row))
                        
                        # Gather column statistics
                        for i, value in enumerate(row):
                            if i not in column_stats:
                                column_stats[i] = {'numeric': 0, 'empty': 0, 'total': 0}
                            
                            column_stats[i]['total'] += 1
                            if not value.strip():
                                column_stats[i]['empty'] += 1
                            elif value.strip().replace('.', '', 1).replace('-', '', 1).isdigit():
                                column_stats[i]['numeric'] += 1
                        
                        row_count += 1
                        if row_count >= max_rows:
                            rows.append("... (truncated due to large size)")
                            break
                    
                    # Add column analysis if we have enough data
                    if row_count > 10 and header:
                        result.append("\nColumn Analysis:")
                        for i, col_name in enumerate(header):
                            if i in column_stats and column_stats[i]['total'] > 0:
                                stats = column_stats[i]
                                pct_numeric = (stats['numeric'] / stats['total']) * 100
                                pct_empty = (stats['empty'] / stats['total']) * 100
                                
                                col_type = 'Numeric' if pct_numeric > 90 else 'Text'
                                if pct_empty > 50:
                                    col_type += " (Sparse)"
                                
                                result.append(f"  {col_name}: {col_type}")
                        
                        result.append("-" * 40)
                    
                    # Add the actual data
                    result.extend(rows)
                    
                    return "\n".join(result)
            except UnicodeDecodeError:
                continue  # Try next encoding
        
        # If all encodings fail, try as plain text
        return read_text_file(file_path)
    except Exception as e:
        # Fall back to simple text reading if CSV parsing fails
        try:
            return read_text_file(file_path)
        except Exception as nested_e:
            return f"Error reading CSV file: {str(e)}, Nested error: {str(nested_e)}"

def read_epub_file(file_path: str) -> str:
    """Extract content from EPUB e-books."""
    if not has_epub_support:
        return "EPUB support not available. Install required packages with: pip install ebooklib beautifulsoup4"
    
    try:
        content = []
        book = epub.read_epub(file_path)
        
        # Extract metadata
        content.append("--- EPUB Metadata ---")
        content.append(f"Title: {book.get_metadata('DC', 'title')[0][0] if book.get_metadata('DC', 'title') else 'Unknown'}")
        
        # Get author(s)
        authors = book.get_metadata('DC', 'creator')
        if authors:
            author_list = [author[0] for author in authors]
            content.append(f"Author(s): {', '.join(author_list)}")
        
        # Get language
        languages = book.get_metadata('DC', 'language')
        if languages:
            content.append(f"Language: {languages[0][0]}")
            
        # Get other metadata
        identifiers = book.get_metadata('DC', 'identifier')
        if identifiers:
            for identifier in identifiers:
                if identifier[1].get('id') == 'ISBN':
                    content.append(f"ISBN: {identifier[0]}")
        
        publishers = book.get_metadata('DC', 'publisher')
        if publishers:
            content.append(f"Publisher: {publishers[0][0]}")
            
        dates = book.get_metadata('DC', 'date')
        if dates:
            content.append(f"Date: {dates[0][0]}")
            
        content.append("-" * 40)
        
        # Get table of contents
        toc = book.toc
        if toc:
            content.append("--- Table of Contents ---")
            
            def process_toc_entries(entries, level=0):
                toc_content = []
                for entry in entries:
                    if isinstance(entry, tuple) and len(entry) >= 2:
                        title, href = entry[0], entry[1]
                        toc_content.append(f"{'  ' * level}• {title}")
                    elif isinstance(entry, list):
                        toc_content.extend(process_toc_entries(entry, level + 1))
                return toc_content
            
            content.extend(process_toc_entries(toc))
            content.append("-" * 40)
        
        # Count items and get document statistics
        content.append("--- Document Statistics ---")
        content.append(f"Spine items: {len(book.spine)}")
        content.append(f"Total items: {len(book.items)}")
        
        # Count images
        image_count = 0
        css_count = 0
        html_count = 0
        
        for item in book.items:
            if item.media_type and item.media_type.startswith('image/'):
                image_count += 1
            elif item.media_type == 'text/css':
                css_count += 1
            elif item.media_type == 'application/xhtml+xml':
                html_count += 1
        
        content.append(f"HTML documents: {html_count}")
        content.append(f"CSS stylesheets: {css_count}")
        content.append(f"Images: {image_count}")
        content.append("-" * 40)
        
        # Extract text content from HTML documents
        content.append("--- Content ---")
        
        # Helper function to extract text from HTML
        def chapter_to_text(html_content):
            try:
                soup = BeautifulSoup(html_content, 'html.parser')
                
                # Extract title if available
                title = soup.find('title')
                title_text = f"Chapter: {title.text}\n" if title else ""
                
                # Remove script and style elements
                for script in soup(["script", "style"]):
                    script.extract()
                
                # Get text
                text = soup.get_text(separator='\n')
                
                # Clean whitespace
                lines = (line.strip() for line in text.splitlines())
                chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
                text = '\n'.join(chunk for chunk in chunks if chunk)
                
                return title_text + text
            except Exception as e:
                return f"[Error processing HTML content: {str(e)}]"
        
        # Process spine documents in order
        processed_items = set()
        for item_id in book.spine:
            item = book.get_item_with_id(item_id[0] if isinstance(item_id, tuple) else item_id)
            if item and item.get_content() and item.media_type == 'application/xhtml+xml':
                processed_items.add(item.id)
                chapter_text = chapter_to_text(item.get_content().decode('utf-8'))
                if chapter_text:
                    content.append(f"--- Document: {item.get_name()} ---")
                    content.append(chapter_text)
                    content.append("-" * 40)
        
        # Process any HTML items not in spine but might contain important content
        for item in book.items:
            if (item.id not in processed_items and 
                item.media_type == 'application/xhtml+xml' and 
                not item.get_name().startswith('nav')):  # Skip navigation files
                chapter_text = chapter_to_text(item.get_content().decode('utf-8'))
                if chapter_text:
                    content.append(f"--- Additional Document: {item.get_name()} ---")
                    content.append(chapter_text)
                    content.append("-" * 40)
        
        return "\n".join(content)
    except Exception as e:
        return f"Error reading EPUB file: {str(e)}"

def read_rtf_file(file_path: str) -> str:
    """Extract text content from RTF files."""
    if not has_rtf_support:
        return "RTF support not available. Install required package with: pip install striprtf"
    
    try:
        # Read RTF content
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            rtf_text = file.read()
        
        # Convert RTF to plain text
        plain_text = striprtf.rtf_to_text(rtf_text)
        
        # Format output
        content = []
        content.append(f"--- RTF Document: {os.path.basename(file_path)} ---")
        content.append(f"Size: {os.path.getsize(file_path) / 1024:.2f} KB")
        content.append("-" * 40)
        content.append(plain_text)
        
        return "\n".join(content)
    except Exception as e:
        return f"Error reading RTF file: {str(e)}"

def extract_file_content(file_path: str) -> Dict[str, Any]:
    """Extract content from a file based on its extension."""
    if not os.path.exists(file_path):
        return {"success": False, "error": f"File not found: {file_path}", "content": ""}
    
    file_extension = os.path.splitext(file_path)[1].lower()
    content = ""
    
    try:
        # Handle different file types
        if file_extension in ['.txt', '.md', '.json', '.html', '.xml', '.log', '.py', '.js', '.css', '.java', '.ini', '.conf', '.cfg']:
            content = read_text_file(file_path)
        elif file_extension == '.pdf':
            content = read_pdf_file(file_path)
        elif file_extension == '.docx':
            content = read_docx_file(file_path)
        elif file_extension == '.xlsx':
            content = read_xlsx_file(file_path)
        elif file_extension == '.pptx':
            content = read_pptx_file(file_path)
        elif file_extension == '.csv':
            content = read_csv_file(file_path)
        elif file_extension == '.epub':
            content = read_epub_file(file_path)
        elif file_extension == '.rtf':
            content = read_rtf_file(file_path)
        else:
            # Try to read as text file first, then fall back to binary warning
            try:
                content = read_text_file(file_path)
            except Exception:
                return {
                    "success": False, 
                    "error": f"Unsupported file type: {file_extension}", 
                    "content": f"This file type ({file_extension}) is not directly supported."
                }
        
        return {"success": True, "content": content, "file_path": file_path, "file_type": file_extension}
    
    except Exception as e:
        return {"success": False, "error": str(e), "content": ""}

def summarize_content(content: str, max_length: int = 500) -> str:
    """Create a brief summary of the content if it's too long."""
    if len(content) <= max_length:
        return content
    
    # Simple summarization: first few characters and last few characters
    first_part = content[:max_length // 2]
    last_part = content[-(max_length // 2):]
    return f"{first_part}\n\n... [Content truncated, total length: {len(content)} characters] ...\n\n{last_part}"

def read_file(path: str, summarize: bool = False, max_summary_length: int = 500) -> Dict[str, Any]:
    """
    Reads and returns the contents of the file at 'path' with appropriate handling per file type.
    
    Args:
        path (str): Path to the file to read
        summarize (bool): Whether to summarize very large content
        max_summary_length (int): Maximum length for summary if summarizing
        
    Returns:
        Dict with keys:
        - success (bool): Whether the read was successful
        - content (str): The file content, possibly summarized
        - file_path (str): Original file path
        - file_type (str): File extension
        - error (str, optional): Error message if success is False
    """
    result = extract_file_content(path)
    
    # Summarize very large content if requested
    if summarize and result["success"] and len(result["content"]) > max_summary_length:
        result["content"] = summarize_content(result["content"], max_summary_length)
        result["summarized"] = True
    
    return result

def print_output(result: Dict[str, Any], output_format: str = "text") -> None:
    """Print the output in the specified format."""
    if output_format == "json":
        print(json.dumps(result, indent=2))
    else:
        if result["success"]:
            print(f"File: {result['file_path']} ({result['file_type']})")
            print("-" * 40)
            print(result["content"])
        else:
            print(f"Error: {result['error']}")

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

def interactive_mode():
    """Run the file reader in interactive mode."""
    print("=" * 60)
    print("       Advanced File Content Extractor")
    print("=" * 60)
    print("Supported file types:")
    print("  • Text files (.txt, .md, .log, etc.)")
    print("  • PDF documents (.pdf)")
    print("  • Microsoft Word documents (.docx)")
    print("  • Microsoft Excel spreadsheets (.xlsx)")
    print("  • Microsoft PowerPoint presentations (.pptx)")
    print("  • CSV files (.csv)")
    print("  • EPUB e-books (.epub)")
    print("  • RTF documents (.rtf)")
    
    # Show which enhanced modules are available
    print("\nEnhanced features available:")
    if has_pymupdf:
        print("  ✓ Advanced PDF processing (PyMuPDF)")
    else:
        print("  ✗ Advanced PDF processing (install pymupdf for better results)")
    
    if has_epub_support:
        print("  ✓ EPUB e-book support")
    else:
        print("  ✗ EPUB support (install ebooklib and beautifulsoup4)")
    
    if has_rtf_support:
        print("  ✓ RTF document support")
    else:
        print("  ✗ RTF support (install striprtf)")
    
    if has_pil:
        print("  ✓ Enhanced image analysis")
    else:
        print("  ✗ Enhanced image analysis (install Pillow)")
    
    print("=" * 60)
    
    # Ask for file path
    file_path = input("Enter the path to the file you want to read: ").strip()
    
    # Check if file exists
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' does not exist!")
        return
    
    # Options
    print("\nOutput options:")
    print("1. Display in terminal")
    print("2. Save to text file")
    print("3. Both display and save")
    
    while True:
        try:
            output_choice = int(input("Enter your choice (1-3): "))
            if 1 <= output_choice <= 3:
                break
            else:
                print("Please enter a number between 1 and 3.")
        except ValueError:
            print("Please enter a valid number.")
    
    # Summarize option for large files
    summarize = input("\nSummarize large content? (y/n): ").lower().startswith('y')
    max_length = 1000
    if summarize:
        try:
            max_length = int(input("Maximum length for summary (default: 1000): ") or "1000")
        except ValueError:
            print("Using default length of 1000 characters.")
    
    # Process file
    print(f"\nProcessing {file_path}...")
    start_time = time.time()
    result = read_file(file_path, summarize, max_length)
    processing_time = time.time() - start_time
    
    # Output
    if not result["success"]:
        print(f"\nError: {result['error']}")
        return
    
    file_size = os.path.getsize(file_path) / 1024  # KB
    print(f"\nFile processed successfully in {processing_time:.2f} seconds.")
    print(f"File size: {file_size:.2f} KB")
    print(f"File type: {result['file_type']}")
    content_length = len(result["content"])
    print(f"Extracted content length: {content_length} characters")
    
    # Display content
    if output_choice in [1, 3]:
        print("\n" + "=" * 60)
        print(f"CONTENT OF {os.path.basename(file_path)}:")
        print("=" * 60)
        print(result["content"])
        print("=" * 60)
    
    # Save to file
    if output_choice in [2, 3]:
        output_path = save_to_file(result["content"], file_path)
        if output_path:
            print(f"\nContent saved to: {output_path}")
    
    print("\nDone!")

if __name__ == '__main__':
    # Check if arguments were provided
    if len(sys.argv) > 1:
        # Command line mode
        import argparse
        parser = argparse.ArgumentParser(description='Read and extract content from various file types.')
        parser.add_argument('file_path', help='Path to the file to read')
        parser.add_argument('--format', choices=['text', 'json'], default='text', help='Output format')
        parser.add_argument('--summarize', action='store_true', help='Summarize large content')
        parser.add_argument('--max-length', type=int, default=1000, help='Maximum length for summary')
        parser.add_argument('--save', action='store_true', help='Save output to text file')
        
        args = parser.parse_args()
        
        # Process the file
        result = read_file(args.file_path, args.summarize, args.max_length)
        print_output(result, args.format)
        
        # Save to file if requested
        if args.save and result["success"]:
            output_path = save_to_file(result["content"], args.file_path)
            if output_path:
                print(f"Content saved to: {output_path}")
    else:
        # Interactive mode
        interactive_mode()