"""PDF reader module with enhanced extraction capabilities."""

import os
from ..core.dependencies import (
    has_pdfplumber, has_pymupdf, has_tabula,
    PyPDF2, pdfplumber, fitz, tabula
)

def read_pdf_file(file_path: str) -> str:
    """Extract text and identify images from PDF files with layout preservation."""
    content = []
    
    # Try pdfplumber first if available
    if has_pdfplumber:
        try:
            content.append("--- Document Metadata ---")
            pdf = pdfplumber.open(file_path)
            
            # Basic document info
            content.append(f"PDF Document: {os.path.basename(file_path)}")
            content.append(f"Number of pages: {len(pdf.pages)}")
            
            # Try to extract metadata
            if hasattr(pdf, 'metadata') and pdf.metadata:
                for key, value in pdf.metadata.items():
                    if value and str(value).strip():
                        # Clean up key name
                        clean_key = key[1:] if isinstance(key, str) and key.startswith('/') else key
                        content.append(f"{clean_key}: {value}")
            
            content.append("-" * 40)
            
            # Process each page
            for i, page in enumerate(pdf.pages):
                content.append(f"--- Page {i + 1} ---")
                
                # Extract page dimensions
                width, height = page.width, page.height
                content.append(f"Page dimensions: {width:.2f} x {height:.2f} points")
                
                # Extract text with layout preservation
                page_text = page.extract_text(x_tolerance=3, y_tolerance=3)
                
                # Try to detect tables
                tables = page.extract_tables()
                if tables:
                    content.append(f"[Contains {len(tables)} table{'s' if len(tables) > 1 else ''}]")
                    
                    # Format each table
                    for t_idx, table in enumerate(tables):
                        content.append(f"\nTable {t_idx + 1}:")
                        
                        # Format the table with proper alignment
                        for row in table:
                            # Clean row data and handle None values
                            cleaned_row = [str(cell).strip() if cell is not None else "" for cell in row]
                            row_text = " | ".join(cleaned_row)
                            content.append(row_text)
                        
                        content.append("")  # Add spacing after table
                
                # Extract images if available
                try:
                    images = page.images
                    if images:
                        content.append(f"[Contains {len(images)} image{'s' if len(images) > 1 else ''}]")
                        for img_idx, img in enumerate(images):
                            if 'width' in img and 'height' in img:
                                content.append(f"  Image {img_idx+1}: {img.get('width', '?')}x{img.get('height', '?')} pixels")
                            else:
                                content.append(f"  Image {img_idx+1}")
                except Exception as e:
                    pass  # Skip if image extraction fails
                
                # Add the page text with layout preservation
                if not page_text or page_text.isspace():
                    content.append("[This page appears to be empty or contains only non-text elements]")
                else:
                    content.append(page_text)
            
            pdf.close()
            return "\n\n".join(content)
        except Exception as e:
            print(f"pdfplumber processing failed: {str(e)}. Trying PyMuPDF.")
    
    # Try PyMuPDF if available
    if has_pymupdf:
        try:
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
            
            # Try using tabula for table extraction if available
            tables_by_page = {}
            if has_tabula:
                try:
                    # Extract tables from all pages
                    all_tables = tabula.read_pdf(file_path, pages='all', multiple_tables=True)
                    
                    # Organize tables by page
                    for idx, table in enumerate(all_tables):
                        # Tabula doesn't provide page numbers directly, so we estimate
                        # This is a limitation, might need adjustment based on PDF structure
                        page_num = idx % len(pdf_document) + 1  # Basic estimation
                        if page_num not in tables_by_page:
                            tables_by_page[page_num] = []
                        tables_by_page[page_num].append(table)
                except Exception as e:
                    print(f"Tabula table extraction failed: {str(e)}")
            
            # Process each page
            for page_num, page in enumerate(pdf_document):
                content.append(f"--- Page {page_num + 1} ---")
                
                # Get page size
                width, height = page.rect.width, page.rect.height
                content.append(f"Page dimensions: {width:.2f} x {height:.2f} points")
                
                # We'll extract content in blocks to preserve layout and sequence
                # Get the page elements in order
                page_elements = []
                
                # First, extract text blocks with position information
                # Sort by y position first (top to bottom), then x position (left to right)
                blocks = page.get_text("dict", sort=True)["blocks"]
                
                # Track text blocks with their position
                text_blocks = []
                for block in blocks:
                    if block["type"] == 0:  # Text block
                        text = ""
                        for line in block["lines"]:
                            for span in line["spans"]:
                                text += span["text"]
                            text += "\n"
                        
                        # Store the text block with its position
                        if text.strip():
                            y_pos = block["bbox"][1]  # Top y-coordinate
                            x_pos = block["bbox"][0]  # Left x-coordinate
                            text_blocks.append({
                                "type": "text",
                                "y_pos": y_pos,
                                "x_pos": x_pos,
                                "text": text.strip()
                            })
                
                # Get image information with position
                image_blocks = []
                for img_index, img in enumerate(page.get_images(full=True)):
                    xref = img[0]  # image reference number
                    
                    # Get image details
                    try:
                        base_image = pdf_document.extract_image(xref)
                        if base_image:
                            # Try to get image position on page
                            rect = None
                            for item in page.get_drawings():
                                if item["type"] == "image" and xref == item.get("xref"):
                                    rect = item["rect"]
                                    break
                            
                            if rect:
                                y_pos = rect[1]  # Top y-coordinate
                                x_pos = rect[0]  # Left x-coordinate
                                img_ext = base_image.get("ext", "")
                                img_width = base_image.get("width", "unknown")
                                img_height = base_image.get("height", "unknown")
                                
                                image_blocks.append({
                                    "type": "image",
                                    "y_pos": y_pos,
                                    "x_pos": x_pos,
                                    "desc": f"Image {img_index+1}: {img_width}x{img_height} {img_ext}"
                                })
                    except:
                        # If we can't get position, add it to the end
                        image_blocks.append({
                            "type": "image",
                            "y_pos": height,  # Put at bottom if position unknown
                            "x_pos": 0,
                            "desc": f"Image {img_index+1}"
                        })
                
                # Merge all blocks and sort by position
                page_elements = text_blocks + image_blocks
                page_elements.sort(key=lambda x: (x["y_pos"], x["x_pos"]))
                
                # Check for tables from tabula
                table_blocks = []
                if page_num + 1 in tables_by_page:
                    tables = tables_by_page[page_num + 1]
                    for t_idx, table in enumerate(tables):
                        # We don't have position info from tabula, so estimate
                        # based on where tables typically appear in the document
                        # This is approximate and may not perfectly match the layout
                        est_y_pos = height * 0.3 * (t_idx + 1)  # Estimate
                        
                        table_blocks.append({
                            "type": "table",
                            "y_pos": est_y_pos,
                            "x_pos": width / 4,  # Center-left of page
                            "table": table
                        })
                
                # Add tables to page elements if we have any
                if table_blocks:
                    page_elements.extend(table_blocks)
                    # Re-sort with tables included
                    page_elements.sort(key=lambda x: (x["y_pos"], x["x_pos"]))
                
                # If we have no elements, check if page has text using plain extraction
                if not page_elements:
                    page_text = page.get_text()
                    if page_text.strip():
                        content.append(page_text.strip())
                    else:
                        content.append("[This page appears to be empty or contains only non-text elements]")
                else:
                    # Output elements in order
                    for element in page_elements:
                        if element["type"] == "text":
                            content.append(element["text"])
                        elif element["type"] == "image":
                            content.append(f"[{element['desc']}]")
                        elif element["type"] == "table":
                            content.append("\n--- Table ---")
                            table_str = element["table"].to_string(index=False)
                            content.append(table_str)
                            content.append("--- End Table ---")
                
                # Check for links and annotations at the end
                links = page.get_links()
                if links:
                    content.append(f"[Contains {len(links)} link{'s' if len(links) > 1 else ''}]")
                
                annots = page.annots()
                if annots:
                    content.append(f"[Contains {len(annots)} annotation{'s' if len(annots) > 1 else ''}]")
            
            return "\n\n".join(content)
        except Exception as e:
            print(f"PyMuPDF processing failed: {str(e)}. Falling back to PyPDF2.")
    
    # Use PyPDF2 as final fallback
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            num_pages = len(pdf_reader.pages)
            
            # Try to get document info
            content.append("--- Document Metadata ---")
            info = pdf_reader.metadata
            if info:
                for key in info:
                    if info.get(key):
                        # Clean up key name by removing leading slash
                        clean_key = key[1:] if key.startswith('/') else key
                        content.append(f"{clean_key}: {info.get(key)}")
            
            # Add document summary
            content.append(f"PDF Document: {os.path.basename(file_path)}")
            content.append(f"Number of pages: {num_pages}")
            content.append("-" * 40)
            
            # Extract text from each page
            for page_num in range(num_pages):
                page = pdf_reader.pages[page_num]
                
                # Start page content
                content.append(f"--- Page {page_num + 1} ---")
                
                # Extract text with better layout handling
                try:
                    page_text = page.extract_text(extraction_mode="layout")
                except:
                    # Fallback to regular extraction if layout mode is not available
                    page_text = page.extract_text()
                
                # Try to detect images and objects
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
                    # Format text with proper paragraph separation
                    paragraphs = page_text.split('\n\n')
                    formatted_text = []
                    for para in paragraphs:
                        # Ensure there are no artificial line breaks within paragraphs
                        # but preserve real paragraph breaks
                        lines = para.split('\n')
                        formatted_para = ' '.join(line.strip() for line in lines if line.strip())
                        if formatted_para:
                            formatted_text.append(formatted_para)
                    
                    content.append('\n\n'.join(formatted_text))
        
        return "\n\n".join(content)
    except Exception as e:
        return f"Error reading PDF file: {str(e)}" 