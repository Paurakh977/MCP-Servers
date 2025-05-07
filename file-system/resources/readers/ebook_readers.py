"""Ebook file readers module for EPUB and other ebook formats."""

# For EPUB e-books
try:
    import ebooklib
    from ebooklib import epub
    from bs4 import BeautifulSoup
    has_epub_support = True
except ImportError:
    has_epub_support = False
    print("EPUB support not available. To read EPUB files: pip install ebooklib beautifulsoup4")

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