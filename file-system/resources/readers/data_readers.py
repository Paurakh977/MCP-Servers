"""Data file readers module for CSV and other data formats."""
import csv
import os

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
        from .text_reader import read_text_file
        return read_text_file(file_path)
    except Exception as e:
        # Fall back to simple text reading if CSV parsing fails
        try:
            from .text_reader import read_text_file
            return read_text_file(file_path)
        except Exception as nested_e:
            return f"Error reading CSV file: {str(e)}, Nested error: {str(nested_e)}" 