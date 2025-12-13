"""Extract redlines (tracked changes) from Microsoft Word documents."""

from typing import List, Dict, Optional
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
import re

# Try to import docx, but handle lxml import errors gracefully
try:
    from docx import Document
    from docx.oxml import parse_xml
    from docx.oxml.ns import qn
    DOCX_AVAILABLE = True
except ImportError as e:
    DOCX_AVAILABLE = False
    print(f"WARNING: python-docx is not available: {e}")
    print("Word document processing will be limited.")
    # Create dummy classes to prevent errors
    class Document:
        def __init__(self, *args, **kwargs):
            raise ImportError("python-docx is not available. lxml installation failed.")
    def parse_xml(*args, **kwargs):
        raise ImportError("python-docx is not available. lxml installation failed.")
    def qn(*args, **kwargs):
        raise ImportError("python-docx is not available. lxml installation failed.")


class WordRedlineExtractor:
    """Extracts tracked changes from Word documents."""
    
    def __init__(self, doc_path: str):
        """Initialize with document path."""
        self.doc_path = Path(doc_path)
        self.document = Document(doc_path)
        self.redlines: List[Dict] = []
        self._extract_redlines()
    
    def _extract_redlines(self) -> None:
        """Extract ONLY actual tracked changes (redlines) from the document.
        
        CRITICAL: This method ONLY extracts actual Word tracked changes:
        - <w:ins> elements (insertions)
        - <w:del> elements (deletions)
        
        It does NOT extract:
        - Regular document text
        - Formatting changes
        - Any text that is not explicitly marked as a tracked change in the XML
        
        Uses a look-ahead strategy to distinguish between:
        - Pure Deletions: standalone <w:del> elements
        - Replacements: <w:del> immediately followed by <w:ins>
        """
        # Word stores tracked changes in the document.xml file
        # We need to access the underlying XML to get revision tracking info
        
        docx_file = zipfile.ZipFile(self.doc_path, 'r')
        document_xml = docx_file.read('word/document.xml')
        docx_file.close()
        
        # Parse XML
        root = ET.fromstring(document_xml)
        
        # Define namespaces
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }
        
        # Track processed insertions to avoid double-processing
        processed_insertions = set()
        
        # Helper function to collect all tracked change elements from a paragraph in order
        def collect_tracked_changes(para):
            """Collect all w:del and w:ins elements from paragraph in document order.
            
            Uses recursive search to find ALL tracked changes within the paragraph,
            regardless of nesting level. Uses depth-first traversal to preserve order.
            """
            w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            changes = []
            seen = set()
            
            def collect_in_order(elem):
                """Recursively collect tracked changes in document order (depth-first)."""
                if elem is None:
                    return
                
                # Check if this element itself is a tracked change
                if elem.tag == f'{w_ns}del' or elem.tag == f'{w_ns}ins':
                    elem_id = id(elem)
                    if elem_id not in seen:
                        changes.append(elem)
                        seen.add(elem_id)
                
                # Recursively process children in order
                for child in elem:
                    collect_in_order(child)
            
            # Start collection from paragraph (depth-first preserves document order)
            collect_in_order(para)
            
            return changes
        
        # Iterate through all paragraphs
        paragraphs = root.findall('.//w:p', namespaces)
        
        print(f"Processing {len(paragraphs)} paragraph(s) for tracked changes")
        
        # Also do a quick global check to see if there are any tracked changes at all
        global_insertions = root.findall('.//w:ins', namespaces)
        global_deletions = root.findall('.//w:del', namespaces)
        print(f"Global search found: {len(global_insertions)} insertion(s), {len(global_deletions)} deletion(s)")
        
        # Process each paragraph's tracked changes to detect adjacent deletions/insertions
        total_tracked_changes_found = 0
        for para in paragraphs:
            # Collect all tracked changes in this paragraph in order
            tracked_changes = collect_tracked_changes(para)
            total_tracked_changes_found += len(tracked_changes)
            
            if not tracked_changes:
                continue
            
            i = 0
            while i < len(tracked_changes):
                change_elem = tracked_changes[i]
                w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
                
                # Check if this is a deletion element
                if change_elem.tag == f'{w_ns}del':
                    del_elem = change_elem
                    
                    # Get deletion metadata
                    del_author = del_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
                    del_date = del_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')
                    
                    # Extract deleted text - preserve all text content
                    del_text_elements = del_elem.findall('.//w:delText', namespaces)
                    del_text_parts = []
                    for elem in del_text_elements:
                        if elem.text:
                            del_text_parts.append(elem.text)
                        # Also check for tail text (text after the element)
                        if elem.tail:
                            del_text_parts.append(elem.tail)
                    old_text = ''.join(del_text_parts)
                    # Preserve original text - only normalize excessive whitespace, don't lose content
                    if old_text.strip():
                        # Replace multiple spaces/tabs/newlines with single space, but preserve structure
                        import re
                        old_text = re.sub(r'\s+', ' ', old_text).strip()
                    else:
                        old_text = old_text
                    
                    # Look ahead: Check if next tracked change is an insertion
                    if i + 1 < len(tracked_changes):
                        next_change = tracked_changes[i + 1]
                        
                        # Condition A: Replacement - deletion immediately followed by insertion
                        if next_change.tag == f'{w_ns}ins':
                            ins_elem = next_change
                            
                            # Get insertion metadata
                            ins_author = ins_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
                            ins_date = ins_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')
                            
                            # Extract inserted text - preserve all text content
                            ins_text_elements = ins_elem.findall('.//w:t', namespaces)
                            ins_text_parts = []
                            for elem in ins_text_elements:
                                if elem.text:
                                    ins_text_parts.append(elem.text)
                                # Also check for tail text (text after the element)
                                if elem.tail:
                                    ins_text_parts.append(elem.tail)
                            new_text = ''.join(ins_text_parts)
                            # Preserve original text - only normalize excessive whitespace, don't lose content
                            if new_text.strip():
                                # Replace multiple spaces/tabs/newlines with single space, but preserve structure
                                import re
                                new_text = re.sub(r'\s+', ' ', new_text).strip()
                            else:
                                new_text = new_text
                            
                            # Only add if both old and new text are non-empty
                            if old_text.strip() or new_text.strip():
                                # Use deletion author/date as primary, fallback to insertion if needed
                                author = del_author if del_author != 'Unknown' else ins_author
                                date = del_date if del_date else ins_date
                                
                                self.redlines.append({
                                    'type': 'replacement',
                                    'old_text': old_text,
                                    'new_text': new_text,
                                    'text': f"{old_text} → {new_text}",  # Include both for better context
                                    'author': author,
                                    'date': date,
                                    'del_element': del_elem,
                                    'ins_element': ins_elem
                                })
                                print(f"  Extracted replacement #{len(self.redlines)}:")
                                print(f"    Old text (deleted): '{old_text}'")
                                print(f"    New text (inserted): '{new_text}'")
                                print(f"    Full replacement: '{old_text}' → '{new_text}'")
                            
                            # Mark this insertion as processed so we skip it in the main loop
                            processed_insertions.add(id(ins_elem))
                            
                            # Skip both deletion and insertion (increment by 2)
                            i += 2
                            continue
                    
                    # Condition B: Pure Deletion - no adjacent insertion
                    if old_text.strip():
                        self.redlines.append({
                            'type': 'deletion',
                            'text': old_text,
                            'old_text': old_text,  # Add for consistency
                            'author': del_author,
                            'date': del_date,
                            'element': del_elem
                        })
                        print(f"  Extracted deletion #{len(self.redlines)}: '{old_text[:50]}...' (length: {len(old_text)})")
                
                # Check if this is an insertion element that hasn't been processed yet
                elif change_elem.tag == f'{w_ns}ins':
                    ins_elem = change_elem
                    
                    # Skip if this insertion was already grouped with a deletion
                    if id(ins_elem) in processed_insertions:
                        i += 1
                        continue
                    
                    # This is a standalone insertion (not part of a replacement)
                    author = ins_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
                    date = ins_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')
                    
                    # Get text content from within the <w:ins> element
                    text_elements = ins_elem.findall('.//w:t', namespaces)
                    text_parts = []
                    for elem in text_elements:
                        if elem.text:
                            text_parts.append(elem.text)
                    text = ''.join(text_parts)
                    
                    # Normalize whitespace but preserve structure
                    text = ' '.join(text.split()) if text.strip() else text
                    
                    # Only add if there's actual text (not empty)
                    if text.strip():
                        self.redlines.append({
                            'type': 'insertion',
                            'text': text,
                            'new_text': text,  # Add for consistency
                            'author': author,
                            'date': date,
                            'element': ins_elem
                        })
                        print(f"  Extracted insertion #{len(self.redlines)}: '{text[:50]}...' (length: {len(text)})")
                
                # Move to next child
                i += 1
        
        print(f"Paragraph-based approach found {total_tracked_changes_found} tracked change element(s) across all paragraphs")
        
        # Fallback: If paragraph-based approach found nothing, use global search
        # (This handles edge cases where structure might be different)
        if not self.redlines:
            print("No redlines found with paragraph-based approach, trying global search...")
            insertions = root.findall('.//w:ins', namespaces)
            deletions = root.findall('.//w:del', namespaces)
            
            print(f"Found {len(insertions)} insertion(s) and {len(deletions)} deletion(s) via global search")
            
            # Process insertions
            for ins in insertions:
                author = ins.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
                date = ins.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')
                text_elements = ins.findall('.//w:t', namespaces)
                text_parts = [elem.text for elem in text_elements if elem.text]
                text = ''.join(text_parts)
                text = ' '.join(text.split()) if text.strip() else text
                
                if text.strip():
                    self.redlines.append({
                        'type': 'insertion',
                        'text': text,
                        'new_text': text,
                        'author': author,
                        'date': date,
                        'element': ins
                    })
            
            # Process deletions
            for del_elem in deletions:
                author = del_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
                date = del_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')
                text_elements = del_elem.findall('.//w:delText', namespaces)
                text_parts = [elem.text for elem in text_elements if elem.text]
                text = ''.join(text_parts)
                text = ' '.join(text.split()) if text.strip() else text
                
                if text.strip():
                    self.redlines.append({
                        'type': 'deletion',
                        'text': text,
                        'old_text': text,
                        'author': author,
                        'date': date,
                        'element': del_elem
                    })
        
        # CRITICAL: Only extract actual tracked changes (w:ins and w:del)
        # Do NOT use alternative methods that might identify regular text as redlines
        # If no XML-based redlines found, that means there are no tracked changes
        if not self.redlines:
            print("No tracked changes (redlines) found in document. Only actual tracked changes are processed.")
        else:
            print(f"Total redlines extracted: {len(self.redlines)}")
    
    def get_redlines(self) -> List[Dict]:
        """Get all extracted redlines."""
        return self.redlines
    
    def get_redlines_summary(self) -> str:
        """Get a text summary of all redlines for AI analysis."""
        summary_parts = []
        for idx, redline in enumerate(self.redlines, 1):
            summary_parts.append(
                f"Redline #{idx}:\n"
                f"Type: {redline['type']}\n"
                f"Text: {redline['text']}\n"
                f"Author: {redline.get('author', 'Unknown')}\n"
                f"Date: {redline.get('date', 'Unknown')}\n"
            )
        return '\n---\n'.join(summary_parts)
    
    def get_document_text(self) -> str:
        """Get full document text."""
        return '\n'.join([para.text for para in self.document.paragraphs])


