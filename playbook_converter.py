"""Convert MS Word playbooks to Markdown format."""

import os
import re
from typing import Optional
from pathlib import Path
import xml.etree.ElementTree as ET
from zipfile import ZipFile


class PlaybookConverter:
    """Converts Word documents to Markdown format."""
    
    def __init__(self):
        """Initialize converter."""
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }
    
    def convert_word_to_markdown(self, word_path: str, output_path: Optional[str] = None) -> str:
        """
        Convert a Word document to Markdown.
        
        Args:
            word_path: Path to the .docx file
            output_path: Optional path to save the markdown file. If None, saves next to Word file.
        
        Returns:
            Path to the created markdown file
        """
        if not os.path.exists(word_path):
            raise FileNotFoundError(f"Word document not found: {word_path}")
        
        if not word_path.endswith('.docx'):
            raise ValueError("Only .docx files are supported")
        
        # Extract text from Word document
        markdown_content = self._extract_text_from_word(word_path)
        
        # Format as markdown
        formatted_markdown = self._format_as_markdown(markdown_content)
        
        # Determine output path
        if output_path is None:
            base_name = os.path.splitext(word_path)[0]
            output_path = f"{base_name}.md"
        
        # Save markdown file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(formatted_markdown)
        
        return output_path
    
    def _extract_text_from_word(self, word_path: str) -> str:
        """Extract text content from Word document XML."""
        with ZipFile(word_path, 'r') as docx:
            # Read document.xml
            try:
                xml_content = docx.read('word/document.xml')
            except KeyError:
                raise ValueError("Invalid Word document: document.xml not found")
            
            root = ET.fromstring(xml_content)
            
            # Extract all text nodes
            text_parts = []
            self._extract_text_recursive(root, text_parts)
            
            return '\n'.join(text_parts)
    
    def _extract_text_recursive(self, element, text_parts: list, in_paragraph: bool = False):
        """Recursively extract text from XML elements."""
        w_ns = self.namespaces['w']
        
        # Check if this is a paragraph
        if element.tag == f'{w_ns}p':
            if text_parts and text_parts[-1] and not text_parts[-1].endswith('\n'):
                text_parts.append('\n')
            in_paragraph = True
        
        # Extract text from w:t elements
        if element.tag == f'{w_ns}t':
            if element.text:
                text_parts.append(element.text)
        
        # Extract tail text
        if element.tail and element.tail.strip():
            text_parts.append(element.tail)
        
        # Process children
        for child in element:
            self._extract_text_recursive(child, text_parts, in_paragraph)
        
        # Add newline after paragraph
        if element.tag == f'{w_ns}p' and text_parts and not text_parts[-1].endswith('\n'):
            text_parts.append('\n')
    
    def _format_as_markdown(self, content: str) -> str:
        """Format extracted text as markdown."""
        lines = content.split('\n')
        formatted_lines = []
        in_list = False
        in_code_block = False
        
        for i, line in enumerate(lines):
            stripped = line.strip()
            
            # Skip empty lines (but preserve structure)
            if not stripped:
                if in_list:
                    in_list = False
                formatted_lines.append('')
                continue
            
            # Detect headings (lines that are all caps or start with specific patterns)
            if self._is_heading(line, i, lines):
                if in_list:
                    in_list = False
                    formatted_lines.append('')
                heading_level = self._get_heading_level(line)
                formatted_lines.append(f"{'#' * heading_level} {stripped}")
                continue
            
            # Detect list items
            if self._is_list_item(stripped):
                if not in_list:
                    formatted_lines.append('')
                in_list = True
                # Convert to markdown list
                list_marker = self._get_list_marker(stripped)
                list_text = self._clean_list_text(stripped)
                formatted_lines.append(f"{list_marker} {list_text}")
                continue
            
            # Regular paragraph
            if in_list:
                in_list = False
                formatted_lines.append('')
            
            # Check for PRINCIPLE: and RESPONSE: patterns
            if stripped.upper().startswith('PRINCIPLE:'):
                formatted_lines.append(f"## {stripped}")
            elif stripped.upper().startswith('RESPONSE:'):
                formatted_lines.append(f"### {stripped}")
            else:
                formatted_lines.append(stripped)
        
        # Clean up multiple consecutive empty lines
        result = []
        prev_empty = False
        for line in formatted_lines:
            is_empty = not line.strip()
            if is_empty and prev_empty:
                continue
            result.append(line)
            prev_empty = is_empty
        
        return '\n'.join(result)
    
    def _is_heading(self, line: str, index: int, all_lines: list) -> bool:
        """Determine if a line is a heading."""
        stripped = line.strip()
        
        # All caps and short (likely heading)
        if stripped.isupper() and len(stripped) < 100 and len(stripped.split()) < 10:
            return True
        
        # Starts with common heading words
        heading_keywords = ['PRINCIPLE', 'SECTION', 'CHAPTER', 'PART', 'GUIDELINE', 'RULE']
        if any(stripped.upper().startswith(kw) for kw in heading_keywords):
            return True
        
        # Line followed by empty line and has few words
        if index < len(all_lines) - 1:
            next_line = all_lines[index + 1].strip()
            if not next_line and len(stripped.split()) < 15:
                return True
        
        return False
    
    def _get_heading_level(self, line: str) -> int:
        """Determine heading level (1-3)."""
        stripped = line.strip()
        
        # All caps = level 1
        if stripped.isupper():
            return 1
        
        # Starts with PRINCIPLE, SECTION, etc. = level 2
        major_keywords = ['PRINCIPLE', 'SECTION', 'CHAPTER', 'PART']
        if any(stripped.upper().startswith(kw) for kw in major_keywords):
            return 2
        
        # Otherwise level 3
        return 3
    
    def _is_list_item(self, line: str) -> bool:
        """Check if line is a list item."""
        # Bullet points
        if line.startswith('•') or line.startswith('·') or line.startswith('▪'):
            return True
        
        # Numbered lists
        if re.match(r'^\d+[\.\)]\s', line):
            return True
        
        # Letter lists
        if re.match(r'^[a-zA-Z][\.\)]\s', line):
            return True
        
        # Dash lists
        if re.match(r'^[-–—]\s', line):
            return True
        
        return False
    
    def _get_list_marker(self, line: str) -> str:
        """Get markdown list marker."""
        # Numbered
        if re.match(r'^\d+[\.\)]\s', line):
            return '1.'
        
        # Letter (convert to dash)
        if re.match(r'^[a-zA-Z][\.\)]\s', line):
            return '-'
        
        # Already a dash or bullet
        return '-'
    
    def _clean_list_text(self, line: str) -> str:
        """Remove list markers from text."""
        # Remove bullet points
        line = re.sub(r'^[•·▪]\s*', '', line)
        # Remove numbered markers
        line = re.sub(r'^\d+[\.\)]\s*', '', line)
        # Remove letter markers
        line = re.sub(r'^[a-zA-Z][\.\)]\s*', '', line)
        # Remove dash markers
        line = re.sub(r'^[-–—]\s*', '', line)
        return line.strip()





