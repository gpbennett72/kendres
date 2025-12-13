#!/usr/bin/env python3
"""Test script to create a minimal Word document with a comment to verify structure."""

import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
import os
import tempfile

def create_test_docx_with_comment():
    """Create a minimal Word document with a comment to test the structure."""
    
    # Create a temporary directory for our test document
    temp_dir = tempfile.mkdtemp()
    test_docx_path = os.path.join(temp_dir, 'test_comment.docx')
    
    print("Creating test Word document with comment...")
    
    # Create minimal Word document structure
    with zipfile.ZipFile(test_docx_path, 'w', zipfile.ZIP_DEFLATED) as docx:
        # Create [Content_Types].xml
        content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>'''
        docx.writestr('[Content_Types].xml', content_types)
        
        # Create _rels/.rels
        rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
        docx.writestr('_rels/.rels', rels)
        
        # Create word/_rels/document.xml.rels
        doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>'''
        docx.writestr('word/_rels/document.xml.rels', doc_rels)
        
        # Create word/document.xml with a comment
        # Structure: paragraph with text, commentRangeStart, commentReference, commentRangeEnd
        document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
<w:r>
<w:t>This is test text.</w:t>
</w:r>
<w:commentRangeStart w:id="1"/>
<w:r>
<w:t>This text has a comment.</w:t>
</w:r>
<w:r>
<w:commentReference w:id="1"/>
<w:t> </w:t>
</w:r>
<w:commentRangeEnd w:id="1"/>
<w:r>
<w:t> More text after comment.</w:t>
</w:r>
</w:p>
</w:body>
</w:document>'''
        docx.writestr('word/document.xml', document_xml)
        
        # Create word/comments.xml
        comments_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:comment w:id="1" w:author="Test Author" w:date="{datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')}">
<w:p>
<w:pPr/>
<w:r>
<w:t>This is a test comment with actual text content.</w:t>
</w:r>
</w:p>
</w:comment>
</w:comments>'''
        docx.writestr('word/comments.xml', comments_xml)
    
    print(f"✓ Test document created: {test_docx_path}")
    print(f"\nTry opening this document in Word. It should:")
    print(f"  1. Open without errors")
    print(f"  2. Show a comment bubble on 'This text has a comment.'")
    print(f"  3. Display the comment text when you hover/click")
    print(f"\nIf this works, we can compare it to your generated document.")
    print(f"If it doesn't work, there's a fundamental issue with Word or the structure.")
    
    return test_docx_path

if __name__ == '__main__':
    create_test_docx_with_comment()
