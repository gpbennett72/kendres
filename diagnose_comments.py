#!/usr/bin/env python3
"""Diagnostic script to check Word document comment structure."""

import zipfile
import xml.etree.ElementTree as ET
import sys
import os

# Try to import lxml, but provide a fallback if it's not available
try:
    from lxml import etree as lxml_etree
    LXML_AVAILABLE = True
except ImportError as e:
    LXML_AVAILABLE = False
    print(f"WARNING: lxml is not available: {e}")
    print("XML validation features will be limited.")
    # Create a dummy lxml_etree object to prevent errors
    class DummyLxmlEtree:
        @staticmethod
        def fromstring(xml_bytes):
            return ET.fromstring(xml_bytes)
        
        class XMLSyntaxError(Exception):
            pass
    
    lxml_etree = DummyLxmlEtree()

def diagnose_docx(docx_path):
    """Diagnose a Word document's comment structure."""
    print(f"\n{'='*80}")
    print(f"DIAGNOSING: {docx_path}")
    print(f"{'='*80}\n")
    
    if not os.path.exists(docx_path):
        print(f"❌ ERROR: File not found: {docx_path}")
        return
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx:
            # Check if comments.xml exists
            print("1. CHECKING comments.xml")
            print("-" * 80)
            try:
                comments_xml = docx.read('word/comments.xml')
                print(f"✓ comments.xml exists ({len(comments_xml)} bytes)")
                
                # Parse with ElementTree
                try:
                    comments_root = ET.fromstring(comments_xml)
                    print(f"✓ XML is well-formed (ElementTree)")
                except ET.ParseError as e:
                    print(f"❌ XML Parse Error (ElementTree): {e}")
                    return
                
                # Parse with lxml for validation
                try:
                    lxml_root = lxml_etree.fromstring(comments_xml)
                    print(f"✓ XML is well-formed (lxml)")
                except lxml_etree.XMLSyntaxError as e:
                    print(f"❌ XML Parse Error (lxml): {e}")
                    return
                
                # Check namespace
                print(f"\n2. CHECKING NAMESPACE")
                print("-" * 80)
                comments_str = comments_xml.decode('utf-8')
                if 'xmlns:w=' in comments_str:
                    print("✓ Uses 'w:' namespace prefix")
                elif 'xmlns:ns0=' in comments_str:
                    print("⚠ Uses 'ns0:' namespace prefix (should be 'w:')")
                else:
                    print("⚠ No standard namespace found")
                
                # Count comments
                namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                comments = comments_root.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comment', namespaces)
                print(f"\n3. COMMENT COUNT")
                print("-" * 80)
                print(f"Found {len(comments)} comment(s)")
                
                # Check each comment
                for idx, comment in enumerate(comments, 1):
                    comment_id = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                    author = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
                    date = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', 'Unknown')
                    
                    print(f"\n   Comment #{idx} (ID: {comment_id})")
                    print(f"   Author: {author}")
                    print(f"   Date: {date}")
                    
                    # Check for paragraphs
                    paragraphs = comment.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p', namespaces)
                    print(f"   Paragraphs: {len(paragraphs)}")
                    
                    # Check for text
                    text_elems = comment.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t', namespaces)
                    print(f"   Text elements: {len(text_elems)}")
                    
                    # Extract all text
                    all_text = []
                    for text_elem in text_elems:
                        if text_elem.text:
                            all_text.append(text_elem.text)
                    
                    if all_text:
                        combined_text = ''.join(all_text)
                        print(f"   Text length: {len(combined_text)} chars")
                        print(f"   First 100 chars: {combined_text[:100]}")
                    else:
                        print(f"   ❌ NO TEXT FOUND!")
                    
                    # Show XML structure
                    comment_xml = ET.tostring(comment, encoding='utf-8').decode('utf-8')
                    print(f"   XML (first 300 chars): {comment_xml[:300]}")
                
            except KeyError:
                print("❌ comments.xml NOT FOUND in document!")
                return
            
            # Check document.xml for comment markers
            print(f"\n4. CHECKING document.xml FOR COMMENT MARKERS")
            print("-" * 80)
            try:
                document_xml = docx.read('word/document.xml')
                root = ET.fromstring(document_xml)
                
                # Find comment range markers
                comment_range_starts = root.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart', namespaces)
                comment_references = root.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentReference', namespaces)
                comment_range_ends = root.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd', namespaces)
                
                print(f"CommentRangeStart markers: {len(comment_range_starts)}")
                print(f"CommentReference markers: {len(comment_references)}")
                print(f"CommentRangeEnd markers: {len(comment_range_ends)}")
                
                # Check for matching IDs
                start_ids = set()
                for start in comment_range_starts:
                    cid = start.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                    if cid:
                        start_ids.add(cid)
                
                ref_ids = set()
                for ref in comment_references:
                    cid = ref.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                    if cid:
                        ref_ids.add(cid)
                
                end_ids = set()
                for end in comment_range_ends:
                    cid = end.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                    if cid:
                        end_ids.add(cid)
                
                print(f"\n5. COMMENT ID MATCHING")
                print("-" * 80)
                print(f"Start IDs: {sorted(start_ids)}")
                print(f"Reference IDs: {sorted(ref_ids)}")
                print(f"End IDs: {sorted(end_ids)}")
                
                # Check for mismatches
                if start_ids != ref_ids:
                    print(f"⚠ WARNING: Start IDs don't match Reference IDs!")
                    print(f"   Missing in refs: {start_ids - ref_ids}")
                    print(f"   Extra in refs: {ref_ids - start_ids}")
                
                if start_ids != end_ids:
                    print(f"⚠ WARNING: Start IDs don't match End IDs!")
                    print(f"   Missing in ends: {start_ids - end_ids}")
                    print(f"   Extra in ends: {end_ids - start_ids}")
                
                # Check if comment IDs in comments.xml match
                comment_ids_in_comments = set()
                for comment in comments:
                    cid = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                    if cid:
                        comment_ids_in_comments.add(cid)
                
                if start_ids != comment_ids_in_comments:
                    print(f"⚠ WARNING: Document comment IDs don't match comments.xml IDs!")
                    print(f"   In document.xml but not in comments.xml: {start_ids - comment_ids_in_comments}")
                    print(f"   In comments.xml but not in document.xml: {comment_ids_in_comments - start_ids}")
                else:
                    print(f"✓ All comment IDs match between document.xml and comments.xml")
                
            except Exception as e:
                print(f"❌ Error reading document.xml: {e}")
                import traceback
                traceback.print_exc()
            
            # Check XML validity
            print(f"\n6. XML VALIDATION")
            print("-" * 80)
            try:
                # Validate comments.xml
                lxml_etree.fromstring(comments_xml)
                print("✓ comments.xml is valid XML")
            except lxml_etree.XMLSyntaxError as e:
                print(f"❌ comments.xml has XML syntax errors:")
                print(f"   {e}")
            
            try:
                # Validate document.xml
                lxml_etree.fromstring(document_xml)
                print("✓ document.xml is valid XML")
            except lxml_etree.XMLSyntaxError as e:
                print(f"❌ document.xml has XML syntax errors:")
                print(f"   {e}")
    
    except zipfile.BadZipFile:
        print(f"❌ ERROR: Not a valid ZIP file (corrupted .docx)")
    except Exception as e:
        print(f"❌ ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python diagnose_comments.py <path_to_docx_file>")
        print("\nExample:")
        print("  python diagnose_comments.py output_MUTUAL_NON-v2.docx")
        sys.exit(1)
    
    diagnose_docx(sys.argv[1])







