"""Main Redline Analysis Agent - Orchestrates the entire workflow."""

import argparse
import os
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv

from playbook_loader import PlaybookLoader
from word_extractor import WordRedlineExtractor
from google_extractor import GoogleDocsRedlineExtractor
from ai_analyzer import AIAnalyzer
from comment_inserter import CommentInserter


# Load environment variables
load_dotenv()


class RedlineAgent:
    """Main agent for analyzing redlines in legal documents."""
    
    def __init__(
        self,
        playbook_path: str,
        ai_provider: str = None,
        model: str = None
    ):
        """Initialize the agent with playbook and AI configuration."""
        self.playbook = PlaybookLoader(playbook_path)
        self.ai_provider = ai_provider or os.getenv('DEFAULT_AI_PROVIDER', 'openai')
        self.model = model or os.getenv('DEFAULT_MODEL', 'gpt-4')
        self.analyzer = AIAnalyzer(self.playbook, self.ai_provider, self.model)
    
    def process_word_document(
        self,
        input_path: str,
        output_path: str,
        create_summary: bool = True,
        use_tracked_changes: bool = False
    ) -> dict:
        """Process a Word document and add comments or tracked changes.
        
        Args:
            input_path: Path to input Word document
            output_path: Path to save output document
            create_summary: Whether to create a summary document
            use_tracked_changes: If True, insert responses as tracked changes (counter redlines).
                                If False, insert as formatted annotations.
        """
        print(f"Extracting redlines from {input_path}...")
        extractor = WordRedlineExtractor(input_path)
        all_redlines = extractor.get_redlines()
        
        # CRITICAL: Filter to ONLY actual tracked changes (insertions, deletions, and replacements)
        # Reject any redline that is not a tracked change
        redlines = []
        for rl in all_redlines:
            if rl.get('type') in ['insertion', 'deletion', 'replacement']:
                # Ensure 'text' field exists and is properly formatted for replacements
                if rl.get('type') == 'replacement':
                    old_text = rl.get('old_text', '')
                    new_text = rl.get('new_text', '')
                    # Update text field to include both old and new text for better context
                    if old_text or new_text:
                        rl['text'] = f"{old_text} → {new_text}"
                    elif 'text' not in rl:
                        rl['text'] = f"[REPLACEMENT: '{old_text}' -> '{new_text}']"
                redlines.append(rl)
            else:
                print(f"⚠ REJECTED: Redline with type '{rl.get('type')}' is not a tracked change. Only 'insertion', 'deletion', and 'replacement' are processed.")
        
        document_text = extractor.get_document_text()
        
        if not redlines:
            print("No tracked changes (redlines) found in document. Only actual tracked changes are processed.")
            return {
                'redlines_count': 0,
                'analyses': [],
                'output_path': None
            }
        
        print(f"Found {len(redlines)} tracked change(s) (redlines). Analyzing each one individually with AI...")
        for idx, rl in enumerate(redlines, 1):
            print(f"  Redline #{idx}: {rl.get('type')} - {rl.get('text', '')[:50]}...")
        
        # Analyze redlines - pass the extractor so we can access XML elements
        analyses = self.analyzer.analyze_redlines(redlines, document_text)
        
        # Store the extractor for use in insertion
        print(f"Analysis complete. Inserting native Word comments (comment bubbles) for each redline...")
        
        # Insert comments or tracked changes - pass extractor for element references
        inserter = CommentInserter(doc_path=input_path)
        inserter.insert_comments_word(
            analyses, 
            output_path, 
            use_tracked_changes=use_tracked_changes,
            extractor=extractor
        )
        
        # Create summary document if requested
        summary_path = None
        if create_summary:
            summary_path = output_path.replace('.docx', '_summary.docx')
            inserter.create_summary_document(analyses, summary_path)
            print(f"Summary document created: {summary_path}")
        
        print(f"Formatted annotations inserted directly in document. Output saved to: {output_path}")
        
        return {
            'redlines_count': len(redlines),
            'analyses': analyses,
            'output_path': output_path,
            'summary_path': summary_path,
            'document_text': document_text
        }
    
    def process_google_doc(
        self,
        doc_id: str,
        output_path: Optional[str] = None,
        create_summary: bool = True
    ) -> dict:
        """Process a Google Doc and add comments."""
        print(f"Extracting redlines from Google Doc {doc_id}...")
        extractor = GoogleDocsRedlineExtractor(doc_id)
        redlines = extractor.get_redlines()
        document_text = extractor.get_document_text()
        
        if not redlines:
            print("No redlines found in document.")
            return {
                'redlines_count': 0,
                'analyses': [],
                'output_path': None
            }
        
        print(f"Found {len(redlines)} redlines. Analyzing with AI...")
        
        # Analyze redlines
        analyses = self.analyzer.analyze_redlines(redlines, document_text)
        
        print(f"Analysis complete. Inserting comments...")
        
        # Insert comments into Google Doc
        inserter = CommentInserter(doc_id=doc_id)
        inserter.insert_comments_google(analyses)
        
        # Create summary document if requested
        summary_path = None
        if create_summary and output_path:
            inserter.create_summary_document(analyses, output_path)
            summary_path = output_path
            print(f"Summary document created: {summary_path}")
        
        print("Comments inserted into Google Doc.")
        
        return {
            'redlines_count': len(redlines),
            'analyses': analyses,
            'output_path': summary_path,
            'document_text': document_text
        }
    
    def analyze_only(
        self,
        input_path: str = None,
        doc_id: str = None
    ) -> dict:
        """Analyze redlines without inserting comments."""
        if input_path:
            extractor = WordRedlineExtractor(input_path)
        elif doc_id:
            extractor = GoogleDocsRedlineExtractor(doc_id)
        else:
            raise ValueError("Must provide either input_path or doc_id")
        
        all_redlines = extractor.get_redlines()
        document_text = extractor.get_document_text()
        
        # Filter to ONLY actual tracked changes (insertions, deletions, and replacements)
        redlines = []
        for rl in all_redlines:
            if rl.get('type') in ['insertion', 'deletion', 'replacement']:
                # Ensure 'text' field exists for backward compatibility
                if rl.get('type') == 'replacement' and 'text' not in rl:
                    # For replacements, create a descriptive text field
                    old_text = rl.get('old_text', '')
                    new_text = rl.get('new_text', '')
                    rl['text'] = f"[REPLACEMENT: '{old_text}' -> '{new_text}']"
                redlines.append(rl)
        
        if not redlines:
            return {
                'redlines_count': 0,
                'analyses': []
            }
        
        analyses = self.analyzer.analyze_redlines(redlines, document_text)
        
        return {
            'redlines_count': len(redlines),
            'analyses': analyses
        }


def main():
    """Command-line interface for the Redline Agent."""
    parser = argparse.ArgumentParser(
        description='Analyze redlines in legal documents against a legal playbook'
    )
    parser.add_argument(
        '--input', '-i',
        required=True,
        help='Input document path (Word .docx) or Google Doc URL/ID'
    )
    parser.add_argument(
        '--playbook', '-p',
        required=True,
        help='Path to legal playbook file'
    )
    parser.add_argument(
        '--output', '-o',
        help='Output document path (default: input_with_comments.docx)'
    )
    parser.add_argument(
        '--format', '-f',
        choices=['word', 'google'],
        default='word',
        help='Document format (default: word)'
    )
    parser.add_argument(
        '--model', '-m',
        help='AI model to use (default: from .env or gpt-4)'
    )
    parser.add_argument(
        '--provider',
        choices=['openai', 'anthropic'],
        help='AI provider (default: from .env or openai)'
    )
    parser.add_argument(
        '--no-summary',
        action='store_true',
        help='Skip creating summary document'
    )
    parser.add_argument(
        '--analyze-only',
        action='store_true',
        help='Only analyze, do not insert comments'
    )
    
    args = parser.parse_args()
    
    # Determine output path
    if not args.output:
        if args.format == 'word':
            input_path = Path(args.input)
            args.output = str(input_path.parent / f"{input_path.stem}_with_comments{input_path.suffix}")
        else:
            args.output = "google_doc_summary.docx"
    
    # Initialize agent
    agent = RedlineAgent(
        playbook_path=args.playbook,
        ai_provider=args.provider,
        model=args.model
    )
    
    # Process document
    if args.analyze_only:
        if args.format == 'word':
            result = agent.analyze_only(input_path=args.input)
        else:
            # Extract doc ID from URL if needed
            doc_id = args.input
            if 'docs.google.com' in args.input:
                # Extract ID from URL
                parts = args.input.split('/')
                doc_id = [p for p in parts if len(p) > 20][0] if any(len(p) > 20 for p in parts) else args.input
            result = agent.analyze_only(doc_id=doc_id)
        
        # Print results
        print(f"\nAnalysis Results:")
        print(f"Redlines found: {result['redlines_count']}")
        for idx, analysis in enumerate(result['analyses'], 1):
            print(f"\nRedline #{idx}:")
            print(f"  Assessment: {analysis.get('assessment', 'N/A')}")
            print(f"  Risk Level: {analysis.get('risk_level', 'N/A')}")
            print(f"  Response: {analysis.get('response', 'N/A')}")
    
    else:
        if args.format == 'word':
            result = agent.process_word_document(
                args.input,
                args.output,
                create_summary=not args.no_summary
            )
        else:
            # Extract doc ID from URL if needed
            doc_id = args.input
            if 'docs.google.com' in args.input:
                parts = args.input.split('/')
                doc_id = [p for p in parts if len(p) > 20][0] if any(len(p) > 20 for p in parts) else args.input
            
            result = agent.process_google_doc(
                doc_id,
                args.output if not args.no_summary else None,
                create_summary=not args.no_summary
            )
        
        print(f"\nProcessing complete!")
        print(f"Redlines analyzed: {result['redlines_count']}")
        if result.get('output_path'):
            print(f"Output saved to: {result['output_path']}")
        if result.get('summary_path'):
            print(f"Summary saved to: {result['summary_path']}")


if __name__ == '__main__':
    main()

