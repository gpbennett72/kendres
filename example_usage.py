"""Example usage of the RedLine Agent."""

from redline_agent import RedlineAgent
import os

# Example 1: Process a Word document
def example_word_document():
    """Example of processing a Word document."""
    agent = RedlineAgent(
        playbook_path="sample_playbook.txt",
        ai_provider="openai",  # or "anthropic"
        model="gpt-4"  # or "claude-3-opus-20240229" for Anthropic
    )
    
    result = agent.process_word_document(
        input_path="contract_with_redlines.docx",
        output_path="contract_with_ai_comments.docx",
        create_summary=True
    )
    
    print(f"Processed {result['redlines_count']} redlines")
    print(f"Output saved to: {result['output_path']}")
    if result.get('summary_path'):
        print(f"Summary saved to: {result['summary_path']}")


# Example 2: Process a Google Doc
def example_google_doc():
    """Example of processing a Google Doc."""
    agent = RedlineAgent(
        playbook_path="sample_playbook.txt",
        ai_provider="openai",
        model="gpt-4"
    )
    
    # Extract doc ID from Google Docs URL
    # URL format: https://docs.google.com/document/d/DOC_ID/edit
    doc_id = "YOUR_GOOGLE_DOC_ID_HERE"
    
    result = agent.process_google_doc(
        doc_id=doc_id,
        output_path="google_doc_summary.docx",
        create_summary=True
    )
    
    print(f"Processed {result['redlines_count']} redlines")
    print("Comments have been added to the Google Doc")


# Example 3: Analyze only (without inserting comments)
def example_analyze_only():
    """Example of analyzing redlines without inserting comments."""
    agent = RedlineAgent(
        playbook_path="sample_playbook.txt",
        ai_provider="openai",
        model="gpt-4"
    )
    
    result = agent.analyze_only(input_path="contract_with_redlines.docx")
    
    print(f"Found {result['redlines_count']} redlines")
    print("\nAnalyses:")
    for idx, analysis in enumerate(result['analyses'], 1):
        print(f"\nRedline #{idx}:")
        print(f"  Assessment: {analysis.get('assessment', 'N/A')}")
        print(f"  Risk Level: {analysis.get('risk_level', 'N/A')}")
        print(f"  Response: {analysis.get('response', 'N/A')}")


if __name__ == "__main__":
    # Make sure you have set up your .env file with API keys
    # and have a playbook file ready
    
    print("RedLine Agent Examples")
    print("=" * 50)
    print("\n1. Word Document Processing:")
    print("   Uncomment example_word_document() to run")
    # example_word_document()
    
    print("\n2. Google Doc Processing:")
    print("   Uncomment example_google_doc() to run")
    # example_google_doc()
    
    print("\n3. Analysis Only:")
    print("   Uncomment example_analyze_only() to run")
    # example_analyze_only()








