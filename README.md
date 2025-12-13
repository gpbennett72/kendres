# RedLine Legal Analysis Agent

An AI-powered agent that extracts redlines (tracked changes) from Microsoft Word documents and Google Docs, analyzes them against a legal playbook, and provides intelligent responses as inline comments.

## Features

- **Word Document Support**: Extracts tracked changes from `.docx` files
- **Google Docs Support**: Extracts revisions and changes from Google Docs
- **AI-Powered Analysis**: Uses AI models to analyze redlines against legal playbooks
- **Inline Comments**: Automatically adds responses as comments in the document
- **Flexible Playbooks**: Support for custom legal playbooks in various formats

## Installation

```bash
pip install -r requirements.txt
```

## Configuration

1. Create a `.env` file with your API credentials:
```env
OPENAI_API_KEY=your_key_here
ANTHROPIC_API_KEY=your_key_here
GOOGLE_CREDENTIALS_PATH=path/to/credentials.json
```

2. For Google Docs, you'll need to set up OAuth2 credentials:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a project and enable Google Docs API
   - Create OAuth2 credentials and download as JSON
   - Place the file in your project directory

## Usage

### Web Interface (Recommended for Testing)

The easiest way to test the agent is using the web interface:

```bash
python app.py
```

Then open your browser to: `http://localhost:5000`

The web interface provides:
- **Word Document Processing**: Upload a Word document and playbook, configure AI settings, and process
- **Google Doc Processing**: Enter a Google Doc ID and playbook content directly
- **Analysis Only Mode**: Analyze redlines without modifying the document
- **Results Display**: View detailed analysis results with risk levels
- **File Downloads**: Download processed documents with comments

### Command Line Interface

```bash
# Analyze a Word document
python redline_agent.py --input document.docx --playbook playbook.txt --output document_with_comments.docx

# Analyze a Google Doc
python redline_agent.py --input "https://docs.google.com/document/d/DOC_ID" --playbook playbook.txt --format google

# Use a specific AI model
python redline_agent.py --input document.docx --playbook playbook.txt --model gpt-4
```

### Python API

```python
from redline_agent import RedlineAgent

agent = RedlineAgent(
    playbook_path="legal_playbook.txt",
    ai_provider="openai",
    model="gpt-4"
)

# Process Word document
result = agent.process_word_document("document.docx", "output.docx")

# Process Google Doc
result = agent.process_google_doc("DOC_ID", "output.docx")
```

## Legal Playbook Format

The playbook can be a text file containing:
- Legal principles and guidelines
- Standard responses to common redline scenarios
- Risk assessment criteria
- Approval/rejection guidelines

Example:
```
PRINCIPLE: All liability caps must be reviewed by legal team
RESPONSE: Request removal of liability cap or reduce to standard $X amount

PRINCIPLE: Indemnification clauses require mutual indemnification
RESPONSE: Suggest mutual indemnification language
```

## Architecture

- `app.py`: Flask web application and API endpoints
- `redline_agent.py`: Main agent orchestrator
- `word_extractor.py`: Word document redline extraction
- `google_extractor.py`: Google Docs redline extraction
- `ai_analyzer.py`: AI-powered analysis engine
- `comment_inserter.py`: Comment insertion utilities
- `playbook_loader.py`: Legal playbook parser
- `templates/`: HTML templates for web interface
- `static/`: CSS and JavaScript for web interface

