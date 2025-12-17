"""Flask web application for RedLine Agent."""

import os
import re
import uuid
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables from .env file FIRST, before any other imports
load_dotenv()

from flask import Flask, render_template, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
import tempfile
import shutil

# Lazy imports to avoid lxml import errors at startup
# These will be imported when actually needed
def get_RedlineAgent():
    from redline_agent import RedlineAgent
    return RedlineAgent

def get_PlaybookLoader():
    from playbook_loader import PlaybookLoader
    return PlaybookLoader

def get_ContractTypesManager():
    from contract_types_manager import ContractTypesManager
    return ContractTypesManager

def get_PlaybookConverter():
    from playbook_converter import PlaybookConverter
    return PlaybookConverter

app = Flask(__name__, static_folder='static', static_url_path='/static')
# Use environment variable for secret key if available, otherwise generate one
app.secret_key = os.environ.get('FLASK_SECRET_KEY', os.urandom(24).hex())
# File size limit removed - no maximum file size restriction

# For Vercel/serverless, use /tmp directory which is writable
if os.environ.get('VERCEL'):
    app.config['UPLOAD_FOLDER'] = '/tmp/redline_uploads'
else:
    app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()

# Ensure upload directory exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'docx', 'txt', 'doc'}

# Standard playbook location
STANDARD_PLAYBOOK_PATH = os.path.join(os.path.dirname(__file__), 'playbooks', 'default_playbook.txt')


def allowed_file(filename):
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_playbook_path(contract_type_id: str = None):
    """Get playbook path for a contract type, or default."""
    if contract_type_id:
        ContractTypesManager = get_ContractTypesManager()
        ContractTypesManager = get_ContractTypesManager()
        manager = ContractTypesManager()
        playbook_path = manager.get_playbook_path(contract_type_id)
        if playbook_path:
            return playbook_path
    
    # Use standard playbook location
    if os.path.exists(STANDARD_PLAYBOOK_PATH):
        return STANDARD_PLAYBOOK_PATH
    
    # Fallback: try to find any playbook in playbooks directory
    playbooks_dir = os.path.join(os.path.dirname(__file__), 'playbooks')
    if os.path.exists(playbooks_dir):
        for file in os.listdir(playbooks_dir):
            if file.endswith('.txt'):
                return os.path.join(playbooks_dir, file)
    
    return None


def extract_parties(document_text: str) -> tuple[str, str]:
    """
    Extract Company and Counterparty names from agreement text.
    Looks for defined term patterns like: Company Name, Inc. ("Company") and ("Counterparty")
    Returns (company_name, counterparty_name).
    """
    if not document_text:
        return "", ""
    
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', document_text)
    
    def clean_name(name: str) -> str:
        """Clean and extract the core company name."""
        name = name.strip(' .,"""\'')
        name = re.sub(r'\s+', ' ', name)
        # Remove leading "and" (common when extracting second party)
        name = re.sub(r'^and\s+', '', name, flags=re.IGNORECASE)
        # Remove trailing descriptors like "a Delaware corporation with offices at..."
        name = re.sub(r',?\s*a\s+\w+\s+(?:corporation|company|LLC|limited|partnership).*$', '', name, flags=re.IGNORECASE)
        # Remove trailing address patterns
        name = re.sub(r',?\s*with\s+(?:offices?|headquarters|principal\s+place).*$', '', name, flags=re.IGNORECASE)
        name = re.sub(r',?\s*located\s+at.*$', '', name, flags=re.IGNORECASE)
        return name.strip(' .,')[:150]
    
    company_name = ""
    counterparty_name = ""
    
    # PRIORITY 1: Look for explicit ("Company") and ("Counterparty") definitions
    # Pattern: [Name, Corp type, address info] ("Company") or (the "Company") or ("Company")
    
    # Find Company - look for text before ("Company") or (the "Company") or ("Company")
    company_patterns = [
        r'([A-Z][A-Za-z0-9 .,&\'()\-]+?(?:Inc\.|LLC|Corp\.|Corporation|Ltd\.|L\.P\.|LLP|Company|Co\.))[^(]*?\(\s*(?:the\s+)?["""]Company["""]\s*\)',
        r'([A-Z][A-Za-z0-9 .,&\'()\-]{5,150})[^(]{0,200}?\(\s*(?:the\s+)?["""]Company["""]\s*\)',
    ]
    
    for pat in company_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            company_name = clean_name(m.group(1))
            break
    
    # Find Counterparty - look for text before ("Counterparty") or (the "Counterparty")
    counterparty_patterns = [
        r'([A-Z][A-Za-z0-9 .,&\'()\-]+?(?:Inc\.|LLC|Corp\.|Corporation|Ltd\.|L\.P\.|LLP|Company|Co\.))[^(]*?\(\s*(?:the\s+)?["""]Counterparty["""]\s*\)',
        r'([A-Z][A-Za-z0-9 .,&\'()\-]{5,150})[^(]{0,200}?\(\s*(?:the\s+)?["""]Counterparty["""]\s*\)',
    ]
    
    for pat in counterparty_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            counterparty_name = clean_name(m.group(1))
            break
    
    # If we found both, return them
    if company_name and counterparty_name:
        return company_name, counterparty_name
    
    # PRIORITY 2: Try other common defined terms (Disclosing Party, Receiving Party, Client, etc.)
    other_party_terms = [
        ('Disclosing Party', 'Receiving Party'),
        ('Discloser', 'Receiver'),
        ('Client', 'Vendor'),
        ('Customer', 'Provider'),
        ('Party A', 'Party B'),
        ('First Party', 'Second Party'),
    ]
    
    for term1, term2 in other_party_terms:
        if not company_name:
            pat1 = rf'([A-Z][A-Za-z0-9 .,&\'()\-]+?(?:Inc\.|LLC|Corp\.|Corporation|Ltd\.|L\.P\.|LLP|Company|Co\.))[^(]*?\(\s*(?:the\s+)?["""]?{term1}["""]?\s*\)'
            m = re.search(pat1, text, re.IGNORECASE)
            if m:
                company_name = clean_name(m.group(1))
        
        if not counterparty_name:
            pat2 = rf'([A-Z][A-Za-z0-9 .,&\'()\-]+?(?:Inc\.|LLC|Corp\.|Corporation|Ltd\.|L\.P\.|LLP|Company|Co\.))[^(]*?\(\s*(?:the\s+)?["""]?{term2}["""]?\s*\)'
            m = re.search(pat2, text, re.IGNORECASE)
            if m:
                counterparty_name = clean_name(m.group(1))
        
        if company_name and counterparty_name:
            break
    
    # PRIORITY 3: Fallback - look for "by and between" pattern
    if not company_name or not counterparty_name:
        between_pat = r'by\s+and\s+between[:\s]+([A-Z][A-Za-z0-9 .,&\'()\-]+?(?:Inc\.|LLC|Corp\.|Corporation|Ltd\.|L\.P\.|LLP|Company|Co\.)).*?(?:and|,)\s+([A-Z][A-Za-z0-9 .,&\'()\-]+?(?:Inc\.|LLC|Corp\.|Corporation|Ltd\.|L\.P\.|LLP|Company|Co\.))'
        m = re.search(between_pat, text, re.IGNORECASE)
        if m:
            if not company_name:
                company_name = clean_name(m.group(1))
            if not counterparty_name:
                counterparty_name = clean_name(m.group(2))
    
    return company_name, counterparty_name


def get_contract_type_name(contract_type_id: str = None) -> str:
    """Resolve contract type ID to its display name."""
    if not contract_type_id:
        return "Default"
    try:
        ContractTypesManager = get_ContractTypesManager()
        manager = ContractTypesManager()
        meta = manager.get_type_by_id(contract_type_id)
        if meta and isinstance(meta, dict):
            return meta.get('name') or contract_type_id
    except Exception:
        pass
    return contract_type_id


@app.route('/')
def index():
    """Main page."""
    return render_template('index.html')

@app.route('/static/<path:filename>')
def serve_static(filename):
    """Serve static files explicitly for Vercel compatibility."""
    static_dir = os.path.join(os.path.dirname(__file__), 'static')
    file_path = os.path.join(static_dir, filename)
    if os.path.exists(file_path):
        return send_file(file_path)
    else:
        from flask import abort
        abort(404)

@app.route('/api/upload', methods=['POST'])
def upload_files():
    """Handle file uploads."""
    try:
        if 'document' not in request.files:
            return jsonify({'error': 'Document file is required'}), 400
        
        document_file = request.files['document']
        
        if document_file.filename == '':
            return jsonify({'error': 'Document file must be selected'}), 400
        
        # Generate session ID for this processing session
        session_id = str(uuid.uuid4())
        session_dir = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        os.makedirs(session_dir, exist_ok=True)
        
        # Save document file
        if document_file and allowed_file(document_file.filename):
            doc_filename = secure_filename(document_file.filename)
            doc_path = os.path.join(session_dir, doc_filename)
            document_file.save(doc_path)
        else:
            return jsonify({'error': 'Invalid document file type'}), 400
        
        # Always use standard playbook location
        playbook_path = get_playbook_path()
        if not playbook_path:
            return jsonify({'error': 'No playbook found. Please configure a playbook in the Admin section.'}), 400
        
        playbook_filename = os.path.basename(playbook_path)
        
        # Store session info
        session['session_id'] = session_id
        session['doc_path'] = doc_path
        session['playbook_path'] = playbook_path
        
        return jsonify({
            'success': True,
            'session_id': session_id,
            'document': doc_filename,
            'playbook': playbook_filename
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/process', methods=['POST'])
def process_document():
    """Process the document with redline analysis."""
    import sys
    print("\n" + "="*80, flush=True)
    print("=== PROCESSING DOCUMENT REQUEST ===", flush=True)
    print("="*80 + "\n", flush=True)
    
    try:
        data = request.json
        session_id = session.get('session_id')
        doc_path = session.get('doc_path')
        playbook_path = session.get('playbook_path')
        
        print(f"Session ID: {session_id}", flush=True)
        print(f"Document path: {doc_path}", flush=True)
        
        if not session_id or not doc_path:
            print("ERROR: Session expired or document not found in session", flush=True)
            return jsonify({'error': 'Session expired. Please upload document again.'}), 400
        
        # Get contract type if provided
        contract_type_id = data.get('contract_type_id')
        
        # Get playbook path based on contract type
        playbook_path = get_playbook_path(contract_type_id)
        print(f"Contract Type ID: {contract_type_id}", flush=True)
        print(f"Playbook path: {playbook_path}", flush=True)
        
        if not os.path.exists(doc_path):
            print(f"ERROR: Document not found at {doc_path}", flush=True)
            return jsonify({'error': 'Document not found. Please upload again.'}), 404
        
        if not playbook_path or not os.path.exists(playbook_path):
            print(f"ERROR: Playbook not found at {playbook_path}", flush=True)
            return jsonify({'error': 'Playbook not found. Please configure a playbook in the Admin section.'}), 404
        
        # Get AI configuration
        ai_provider = data.get('ai_provider', 'openai')
        model = data.get('model', 'gpt-4')
        create_summary = data.get('create_summary', True)
        use_tracked_changes = data.get('use_tracked_changes', False)
        
        print(f"AI Provider: {ai_provider}", flush=True)
        print(f"AI Model: {model}", flush=True)
        print(f"Create Summary: {create_summary}", flush=True)
        print(f"Use Tracked Changes: {use_tracked_changes}", flush=True)
        print("\n" + "-"*80, flush=True)
        print("Starting document processing...", flush=True)
        print("-"*80 + "\n", flush=True)
        
        # Initialize agent
        RedlineAgent = get_RedlineAgent()
        agent = RedlineAgent(
            playbook_path=playbook_path,
            ai_provider=ai_provider,
            model=model
        )
        
        # Process document
        session_dir = os.path.dirname(doc_path)
        output_filename = f"output_{os.path.basename(doc_path)}"
        output_path = os.path.join(session_dir, output_filename)
        
        print(f"Output path: {output_path}\n", flush=True)
        
        result = agent.process_word_document(
            input_path=doc_path,
            output_path=output_path,
            create_summary=create_summary,
            use_tracked_changes=use_tracked_changes
        )
        
        print("\n" + "-"*80, flush=True)
        print("Document processing complete!", flush=True)
        print(f"Redlines found: {result.get('redlines_count', 0)}", flush=True)
        print(f"Analyses completed: {len(result.get('analyses', []))}", flush=True)
        print("-"*80 + "\n", flush=True)
        
        # Store output paths in session
        session['output_path'] = output_path
        if result.get('summary_path'):
            session['summary_path'] = result['summary_path']
        
        # Prepare response with analysis details
        analyses_data = []
        for analysis in result.get('analyses', []):
            redline = analysis.get('redline', {})
            redline_type = redline.get('type', 'Unknown')
            
            # Handle text field - for replacements, show both old and new text
            if redline_type == 'replacement':
                old_text = redline.get('old_text', '')
                new_text = redline.get('new_text', '')
                display_text = f"'{old_text}' → '{new_text}'"
            else:
                display_text = redline.get('text', '')
            
            # Truncate if too long
            if len(display_text) > 100:
                display_text = display_text[:100] + '...'
            
            # Combine assessment and comment_text into a single assessment field
            assessment = analysis.get('assessment', '')
            comment_text = analysis.get('comment_text', '')
            
            # Combine assessment and comment if both exist
            if assessment and comment_text:
                combined_assessment = f"{assessment}\n\n{comment_text}"
            elif comment_text:
                combined_assessment = comment_text
            else:
                combined_assessment = assessment
            
            analyses_data.append({
                'type': redline_type,
                'text': display_text,
                'assessment': combined_assessment,
                'risk_level': analysis.get('risk_level', 'Medium'),
                'response': analysis.get('response', '')
            })
        
        document_text = result.get('document_text', '')
        party_one, party_two = extract_parties(document_text)
        contract_type_id = data.get('contract_type_id')
        metadata = {
            'document': os.path.basename(doc_path),
            'party_one': party_one or 'Not detected',
            'party_two': party_two or 'Not detected',
            'counterparty': party_two or 'Not detected',
            'contract_type': get_contract_type_name(contract_type_id),
            'playbook': os.path.basename(playbook_path),
            'ai_provider': ai_provider,
            'model': model,
            'redlines_count': result['redlines_count']
        }
        
        response = {
            'success': True,
            'redlines_count': result['redlines_count'],
            'analyses': analyses_data,
            'output_file': output_filename,
            'summary_file': os.path.basename(result.get('summary_path', '')) if result.get('summary_path') else None,
            'metadata': metadata
        }
        
        print("\n" + "="*80, flush=True)
        print("=== PROCESSING COMPLETE - SENDING RESPONSE ===", flush=True)
        print(f"Redlines analyzed: {result['redlines_count']}", flush=True)
        print(f"Analyses in response: {len(analyses_data)}", flush=True)
        print("="*80 + "\n", flush=True)
        
        return jsonify(response)
    
    except Exception as e:
        import traceback
        print("\n" + "="*80, flush=True)
        print("ERROR DURING PROCESSING:", flush=True)
        print(f"Type: {type(e).__name__}", flush=True)
        print(f"Message: {str(e)}", flush=True)
        print("\nTraceback:", flush=True)
        traceback.print_exc()
        print("="*80 + "\n", flush=True)
        return jsonify({'error': str(e)}), 500


@app.route('/api/process-google', methods=['POST'])
def process_google_doc():
    """Process a Google Doc."""
    try:
        data = request.json
        doc_id = data.get('doc_id')
        
        if not doc_id:
            return jsonify({'error': 'Google Doc ID is required'}), 400
        
        # Get contract type if provided
        contract_type_id = data.get('contract_type_id')
        
        # Get playbook path based on contract type
        playbook_path = get_playbook_path(contract_type_id)
        if not playbook_path:
            return jsonify({'error': 'No playbook found. Please configure a playbook in the Admin section.'}), 400
        
        # Get AI configuration
        ai_provider = data.get('ai_provider', 'openai')
        model = data.get('model', 'gpt-4')
        
        # Initialize agent
        RedlineAgent = get_RedlineAgent()
        agent = RedlineAgent(
            playbook_path=playbook_path,
            ai_provider=ai_provider,
            model=model
        )
        
        # Process Google Doc
        result = agent.process_google_doc(
            doc_id=doc_id,
            create_summary=False
        )
        
        # Prepare response
        analyses_data = []
        for analysis in result.get('analyses', []):
            # Combine assessment and comment_text into a single assessment field
            assessment = analysis.get('assessment', '')
            comment_text = analysis.get('comment_text', '')
            
            # Combine assessment and comment if both exist
            if assessment and comment_text:
                combined_assessment = f"{assessment}\n\n{comment_text}"
            elif comment_text:
                combined_assessment = comment_text
            else:
                combined_assessment = assessment
            
            analyses_data.append({
                'type': analysis['redline'].get('type', 'Unknown'),
                'text': analysis['redline'].get('text', '')[:100] + '...' if len(analysis['redline'].get('text', '')) > 100 else analysis['redline'].get('text', ''),
                'assessment': combined_assessment,
                'risk_level': analysis.get('risk_level', 'Medium'),
                'response': analysis.get('response', '')
            })
        
        return jsonify({
            'success': True,
            'redlines_count': result['redlines_count'],
            'analyses': analyses_data,
            'message': 'Comments have been added to the Google Doc'
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/download/<session_id>/<filename>')
def download_file(session_id, filename):
    """Download processed files."""
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], session_id, filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/analyze-only', methods=['POST'])
def analyze_only():
    """Analyze redlines without processing document."""
    try:
        data = request.json
        session_id = session.get('session_id')
        doc_path = session.get('doc_path')
        
        if not session_id or not doc_path:
            return jsonify({'error': 'Session expired. Please upload document again.'}), 400
        
        # Get contract type if provided
        contract_type_id = data.get('contract_type_id')
        
        # Get playbook path based on contract type
        playbook_path = get_playbook_path(contract_type_id)
        
        if not os.path.exists(doc_path):
            return jsonify({'error': 'Document not found. Please upload again.'}), 404
        
        if not playbook_path or not os.path.exists(playbook_path):
            return jsonify({'error': 'Playbook not found. Please configure a playbook in the Admin section.'}), 404
        
        # Get AI configuration
        ai_provider = data.get('ai_provider', 'openai')
        model = data.get('model', 'gpt-4')
        
        # Initialize agent
        RedlineAgent = get_RedlineAgent()
        agent = RedlineAgent(
            playbook_path=playbook_path,
            ai_provider=ai_provider,
            model=model
        )
        
        # Analyze only
        result = agent.analyze_only(input_path=doc_path)
        
        # Prepare response
        analyses_data = []
        for analysis in result.get('analyses', []):
            redline = analysis.get('redline', {})
            redline_type = redline.get('type', 'Unknown')
            
            # Handle text field - for replacements, show both old and new text
            if redline_type == 'replacement':
                old_text = redline.get('old_text', '')
                new_text = redline.get('new_text', '')
                display_text = f"'{old_text}' → '{new_text}'"
            else:
                display_text = redline.get('text', '')
            
            # Combine assessment and comment_text into a single assessment field
            assessment = analysis.get('assessment', '')
            comment_text = analysis.get('comment_text', '')
            
            # Combine assessment and comment if both exist
            if assessment and comment_text:
                combined_assessment = f"{assessment}\n\n{comment_text}"
            elif comment_text:
                combined_assessment = comment_text
            else:
                combined_assessment = assessment
            
            analyses_data.append({
                'type': redline_type,
                'text': display_text,
                'author': analysis['redline'].get('author', 'Unknown'),
                'date': analysis['redline'].get('date', 'Unknown'),
                'assessment': combined_assessment,
                'risk_level': analysis.get('risk_level', 'Medium'),
                'response': analysis.get('response', '')
            })
        
        return jsonify({
            'success': True,
            'redlines_count': result['redlines_count'],
            'analyses': analyses_data
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/admin/playbook', methods=['GET'])
def get_playbook():
    """Get current playbook content (admin)."""
    try:
        playbook_path = get_playbook_path()
        if not playbook_path or not os.path.exists(playbook_path):
            return jsonify({'error': 'Playbook not found'}), 404
        
        with open(playbook_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        return jsonify({
            'success': True,
            'content': content,
            'path': playbook_path
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/admin/playbook', methods=['POST'])
def update_playbook():
    """Update playbook content (admin)."""
    try:
        data = request.json
        content = data.get('content')
        
        if content is None:
            return jsonify({'error': 'Content is required'}), 400
        
        # Ensure playbooks directory exists
        playbooks_dir = os.path.join(os.path.dirname(__file__), 'playbooks')
        os.makedirs(playbooks_dir, exist_ok=True)
        
        # Write to standard playbook location
        with open(STANDARD_PLAYBOOK_PATH, 'w', encoding='utf-8') as f:
            f.write(content)
        
        return jsonify({
            'success': True,
            'message': 'Playbook updated successfully'
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/admin/playbook/download', methods=['GET'])
def download_playbook():
    """Download current playbook (admin)."""
    try:
        playbook_path = get_playbook_path()
        if not playbook_path or not os.path.exists(playbook_path):
            return jsonify({'error': 'Playbook not found'}), 404
        
        return send_file(playbook_path, as_attachment=True, download_name='default_playbook.txt')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/admin/playbook/upload', methods=['POST'])
def upload_playbook():
    """Upload new playbook file (admin)."""
    try:
        if 'playbook' not in request.files:
            return jsonify({'error': 'Playbook file is required'}), 400
        
        playbook_file = request.files['playbook']
        
        if playbook_file.filename == '':
            return jsonify({'error': 'File must be selected'}), 400
        
        if not playbook_file.filename.endswith('.txt'):
            return jsonify({'error': 'Only .txt files are allowed'}), 400
        
        # Ensure playbooks directory exists
        playbooks_dir = os.path.join(os.path.dirname(__file__), 'playbooks')
        os.makedirs(playbooks_dir, exist_ok=True)
        
        # Save to standard playbook location
        playbook_file.save(STANDARD_PLAYBOOK_PATH)
        
        return jsonify({
            'success': True,
            'message': 'Playbook uploaded and updated successfully'
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# Contract Types Management Endpoints
@app.route('/api/contract-types', methods=['GET'])
def get_contract_types():
    """Get all contract types."""
    try:
        ContractTypesManager = get_ContractTypesManager()
        manager = ContractTypesManager()
        types = manager.get_all_types()
        return jsonify({
            'success': True,
            'contract_types': types
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/contract-types', methods=['POST'])
def add_contract_type():
    """Add a new contract type."""
    try:
        data = request.json
        name = data.get('name')
        description = data.get('description', '')
        playbook = data.get('playbook', 'default_playbook.txt')
        
        if not name:
            return jsonify({'error': 'Contract type name is required'}), 400
        
        ContractTypesManager = get_ContractTypesManager()
        manager = ContractTypesManager()
        new_type = manager.add_type(name, description, playbook)
        
        return jsonify({
            'success': True,
            'contract_type': new_type,
            'message': 'Contract type added successfully'
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/contract-types/<type_id>', methods=['PUT'])
def update_contract_type(type_id):
    """Update an existing contract type."""
    try:
        data = request.json
        ContractTypesManager = get_ContractTypesManager()
        manager = ContractTypesManager()
        
        success = manager.update_type(
            type_id,
            name=data.get('name'),
            description=data.get('description'),
            playbook=data.get('playbook')
        )
        
        if success:
            updated_type = manager.get_type_by_id(type_id)
            return jsonify({
                'success': True,
                'contract_type': updated_type,
                'message': 'Contract type updated successfully'
            })
        else:
            return jsonify({'error': 'Contract type not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/contract-types/<type_id>', methods=['DELETE'])
def delete_contract_type(type_id):
    """Delete a contract type."""
    try:
        ContractTypesManager = get_ContractTypesManager()
        manager = ContractTypesManager()
        success = manager.delete_type(type_id)
        
        if success:
            return jsonify({
                'success': True,
                'message': 'Contract type deleted successfully'
            })
        else:
            return jsonify({'error': 'Contract type not found or cannot be deleted'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# Playbook Converter Endpoint
@app.route('/api/admin/convert-playbook', methods=['POST'])
def convert_playbook():
    """Convert Word playbook to Markdown."""
    try:
        if 'playbook' not in request.files:
            return jsonify({'error': 'Playbook file is required'}), 400
        
        playbook_file = request.files['playbook']
        
        if playbook_file.filename == '':
            return jsonify({'error': 'File must be selected'}), 400
        
        if not playbook_file.filename.endswith('.docx'):
            return jsonify({'error': 'Only .docx files are allowed'}), 400
        
        # Save uploaded file temporarily
        temp_dir = tempfile.mkdtemp()
        temp_word_path = os.path.join(temp_dir, secure_filename(playbook_file.filename))
        playbook_file.save(temp_word_path)
        
        # Convert to markdown
        PlaybookConverter = get_PlaybookConverter()
        converter = PlaybookConverter()
        base_name = os.path.splitext(playbook_file.filename)[0]
        output_path = os.path.join(temp_dir, f"{base_name}.md")
        
        markdown_path = converter.convert_word_to_markdown(temp_word_path, output_path)
        
        # Read markdown content
        with open(markdown_path, 'r', encoding='utf-8') as f:
            markdown_content = f.read()
        
        # Clean up temp files
        try:
            os.remove(temp_word_path)
            os.remove(markdown_path)
            os.rmdir(temp_dir)
        except:
            pass
        
        return jsonify({
            'success': True,
            'markdown': markdown_content,
            'filename': f"{base_name}.md"
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    # Clean up old temp files on startup
    try:
        for item in os.listdir(app.config['UPLOAD_FOLDER']):
            item_path = os.path.join(app.config['UPLOAD_FOLDER'], item)
            if os.path.isdir(item_path):
                try:
                    shutil.rmtree(item_path)
                except:
                    pass
    except:
        pass
    
    print("Starting RedLine Agent Web Interface...")
    print("Open your browser to: http://localhost:5001")
    app.run(debug=True, host='0.0.0.0', port=5001)

