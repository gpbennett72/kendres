"""Extract redlines from Google Docs."""

from typing import List, Dict, Optional
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
import re


# Scopes required for Google Docs API
SCOPES = ['https://www.googleapis.com/auth/documents.readonly',
          'https://www.googleapis.com/auth/drive.readonly',
          'https://www.googleapis.com/auth/documents']


class GoogleDocsRedlineExtractor:
    """Extracts tracked changes and revisions from Google Docs."""
    
    def __init__(self, doc_id: str, credentials_path: Optional[str] = None):
        """Initialize with Google Doc ID and credentials path."""
        self.doc_id = doc_id
        self.credentials_path = credentials_path or os.getenv('GOOGLE_CREDENTIALS_PATH', 'credentials.json')
        self.service = self._authenticate()
        self.redlines: List[Dict] = []
        self._extract_redlines()
    
    def _authenticate(self):
        """Authenticate and return Google Docs service."""
        creds = None
        
        # Check for existing token
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        
        # If no valid credentials, get new ones
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(self.credentials_path):
                    raise FileNotFoundError(
                        f"Google credentials not found at {self.credentials_path}. "
                        "Please download OAuth2 credentials from Google Cloud Console."
                    )
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_path, SCOPES)
                creds = flow.run_local_server(port=0)
            
            # Save credentials for next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        
        return build('docs', 'v1', credentials=creds)
    
    def _extract_redlines(self) -> None:
        """Extract tracked changes from Google Doc."""
        # Get document revisions
        drive_service = build('drive', 'v3', credentials=self.service._http.credentials)
        
        try:
            # Get revisions
            revisions = drive_service.revisions().list(fileId=self.doc_id).execute()
            
            # Get current document
            doc = self.service.documents().get(documentId=self.doc_id).execute()
            
            # Extract suggested edits (comments and suggestions)
            self._extract_suggestions(doc)
            
            # Compare revisions if available
            if revisions.get('revisions'):
                self._extract_revision_changes(revisions, doc)
        
        except Exception as e:
            # Fallback: extract from document content and comments
            doc = self.service.documents().get(documentId=self.doc_id).execute()
            self._extract_suggestions(doc)
            self._extract_comments(doc)
    
    def _extract_suggestions(self, doc: Dict) -> None:
        """Extract suggested edits from document."""
        # Google Docs stores suggestions in the document structure
        # Look for suggested insertions and deletions
        
        if 'body' in doc and 'content' in doc['body']:
            for element in doc['body']['content']:
                if 'paragraph' in element:
                    self._process_paragraph_suggestions(element['paragraph'])
                elif 'table' in element:
                    self._process_table_suggestions(element['table'])
    
    def _process_paragraph_suggestions(self, paragraph: Dict) -> None:
        """Process suggestions in a paragraph."""
        if 'elements' not in paragraph:
            return
        
        for elem in paragraph['elements']:
            if 'textRun' in elem:
                text_run = elem['textRun']
                text = text_run.get('content', '')
                
                # Check for suggested changes
                if 'suggestedInsertion' in text_run:
                    self.redlines.append({
                        'type': 'suggested_insertion',
                        'text': text,
                        'author': text_run['suggestedInsertion'].get('author', {}).get('displayName', 'Unknown'),
                        'date': text_run['suggestedInsertion'].get('date', ''),
                        'suggestion_id': text_run['suggestedInsertion'].get('suggestionId', '')
                    })
                
                if 'suggestedDeletion' in text_run:
                    self.redlines.append({
                        'type': 'suggested_deletion',
                        'text': text,
                        'author': text_run['suggestedDeletion'].get('author', {}).get('displayName', 'Unknown'),
                        'date': text_run['suggestedDeletion'].get('date', ''),
                        'suggestion_id': text_run['suggestedDeletion'].get('suggestionId', '')
                    })
    
    def _process_table_suggestions(self, table: Dict) -> None:
        """Process suggestions in a table."""
        if 'tableRows' not in table:
            return
        
        for row in table['tableRows']:
            if 'tableCells' not in row:
                continue
            for cell in row['tableCells']:
                if 'content' in cell:
                    for element in cell['content']:
                        if 'paragraph' in element:
                            self._process_paragraph_suggestions(element['paragraph'])
    
    def _extract_revision_changes(self, revisions: Dict, current_doc: Dict) -> None:
        """Extract changes by comparing revisions."""
        # This is a simplified version - full implementation would compare document states
        for rev in revisions.get('revisions', [])[:10]:  # Limit to last 10 revisions
            if rev.get('published', False):
                continue  # Skip published revisions
            
            self.redlines.append({
                'type': 'revision',
                'text': f"Revision by {rev.get('lastModifyingUser', {}).get('displayName', 'Unknown')}",
                'author': rev.get('lastModifyingUser', {}).get('displayName', 'Unknown'),
                'date': rev.get('modifiedTime', ''),
                'revision_id': rev.get('id', '')
            })
    
    def _extract_comments(self, doc: Dict) -> None:
        """Extract comments from document."""
        # Comments are stored separately and can indicate areas needing review
        # This would require additional API calls to get comments
        pass
    
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
        doc = self.service.documents().get(documentId=self.doc_id).execute()
        text_parts = []
        
        if 'body' in doc and 'content' in doc['body']:
            for element in doc['body']['content']:
                if 'paragraph' in element:
                    para = element['paragraph']
                    if 'elements' in para:
                        for elem in para['elements']:
                            if 'textRun' in elem:
                                text_parts.append(elem['textRun'].get('content', ''))
        
        return ''.join(text_parts)








