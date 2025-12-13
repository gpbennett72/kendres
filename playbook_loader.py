"""Legal playbook loader and parser."""

from typing import List, Dict, Optional
from pathlib import Path
import re


class PlaybookLoader:
    """Loads and parses legal playbooks from text files."""
    
    def __init__(self, playbook_path: str):
        """Initialize with playbook file path."""
        self.playbook_path = Path(playbook_path)
        self.principles: List[Dict[str, str]] = []
        self.load_playbook()
    
    def load_playbook(self) -> None:
        """Load and parse the playbook file."""
        if not self.playbook_path.exists():
            raise FileNotFoundError(f"Playbook not found: {self.playbook_path}")
        
        with open(self.playbook_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        self._parse_playbook(content)
    
    def _parse_playbook(self, content: str) -> None:
        """Parse playbook content into structured principles."""
        # Support multiple formats:
        # 1. PRINCIPLE: ... RESPONSE: ...
        # 2. Section headers with bullet points
        # 3. Simple text guidelines
        
        lines = content.split('\n')
        current_principle = None
        current_response = None
        
        for line in lines:
            line = line.strip()
            if not line:
                if current_principle:
                    self.principles.append({
                        'principle': current_principle,
                        'response': current_response or ''
                    })
                    current_principle = None
                    current_response = None
                continue
            
            # Check for PRINCIPLE: pattern
            if line.upper().startswith('PRINCIPLE:'):
                if current_principle:
                    self.principles.append({
                        'principle': current_principle,
                        'response': current_response or ''
                    })
                current_principle = line[10:].strip()
                current_response = None
            elif line.upper().startswith('RESPONSE:'):
                current_response = line[9:].strip()
            elif current_principle and not current_response:
                # Continuation of principle
                current_principle += ' ' + line
            elif current_response is not None:
                # Continuation of response
                current_response += ' ' + line
        
        # Add last principle if exists
        if current_principle:
            self.principles.append({
                'principle': current_principle,
                'response': current_response or ''
            })
        
        # If no structured principles found, treat entire content as playbook
        if not self.principles:
            self.principles.append({
                'principle': 'General Legal Guidelines',
                'response': content
            })
    
    def get_playbook_text(self) -> str:
        """Get full playbook text for AI context."""
        text_parts = []
        for item in self.principles:
            text_parts.append(f"PRINCIPLE: {item['principle']}")
            if item['response']:
                text_parts.append(f"RESPONSE: {item['response']}")
        return '\n\n'.join(text_parts)
    
    def get_principles(self) -> List[Dict[str, str]]:
        """Get list of principles."""
        return self.principles








