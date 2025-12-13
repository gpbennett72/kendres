"""Contract types management system."""

import os
import json
from typing import List, Dict, Optional
from pathlib import Path


class ContractTypesManager:
    """Manages contract types and their associated playbooks."""
    
    def __init__(self, config_path: Optional[str] = None):
        """Initialize with config file path."""
        if config_path is None:
            config_path = os.path.join(os.path.dirname(__file__), 'contract_types.json')
        self.config_path = Path(config_path)
        self.contract_types: List[Dict] = []
        self.load_config()
    
    def load_config(self) -> None:
        """Load contract types from JSON file."""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.contract_types = data.get('contract_types', [])
            except (json.JSONDecodeError, IOError) as e:
                print(f"Warning: Could not load contract types config: {e}")
                self.contract_types = []
        else:
            # Initialize with default contract types
            self.contract_types = [
                {
                    'id': 'default',
                    'name': 'Default',
                    'description': 'Default contract type',
                    'playbook': 'default_playbook.txt'
                }
            ]
            self.save_config()
    
    def save_config(self) -> bool:
        """Save contract types to JSON file."""
        try:
            data = {
                'contract_types': self.contract_types
            }
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            return True
        except IOError as e:
            print(f"Error saving contract types config: {e}")
            return False
    
    def get_all_types(self) -> List[Dict]:
        """Get all contract types."""
        return self.contract_types
    
    def get_type_by_id(self, type_id: str) -> Optional[Dict]:
        """Get contract type by ID."""
        for ct in self.contract_types:
            if ct.get('id') == type_id:
                return ct
        return None
    
    def add_type(self, name: str, description: str = '', playbook: str = 'default_playbook.txt') -> Dict:
        """Add a new contract type."""
        # Generate ID from name
        type_id = name.lower().replace(' ', '_').replace('-', '_')
        # Ensure unique ID
        existing_ids = [ct.get('id') for ct in self.contract_types]
        counter = 1
        original_id = type_id
        while type_id in existing_ids:
            type_id = f"{original_id}_{counter}"
            counter += 1
        
        new_type = {
            'id': type_id,
            'name': name,
            'description': description,
            'playbook': playbook
        }
        self.contract_types.append(new_type)
        self.save_config()
        return new_type
    
    def update_type(self, type_id: str, name: str = None, description: str = None, playbook: str = None) -> bool:
        """Update an existing contract type."""
        for ct in self.contract_types:
            if ct.get('id') == type_id:
                if name is not None:
                    ct['name'] = name
                if description is not None:
                    ct['description'] = description
                if playbook is not None:
                    ct['playbook'] = playbook
                self.save_config()
                return True
        return False
    
    def delete_type(self, type_id: str) -> bool:
        """Delete a contract type."""
        if type_id == 'default':
            return False  # Cannot delete default type
        
        original_count = len(self.contract_types)
        self.contract_types = [ct for ct in self.contract_types if ct.get('id') != type_id]
        
        if len(self.contract_types) < original_count:
            self.save_config()
            return True
        return False
    
    def get_playbook_path(self, type_id: str) -> Optional[str]:
        """Get playbook path for a contract type."""
        contract_type = self.get_type_by_id(type_id)
        if not contract_type:
            return None
        
        playbook_name = contract_type.get('playbook', 'default_playbook.txt')
        playbooks_dir = os.path.join(os.path.dirname(__file__), 'playbooks')
        playbook_path = os.path.join(playbooks_dir, playbook_name)
        
        if os.path.exists(playbook_path):
            return playbook_path
        
        # Fallback to default
        default_path = os.path.join(playbooks_dir, 'default_playbook.txt')
        if os.path.exists(default_path):
            return default_path
        
        return None





