#!/usr/bin/env python3
"""
GURI Decompositor Module
Decomposes GURI strings into their components and maps them to real data values.
"""

import logging
from datetime import datetime
from typing import Optional, Dict, Any, List, Tuple
from guri.component_library import GURIComponentLibrary


class GURIDecompositor:
    """Decomposes GURI strings and maps components to real data values."""
    
    def __init__(self, database=None):
        """
        Initialize the decompositor.
        
        Args:
            database: GURIDatabase instance for looking up records
        """
        self.logger = logging.getLogger(__name__)
        self.database = database
        self.component_library = GURIComponentLibrary(database=database) if database else None
    
    def decompose(self, guri: str) -> Dict[str, Any]:
        """
        Decompose a GURI into its components and return structured data.
        
        Args:
            guri: GURI string to decompose
            
        Returns:
            Dictionary containing:
            - components: List of component strings
            - component_data: List of dicts with component details
            - record_data: Full record data if found in database
            - is_valid: Whether GURI format is valid
            - errors: List of error messages
        """
        result = {
            'guri': guri,
            'components': [],
            'component_data': [],
            'record_data': None,
            'is_valid': False,
            'errors': [],
            'global_type': None
        }
        
        if not guri:
            result['errors'].append("Empty GURI string")
            return result
        
        # Split by 'x'
        components = guri.split('x')
        result['components'] = components
        
        # Expected format: {5hex}x{5hex}x{8hex}x{8hex}x{3hex}x{2hex}
        expected_lengths = [5, 5, 8, 8, 3, 2]
        component_names = [
            "Component 1 (Sender)",
            "Component 2 (Recipients)",
            "Component 3 (Subject)",
            "Component 4 (DateTime)",
            "Component 5 (Risk/Location)",
            "Component 6 (Doc Type)"
        ]
        
        # Validate number of components
        if len(components) != 6:
            result['errors'].append(f"Invalid component count: expected 6, found {len(components)}")
            return result
        
        # Try to find record in database
        record_data = None
        if self.database:
            record_data = self._find_record_by_guri(guri)
            result['record_data'] = record_data
        
        # Determine Component 5 label based on document type
        doc_type_component = components[5] if len(components) > 5 else None
        if doc_type_component:
            # Get global type (define locally to avoid circular import)
            GLOBAL_TYPES = {
                "01": "Global Type 1: Email",
                "0a": "Global Type 2: General Documents",
                "0b": "Global Type 2: General Documents",
                "0c": "Global Type 2: General Documents",
                "0d": "Global Type 2: General Documents",
                "0e": "Global Type 2: General Documents",
                "0f": "Global Type 2: General Documents",
                "10": "Global Type 2: General Documents",
                "11": "Global Type 2: General Documents",
                "12": "Global Type 2: General Documents",
                "20": "Global Type 3: Aliniant Internal Documents",
                "21": "Global Type 3: Aliniant Internal Documents",
                "22": "Global Type 3: Aliniant Internal Documents",
                "23": "Global Type 3: Aliniant Internal Documents",
                "24": "Global Type 3: Aliniant Internal Documents",
                "25": "Global Type 3: Aliniant Internal Documents",
                "26": "Global Type 3: Aliniant Internal Documents",
                "27": "Global Type 3: Aliniant Internal Documents",
                "2c": "Global Type 3: Aliniant Internal Documents",
                "2a": "Global Type 4: Sensitive Documents",
                "2b": "Global Type 4: Sensitive Documents",
                "29": "Global Type 4: Sensitive Documents",
                "28": "Global Type 4: Sensitive Documents",
                "99": "Global Type 2: General Documents"
            }
            result['global_type'] = GLOBAL_TYPES.get(doc_type_component, "Unknown")
            
            # Update Component 5 label
            if doc_type_component == "01":  # Email
                component_names[4] = "Component 5 (Risk)"
            else:
                component_names[4] = "Component 5 (Location)"
        
        # Process each component
        all_valid = True
        all_hex_valid = True
        
        for i, (component, expected_len, name) in enumerate(zip(components, expected_lengths, component_names)):
            actual_len = len(component)
            is_length_valid = actual_len == expected_len
            is_hex = all(c in '0123456789abcdefABCDEF' for c in component)
            
            if not is_length_valid:
                all_valid = False
            if not is_hex:
                all_hex_valid = False
            
            # Get real value if record found or from component library
            real_value = None
            decoded_value = None
            
            # First try to get from record data
            if record_data:
                real_value = self._get_real_value_for_component(i, record_data)
            # If not found, try component library
            elif self.component_library:
                real_value = self.component_library.get_component_value(i, component)
            
            # Decode Component 4 (datetime) from hex
            if i == 3:  # Component 4 (0-indexed as 3)
                decoded_value = self._decode_datetime_component(component)
            
            # Decode Component 6 (document type)
            if i == 5:  # Component 6 (0-indexed as 5)
                # Define document types locally to avoid circular import
                DOCUMENT_TYPES = {
                    "01": "Email",
                    "0a": "Word Document",
                    "0b": "Excel Spreadsheet",
                    "0c": "PowerPoint Presentation",
                    "0d": "PDF Document",
                    "0e": "Text File",
                    "0f": "Image File",
                    "10": "Video File",
                    "11": "Audio File",
                    "12": "Archive/Zip",
                    "20": "Aliniant Policy Document",
                    "21": "Aliniant Technical Document - .docx",
                    "22": "Aliniant Technical Document - .pptx",
                    "23": "Aliniant Compliance Document",
                    "24": "Aliniant Pre-sales Document",
                    "25": "Aliniant Contract Document",
                    "26": "Legal Filing",
                    "27": "Aliniant Finance - Invoice Out",
                    "2c": "Aliniant Financial - Invoice In",
                    "2a": "Protocol A",
                    "2b": "Protocol B",
                    "29": "Sensitive Correspondence",
                    "28": "Sensitive Internal",
                    "99": "Other"
                }
                decoded_value = DOCUMENT_TYPES.get(component, f"Unknown ({component})")
            
            component_info = {
                'index': i + 1,
                'name': name,
                'component': component,
                'expected_length': expected_len,
                'actual_length': actual_len,
                'is_length_valid': is_length_valid,
                'is_hex': is_hex,
                'real_value': real_value,
                'decoded_value': decoded_value,
                'status': self._get_component_status(is_length_valid, is_hex)
            }
            
            result['component_data'].append(component_info)
        
        # Overall validation
        result['is_valid'] = all_valid and all_hex_valid
        
        if not all_valid:
            result['errors'].append("Component length mismatch")
        if not all_hex_valid:
            result['errors'].append("Non-hexadecimal characters detected")
        
        return result
    
    def _find_record_by_guri(self, guri: str) -> Optional[Dict[str, Any]]:
        """
        Find a record in the database by GURI.
        
        Args:
            guri: GURI string to search for
            
        Returns:
            Record dictionary or None if not found
        """
        if not self.database:
            return None
        
        try:
            # Use the database method if available
            if hasattr(self.database, 'get_record_by_guri'):
                return self.database.get_record_by_guri(guri)
            
            # Fallback to direct query
            conn = self.database._get_connection()
            cursor = conn.cursor()
            
            if self.database.db_type in {"postgres", "mysql"}:
                query = '''
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, 
                           document_type, sensitivity, created_at
                    FROM guri_records
                    WHERE guri = %s
                    LIMIT 1
                '''
            else:  # sqlite
                query = '''
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, 
                           document_type, sensitivity, created_at
                    FROM guri_records
                    WHERE guri = ?
                    LIMIT 1
                '''
            
            cursor.execute(query, (guri,))
            row = cursor.fetchone()
            cursor.close()
            conn.close()
            
            if row:
                return {
                    'id': row[0],
                    'guri': row[1],
                    'sender': row[2],
                    'recipients': row[3],
                    'subject': row[4],
                    'datetime': row[5],
                    'avg_risk': row[6],
                    'document_type': row[7],
                    'sensitivity': row[8] if len(row) > 8 else '[SEC1:(U)EXTERNAL/UNRATED]',
                    'created_at': row[9] if len(row) > 9 else row[8]
                }
        except Exception as e:
            self.logger.error(f"Error finding record by GURI: {e}")
        
        return None
    
    def _get_real_value_for_component(self, component_index: int, record_data: Dict[str, Any]) -> Optional[str]:
        """
        Get the real value for a component from record data.
        
        Args:
            component_index: Component index (0-5)
            record_data: Record data dictionary
            
        Returns:
            Real value string or None
        """
        mapping = {
            0: 'sender',
            1: 'recipients',
            2: 'subject',
            3: 'datetime',
            4: 'avg_risk',
            5: 'document_type'
        }
        
        field = mapping.get(component_index)
        if field and field in record_data:
            value = record_data[field]
            
            # For subject, truncate if too long
            if field == 'subject' and value and len(value) > 100:
                return value[:100] + "..."
            
            return str(value) if value else None
        
        return None
    
    def _decode_datetime_component(self, hex_val: str) -> str:
        """
        Decode Component 4 (datetime) from hex to readable time.
        
        Args:
            hex_val: 8-character hex string representing seconds since midnight
            
        Returns:
            Decoded time string
        """
        try:
            seconds = int(hex_val, 16)
            hours = seconds // 3600
            minutes = (seconds % 3600) // 60
            secs = seconds % 60
            return f"{hours:02d}:{minutes:02d}:{secs:02d} ({seconds:,} seconds since midnight)"
        except Exception as e:
            self.logger.warning(f"Error decoding datetime component '{hex_val}': {e}")
            return f"{hex_val} (hex - decode error)"
    
    def _get_component_status(self, is_length_valid: bool, is_hex: bool) -> str:
        """
        Get status string for a component.
        
        Args:
            is_length_valid: Whether length is correct
            is_hex: Whether contains only hex characters
            
        Returns:
            Status string
        """
        if is_length_valid and is_hex:
            return "Valid"
        elif not is_hex:
            return "Invalid Hex"
        else:
            return "Wrong Length"
    
    def get_component_mapping(self, guri: str) -> Dict[str, str]:
        """
        Get a simple mapping of component names to real values.
        
        Args:
            guri: GURI string to decompose
            
        Returns:
            Dictionary mapping component names to real values
        """
        result = self.decompose(guri)
        mapping = {}
        
        for comp_data in result['component_data']:
            name = comp_data['name']
            real_value = comp_data.get('real_value')
            decoded_value = comp_data.get('decoded_value')
            
            if real_value:
                mapping[name] = real_value
            elif decoded_value:
                mapping[name] = decoded_value
            else:
                mapping[name] = comp_data['component']
        
        return mapping

