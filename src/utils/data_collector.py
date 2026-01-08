"""
Data Collector Module
Collects all processing data in memory before writing to Excel
"""

import os
import json
from datetime import datetime
from pathlib import Path


class DataCollector:
    """Singleton class to collect processing data and persist to disk between subprocesses"""
    
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(DataCollector, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        """Initialize data collector"""
        if self._initialized:
            return
        
        self.reset()
        self._initialized = True
    
    def reset(self):
        """Reset all collected data"""
        self.data = {
            'tcs_schedule': None,
            'tcs_input': None,
            'emb_height': None,
            'pavement_input': None,
            'constant_fill': None,
            'formula_applier': None,
            'pavement_input_with_internal': None,
            'final_sum_applier': None,
            'metadata': {
                'session_id': None,
                'collected_at': {},
                'is_merged': True
            }
        }
        self.session_id = None
        self.output_file = None
        self.session_data_dir = None
        self.collect_dir = None
        self.manifest_path = None

    def _ensure_dirs(self):
        """Ensure session directories and manifest are initialized"""
        if self.collect_dir is None:
            # Prefer env-provided session data dir
            session_data_dir = os.getenv('SESSION_DATA_DIR')
            if session_data_dir:
                self.session_data_dir = Path(session_data_dir)
            else:
                # Fallback to current working directory
                self.session_data_dir = Path.cwd() / 'data' / 'sessions' / (self.session_id or 'default')
            self.collect_dir = self.session_data_dir / 'collected'
            self.collect_dir.mkdir(parents=True, exist_ok=True)
            self.manifest_path = self.collect_dir / 'manifest.json'
            if not self.manifest_path.exists():
                with open(self.manifest_path, 'w') as f:
                    json.dump({'scripts': {}, 'metadata': {}}, f, indent=2)
    
    def set_session_info(self, session_id, output_file, is_merged=True):
        """Set session information"""
        self.session_id = session_id
        self.output_file = output_file
        self.data['metadata']['session_id'] = session_id
        self.data['metadata']['is_merged'] = is_merged
        self._ensure_dirs()
        # Update manifest metadata
        try:
            with open(self.manifest_path, 'r') as f:
                manifest = json.load(f)
        except Exception:
            manifest = {'scripts': {}, 'metadata': {}}
        manifest['metadata'].update({
            'session_id': session_id,
            'is_merged': is_merged,
            'output_file': str(output_file),
            'updated_at': datetime.now().isoformat()
        })
        with open(self.manifest_path, 'w') as f:
            json.dump(manifest, f, indent=2)
    
    def add_data(self, script_name, data_dict):
        """
        Add data from a processing script
        
        Args:
            script_name (str): Name of the script (e.g., 'tcs_schedule', 'tcs_input')
            data_dict (dict): Dictionary containing:
                - 'sheet_name': Name of the Excel sheet to write to
                - 'dataframe': pandas DataFrame to write
                - 'start_row': Starting row (0-indexed)
                - 'start_col': Starting column (0-indexed)
                - 'write_header': Whether to include headers
                - 'write_index': Whether to include index
                OR
                - 'workbook_modifications': List of modifications to make
        """
        # Lazy init session info from env if not set
        if not self.session_id:
            self.set_session_info(
                os.getenv('SESSION_ID', 'default'),
                os.getenv('SESSION_OUTPUT_FILE', ''),
                os.getenv('IS_MERGED', 'True').lower() == 'true'
            )
        self._ensure_dirs()

        # Persist DataFrame to CSV and update manifest
        df = data_dict.get('dataframe')
        sheet_name = data_dict.get('sheet_name')
        if df is None or sheet_name is None:
            raise ValueError(f"add_data requires 'dataframe' and 'sheet_name' for {script_name}")

        data_file = self.collect_dir / f"{script_name}.csv"
        df.to_csv(data_file, index=False)

        # Update manifest
        with open(self.manifest_path, 'r') as f:
            manifest = json.load(f)
        manifest.setdefault('scripts', {})[script_name] = {
            'type': 'dataframe',
            'sheet_name': sheet_name,
            'start_row': int(data_dict.get('start_row', 0)),
            'start_col': int(data_dict.get('start_col', 0)),
            'write_header': bool(data_dict.get('write_header', False)),
            'write_index': bool(data_dict.get('write_index', False)),
            'data_file': str(data_file.name)
        }
        manifest.setdefault('metadata', {}).setdefault('collected_at', {})[script_name] = datetime.now().isoformat()
        with open(self.manifest_path, 'w') as f:
            json.dump(manifest, f, indent=2)

        print(f"[DataCollector] Stored data for {script_name} -> {data_file.name}")
    
    def add_formula_data(self, script_name, formula_data):
        """
        Add formula data from formula appliers
        
        Args:
            script_name (str): Name of the script
            formula_data (dict): Dictionary containing formula mappings
        """
        if not self.session_id:
            self.set_session_info(
                os.getenv('SESSION_ID', 'default'),
                os.getenv('SESSION_OUTPUT_FILE', ''),
                os.getenv('IS_MERGED', 'True').lower() == 'true'
            )
        self._ensure_dirs()
        formulas_file = self.collect_dir / f"{script_name}_formulas.json"
        with open(formulas_file, 'w') as f:
            json.dump(formula_data, f, indent=2)
        with open(self.manifest_path, 'r') as f:
            manifest = json.load(f)
        manifest.setdefault('scripts', {})[script_name] = {
            'type': 'formulas',
            'data_file': str(formulas_file.name)
        }
        manifest.setdefault('metadata', {}).setdefault('collected_at', {})[script_name] = datetime.now().isoformat()
        with open(self.manifest_path, 'w') as f:
            json.dump(manifest, f, indent=2)
        print(f"[DataCollector] Stored formula data for {script_name} -> {formulas_file.name}")
    
    def add_workbook_modifications(self, script_name, modifications):
        """
        Add workbook modifications (for scripts that need to modify cells directly)
        
        Args:
            script_name (str): Name of the script
            modifications (list): List of modification dictionaries containing:
                - 'sheet': Sheet name
                - 'cell' or ('row', 'col'): Cell location
                - 'value' or 'formula': Value to write
        """
        if not self.session_id:
            self.set_session_info(
                os.getenv('SESSION_ID', 'default'),
                os.getenv('SESSION_OUTPUT_FILE', ''),
                os.getenv('IS_MERGED', 'True').lower() == 'true'
            )
        self._ensure_dirs()
        mods_file = self.collect_dir / f"{script_name}_modifications.json"
        with open(mods_file, 'w') as f:
            json.dump(modifications, f, indent=2)
        with open(self.manifest_path, 'r') as f:
            manifest = json.load(f)
        manifest.setdefault('scripts', {})[script_name] = {
            'type': 'modifications',
            'data_file': str(mods_file.name)
        }
        manifest.setdefault('metadata', {}).setdefault('collected_at', {})[script_name] = datetime.now().isoformat()
        with open(self.manifest_path, 'w') as f:
            json.dump(manifest, f, indent=2)
        print(f"[DataCollector] Stored workbook modifications for {script_name} -> {mods_file.name}")
    
    def get_data(self, script_name):
        """Get data for a specific script"""
        return self.data.get(script_name)
    
    def get_all_data(self):
        """Get all collected data from manifest (disk-backed)"""
        self._ensure_dirs()
        try:
            with open(self.manifest_path, 'r') as f:
                manifest = json.load(f)
        except Exception:
            manifest = {'scripts': {}, 'metadata': {}}
        return manifest
    
    def has_data(self, script_name):
        """Check if data exists for a script"""
        return self.data.get(script_name) is not None
    
    def save_debug_info(self, output_dir):
        """Save collected data info for debugging"""
        debug_file = Path(output_dir) / f"data_collector_debug_{self.session_id}.json"
        debug_info = {
            'session_id': self.session_id,
            'collected_scripts': [k for k, v in self.data.items() if v is not None and k != 'metadata'],
            'metadata': self.data['metadata']
        }
        with open(debug_file, 'w') as f:
            json.dump(debug_info, f, indent=2, default=str)
        print(f"[DataCollector] Debug info saved to {debug_file}")


def get_collector():
    """Get the singleton data collector instance"""
    return DataCollector()
