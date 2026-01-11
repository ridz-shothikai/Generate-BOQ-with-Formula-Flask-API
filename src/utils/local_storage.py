"""
Local Storage Utility Module
Provides consistent interface for local file storage organized by session_id
"""

import os
from pathlib import Path
from dotenv import load_dotenv
import shutil

load_dotenv()

# Local storage configuration
LOCAL_STORAGE_DIR = Path(os.getenv('LOCAL_STORAGE_DIR', './data/sessions'))


class LocalStorageManager:
    """Manage local storage with session-based organization"""
    
    def __init__(self, base_dir=LOCAL_STORAGE_DIR):
        self.base_dir = Path(base_dir)
        self.base_dir.mkdir(parents=True, exist_ok=True)
    
    def get_session_data_dir(self, session_id):
        """Get or create data directory for a session"""
        data_dir = self.base_dir / session_id / 'data'
        data_dir.mkdir(parents=True, exist_ok=True)
        return data_dir
    
    def get_session_output_dir(self, session_id):
        """Get or create output directory for a session"""
        output_dir = self.base_dir / session_id / 'output'
        output_dir.mkdir(parents=True, exist_ok=True)
        return output_dir
    
    def get_session_dir(self, session_id):
        """Get or create base directory for a session"""
        session_dir = self.base_dir / session_id
        session_dir.mkdir(parents=True, exist_ok=True)
        return session_dir
    
    def save_file(self, session_id, filename, file_path, file_type='data'):
        """
        Save a file to local storage
        
        Args:
            session_id: Session identifier
            filename: Name of the file
            file_path: Path to the source file
            file_type: 'data' or 'output'
        
        Returns:
            Path to the saved file
        """
        if file_type == 'data':
            target_dir = self.get_session_data_dir(session_id)
        elif file_type == 'output':
            target_dir = self.get_session_output_dir(session_id)
        else:
            target_dir = self.get_session_dir(session_id)
        
        target_path = target_dir / filename
        
        try:
            shutil.copy2(file_path, target_path)
            print(f"[LOCAL] Saved {filename} to {target_path}")
            return target_path
        except Exception as e:
            print(f"[ERROR] Failed to save file: {str(e)}")
            raise
    
    def get_file(self, session_id, filename, file_type='data'):
        """
        Get path to a file in local storage
        
        Args:
            session_id: Session identifier
            filename: Name of the file
            file_type: 'data' or 'output'
        
        Returns:
            Path to the file if exists, None otherwise
        """
        if file_type == 'data':
            target_dir = self.get_session_data_dir(session_id)
        elif file_type == 'output':
            target_dir = self.get_session_output_dir(session_id)
        else:
            target_dir = self.get_session_dir(session_id)
        
        file_path = target_dir / filename
        return file_path if file_path.exists() else None
    
    def list_files(self, session_id, file_type='data'):
        """
        List all files in a session directory
        
        Args:
            session_id: Session identifier
            file_type: 'data' or 'output'
        
        Returns:
            List of file paths
        """
        if file_type == 'data':
            target_dir = self.get_session_data_dir(session_id)
        elif file_type == 'output':
            target_dir = self.get_session_output_dir(session_id)
        else:
            target_dir = self.get_session_dir(session_id)
        
        if not target_dir.exists():
            return []
        
        return list(target_dir.glob('*'))
    
    def delete_file(self, session_id, filename, file_type='data'):
        """
        Delete a file from local storage
        
        Args:
            session_id: Session identifier
            filename: Name of the file
            file_type: 'data' or 'output'
        
        Returns:
            True if deleted, False if not found
        """
        file_path = self.get_file(session_id, filename, file_type)
        
        if file_path and file_path.exists():
            try:
                os.remove(file_path)
                print(f"[LOCAL] Deleted {file_path}")
                return True
            except Exception as e:
                print(f"[ERROR] Failed to delete file: {str(e)}")
                raise
        
        return False
    
    def delete_session(self, session_id):
        """
        Delete entire session directory
        
        Args:
            session_id: Session identifier
        
        Returns:
            True if deleted, False if not found
        """
        session_dir = self.base_dir / session_id
        
        if session_dir.exists():
            try:
                shutil.rmtree(session_dir)
                print(f"[LOCAL] Deleted session directory: {session_dir}")
                return True
            except Exception as e:
                print(f"[ERROR] Failed to delete session directory: {str(e)}")
                raise
        
        return False
    
    def get_session_info(self, session_id):
        """
        Get information about session storage
        
        Args:
            session_id: Session identifier
        
        Returns:
            Dictionary with session storage info
        """
        session_dir = self.base_dir / session_id
        data_dir = self.get_session_data_dir(session_id)
        output_dir = self.get_session_output_dir(session_id)
        
        info = {
            'session_id': session_id,
            'base_path': str(session_dir),
            'exists': session_dir.exists(),
            'data_files': len(list(data_dir.glob('*'))) if data_dir.exists() else 0,
            'output_files': len(list(output_dir.glob('*'))) if output_dir.exists() else 0,
            'total_size_bytes': 0
        }
        
        # Calculate total size
        if session_dir.exists():
            for item in session_dir.rglob('*'):
                if item.is_file():
                    info['total_size_bytes'] += item.stat().st_size
        
        return info


# Singleton instance
_local_storage = None


def get_local_storage_manager(base_dir=None):
    """Get or create local storage manager singleton"""
    global _local_storage
    if _local_storage is None:
        _local_storage = LocalStorageManager(base_dir or LOCAL_STORAGE_DIR)
    return _local_storage
