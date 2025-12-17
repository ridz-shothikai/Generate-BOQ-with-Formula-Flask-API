"""
Session Manager for MongoDB
"""
from pymongo import MongoClient
import os
from datetime import datetime
import traceback
from dotenv import load_dotenv
import sys
from pathlib import Path

# Add project root to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / 'src' / 'utils'))
from gcs_utils import get_gcs_handler



class SessionManager:
    def __init__(self):
        load_dotenv()
        
        self.client = MongoClient(os.getenv('MONGO_URI'))
        self.db = self.client[os.getenv('MONGO_DB_NAME')]
        self.sessions = self.db[os.getenv('SESSION_COLLECTION', 'sessions')]
    
    def generate_session_id(self):
        """Generate session ID in format YYYYMMDD_HHMMSS"""
        return datetime.now().strftime("%Y%m%d_%H%M%S")
    
    def create_session(self, session_id=None):
        """Create new session - generates session_id if not provided"""
        if session_id is None:
            session_id = self.generate_session_id()
        
        session_data = {
            'session_id': session_id,
            'status': 'uploaded',
            'created_at': datetime.now(),
            'updated_at': datetime.now(),
            'input_files': [],  # This will store file info objects
            'uploaded_files': {},  # This stores the categorized files
            'output_file': None,
            'processing_info': {
                'started_at': None,
                'completed_at': None,
                'execution_time_seconds': None
            },
            'error_info': {
                'has_error': False,
                'error_message': None,
                'error_traceback': None,
                'failed_script': None
            },
            'zip_download_count': 0,
            'files_generated': {
                'main_carriageway_and_boq': f"{session_id}_main_carriageway_and_boq.xlsx"
            },
            'metadata': {},
            'progress': {}
        }
        self.sessions.insert_one(session_data)
        return session_data
    
    def update_session_status(self, session_id, status, **updates):
        """Update session status and other fields"""
        update_data = {
            'status': status,
            'updated_at': datetime.now(),
            **updates
        }
        self.sessions.update_one(
            {'session_id': session_id},
            {'$set': update_data}
        )
    
    def update_progress(self, session_id, progress_data):
        """Update processing progress"""
        self.sessions.update_one(
            {'session_id': session_id},
            {'$set': {
                'progress': progress_data,
                'updated_at': datetime.now()
            }}
        )
    
    def add_input_file(self, session_id, file_info):
        """Add input file info to session"""
        self.sessions.update_one(
            {'session_id': session_id},
            {'$push': {'input_files': file_info}, '$set': {'updated_at': datetime.now()}}
        )
    
    def set_error(self, session_id, error_message, traceback_str, failed_script):
        """Set error information for session"""
        self.sessions.update_one(
            {'session_id': session_id},
            {'$set': {
                'status': 'failed',
                'updated_at': datetime.now(),
                'error_info': {
                    'has_error': True,
                    'error_message': error_message,
                    'error_traceback': traceback_str,
                    'failed_script': failed_script
                }
            }}
        )
    
    def set_output_file(self, session_id, output_info):
        """Set output file information"""
        self.sessions.update_one(
            {'session_id': session_id},
            {'$set': {
                'output_file': output_info,
                'status': 'completed',
                'updated_at': datetime.now(),
                'processing_info.completed_at': datetime.now()
            }}
        )
    
    def set_boq_file(self, session_id, boq_info):
        """Set BOQ file information"""
        self.sessions.update_one(
            {'session_id': session_id},
            {'$set': {
                'boq_file': boq_info,
                'updated_at': datetime.now()
            }}
        )
    
    def set_processing_started(self, session_id):
        """Set processing start time"""
        self.sessions.update_one(
            {'session_id': session_id},
            {'$set': {
                'processing_info.started_at': datetime.now(),
                'updated_at': datetime.now()
            }}
        )
    
    def get_session(self, session_id):
        """Get session by ID"""
        return self.sessions.find_one({'session_id': session_id})
    
    def session_exists(self, session_id):
        """Check if session exists"""
        return self.sessions.find_one({'session_id': session_id}) is not None
    
    def update_session_data(self, session_id, data):
        """Update session with custom data (for GCS paths, etc.)"""
        self.sessions.update_one(
            {'session_id': session_id},
            {'$set': {
                **data,
                'updated_at': datetime.now()
            }}
        )
    
    def delete_session(self, session_id):
        """Delete a session (for cleanup)"""
        result = self.sessions.delete_one({'session_id': session_id})
        return result.deleted_count > 0
    
    def get_all_sessions(self, limit=10, skip=0, sort_by='created_at', sort_order=-1):
        """
        Get all sessions with pagination
        
        Args:
            limit: Number of sessions to return
            skip: Number of sessions to skip
            sort_by: Field to sort by (default: created_at)
            sort_order: 1 for ascending, -1 for descending
        
        Returns:
            List of sessions and total count
        """
        sessions_cursor = self.sessions.find().sort(sort_by, sort_order).skip(skip).limit(limit)
        
        # Initialize GCS handler for generating URLs
        gcs = get_gcs_handler()
        
        sessions = []
        for session in sessions_cursor:
            # Count files from uploaded_files instead of input_files
            uploaded_files_count = len(session.get('uploaded_files', {}))
            input_files_count = len(session.get('input_files', []))
            
            # Use uploaded_files count if available, otherwise fall back to input_files
            files_count = uploaded_files_count if uploaded_files_count > 0 else input_files_count
            
            # Process uploaded files to include GCS URLs
            uploaded_files_with_urls = {}
            if session.get('uploaded_files'):
                for file_key, file_info in session['uploaded_files'].items():
                    file_data = dict(file_info)  # Copy the file info
                    if 'gcs_path' in file_info:
                        try:
                            # Generate signed URL for the file
                            signed_url = gcs.generate_signed_url(
                                file_info['gcs_path'],
                                expires_in_seconds=3600,  # 1 hour
                                response_disposition=f'attachment; filename="{file_info.get("filename", "file.xlsx")}"',
                                response_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
                            file_data['gcs_url'] = signed_url
                        except Exception as e:
                            # If URL generation fails, continue without URL
                            file_data['gcs_url'] = None
                    uploaded_files_with_urls[file_key] = file_data
            
            # Convert MongoDB document to JSON-serializable format
            session_data = {
                'session_id': session['session_id'],
                'status': session['status'],
                'created_at': session['created_at'].isoformat() if session.get('created_at') else None,
                'updated_at': session['updated_at'].isoformat() if session.get('updated_at') else None,
                'input_files_count': files_count,  # Use the correct count
                'has_error': session.get('error_info', {}).get('has_error', False),
                'error_message': session.get('error_info', {}).get('error_message'),
                'output_file': session.get('output_file'),
                'boq_file': session.get('boq_file'),
                'zip_download_count': session.get('zip_download_count', 0),
                'uploaded_files': uploaded_files_with_urls  # Use the processed files with URLs
            }
            
            # Add processing info if available
            if session.get('processing_info'):
                session_data['processing_info'] = {
                    'started_at': session['processing_info'].get('started_at').isoformat() if session['processing_info'].get('started_at') else None,
                    'completed_at': session['processing_info'].get('completed_at').isoformat() if session['processing_info'].get('completed_at') else None,
                    'execution_time_seconds': session['processing_info'].get('execution_time_seconds')
                }
            
            sessions.append(session_data)
        
        return sessions, self.sessions.count_documents({})
    
    def increment_download_count(self, session_id):
        """Increment ZIP download count"""
        self.sessions.update_one(
            {'session_id': session_id},
            {
                '$inc': {'zip_download_count': 1},
                '$set': {'updated_at': datetime.now()}
            }
        )