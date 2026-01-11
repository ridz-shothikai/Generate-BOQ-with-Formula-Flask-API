"""
Google Cloud Storage Utility Module with Local Storage Fallback
Handles all GCS operations for file uploads and downloads
Falls back to local storage if GCS operations fail
"""

import os
import tempfile
from pathlib import Path
from dotenv import load_dotenv
from datetime import timedelta
import shutil

try:
    from google.cloud import storage
    GCS_AVAILABLE = True
except ImportError:
    GCS_AVAILABLE = False

load_dotenv()

# GCS Configuration from environment variables
GCS_BUCKET_NAME = os.getenv('GCS_BUCKET_NAME')
GCS_PROJECT_ID = os.getenv('GCS_PROJECT_ID')
GCS_CREDENTIALS_PATH = os.getenv('GCS_CREDENTIALS_PATH')

# Local storage configuration
LOCAL_STORAGE_DIR = Path(os.getenv('LOCAL_STORAGE_DIR', './data/sessions'))

class GCSHandler:
    """Handle GCS operations for the project with local fallback"""
    
    def __init__(self):
        self.gcs_enabled = False
        self.gcs_client = None
        self.bucket = None
        self.local_storage_dir = LOCAL_STORAGE_DIR
        self.local_storage_dir.mkdir(parents=True, exist_ok=True)
        
        # Try to initialize GCS
        if GCS_AVAILABLE and GCS_BUCKET_NAME and GCS_PROJECT_ID and GCS_CREDENTIALS_PATH:
            try:
                if os.path.exists(GCS_CREDENTIALS_PATH):
                    os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = GCS_CREDENTIALS_PATH
                
                self.gcs_client = storage.Client(project=GCS_PROJECT_ID)
                self.bucket = self.gcs_client.bucket(GCS_BUCKET_NAME)
                # Test connection
                self.bucket.exists()
                self.gcs_enabled = True
                print("[GCS] GCS connection initialized successfully")
            except Exception as e:
                print(f"[WARNING] GCS initialization failed: {str(e)}")
                print("[FALLBACK] Using local storage instead")
                self.gcs_enabled = False
        else:
            print("[INFO] GCS not configured. Using local storage.")
    
    def _get_local_path(self, gcs_path):
        """Convert GCS path to local storage path"""
        # gcs_path format: sessions/{session_id}/data/{filename} or sessions/{session_id}/output/{filename}
        local_path = self.local_storage_dir / gcs_path
        return local_path
    
    def upload_file(self, local_path, gcs_path):
        """Upload a file to GCS with fallback to local storage"""
        local_file_path = self._get_local_path(gcs_path)
        
        # First, always save to local storage (as backup)
        try:
            local_file_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(local_path, local_file_path)
            print(f"[LOCAL] Saved {local_path} to {local_file_path}")
        except Exception as e:
            print(f"[ERROR] Failed to save to local storage: {str(e)}")
            raise
        
        # Then try to upload to GCS if enabled
        if self.gcs_enabled:
            try:
                blob = self.bucket.blob(gcs_path)
                blob.upload_from_filename(local_path)
                print(f"[GCS] Uploaded {local_path} to gs://{GCS_BUCKET_NAME}/{gcs_path}")
                return {'storage': 'gcs', 'gcs_path': gcs_path, 'local_path': str(local_file_path)}
            except Exception as e:
                print(f"[WARNING] GCS upload failed: {str(e)}")
                print(f"[FALLBACK] Using local storage at {local_file_path}")
                return {'storage': 'local', 'local_path': str(local_file_path), 'gcs_path': gcs_path}
        else:
            print(f"[LOCAL] GCS not available, using local storage at {local_file_path}")
            return {'storage': 'local', 'local_path': str(local_file_path), 'gcs_path': gcs_path}
    
    def download_file(self, gcs_path, local_path):
        """Download a file from GCS with fallback to local storage"""
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        
        # First try local storage
        local_source_path = self._get_local_path(gcs_path)
        if local_source_path.exists():
            try:
                shutil.copy2(local_source_path, local_path)
                print(f"[LOCAL] Downloaded {local_source_path} to {local_path}")
                return
            except Exception as e:
                print(f"[WARNING] Failed to copy from local storage: {str(e)}")
        
        # Then try GCS if enabled
        if self.gcs_enabled:
            try:
                blob = self.bucket.blob(gcs_path)
                blob.download_to_filename(local_path)
                print(f"[GCS] Downloaded gs://{GCS_BUCKET_NAME}/{gcs_path} to {local_path}")
                return
            except Exception as e:
                print(f"[ERROR] Failed to download from GCS: {str(e)}")
                raise
        else:
            raise FileNotFoundError(f"File not found in local storage: {local_source_path}")
    
    def download_to_temp(self, gcs_path, suffix=''):
        """Download a file from GCS/local to a temporary location"""
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        temp_path = temp_file.name
        temp_file.close()
        
        self.download_file(gcs_path, temp_path)
        return temp_path
    
    def file_exists(self, gcs_path):
        """Check if a file exists in GCS or local storage"""
        # Check local storage first
        local_file_path = self._get_local_path(gcs_path)
        if local_file_path.exists():
            return True
        
        # Check GCS if enabled
        if self.gcs_enabled:
            try:
                blob = self.bucket.blob(gcs_path)
                return blob.exists()
            except Exception as e:
                print(f"[WARNING] GCS existence check failed: {str(e)}")
                return False
        
        return False
    
    def list_files(self, prefix=''):
        """List files in storage (local or GCS) with given prefix"""
        files = []
        
        # List from local storage
        try:
            local_prefix_path = self.local_storage_dir / prefix
            if local_prefix_path.exists():
                for item in local_prefix_path.rglob('*'):
                    if item.is_file():
                        relative_path = str(item.relative_to(self.local_storage_dir)).replace('\\', '/')
                        files.append(relative_path)
        except Exception as e:
            print(f"[WARNING] Failed to list local files: {str(e)}")
        
        # List from GCS if enabled
        if self.gcs_enabled:
            try:
                blobs = self.bucket.list_blobs(prefix=prefix)
                for blob in blobs:
                    if blob.name not in files:
                        files.append(blob.name)
            except Exception as e:
                print(f"[WARNING] Failed to list GCS files: {str(e)}")
        
        return files
    
    def delete_file(self, gcs_path):
        """Delete a file from GCS and local storage"""
        errors = []
        
        # Delete from local storage
        local_file_path = self._get_local_path(gcs_path)
        try:
            if local_file_path.exists():
                os.remove(local_file_path)
                print(f"[LOCAL] Deleted {local_file_path}")
        except Exception as e:
            print(f"[WARNING] Failed to delete local file: {str(e)}")
            errors.append(f"Local delete failed: {str(e)}")
        
        # Delete from GCS if enabled
        if self.gcs_enabled:
            try:
                blob = self.bucket.blob(gcs_path)
                blob.delete()
                print(f"[GCS] Deleted gs://{GCS_BUCKET_NAME}/{gcs_path}")
            except Exception as e:
                print(f"[WARNING] Failed to delete from GCS: {str(e)}")
                errors.append(f"GCS delete failed: {str(e)}")
        
        if errors:
            print(f"[INFO] Delete completed with warnings: {errors}")
    
    def get_gcs_path(self, session_id, filename, file_type='data'):
        """Generate standardized GCS path"""
        if file_type == 'data':
            return f"sessions/{session_id}/data/{filename}"
        elif file_type == 'output':
            return f"sessions/{session_id}/output/{filename}"
        else:
            return f"sessions/{session_id}/{filename}"
    
    def generate_signed_url(self, gcs_path, expires_in_seconds=600, response_disposition=None, response_type=None):
        """
        Generate a V4 signed URL for direct download from GCS
        Falls back to local file path if GCS is unavailable
        
        Args:
            gcs_path: Path to the file in GCS
            expires_in_seconds: URL expiration time in seconds (default: 10 minutes)
            response_disposition: Content-Disposition header value
            response_type: Content-Type header value
        
        Returns:
            Signed URL string for GCS or local file path
        """
        # Try GCS first if enabled
        if self.gcs_enabled:
            try:
                blob = self.bucket.blob(gcs_path)
                
                # Build query parameters for signed URL
                query_parameters = {}
                if response_disposition:
                    query_parameters['response-content-disposition'] = response_disposition
                if response_type:
                    query_parameters['response-content-type'] = response_type
                
                # Generate signed URL
                url = blob.generate_signed_url(
                    version="v4",
                    expiration=timedelta(seconds=expires_in_seconds),
                    method="GET",
                    query_parameters=query_parameters if query_parameters else None
                )
                
                print(f"[GCS] Generated signed URL for: gs://{GCS_BUCKET_NAME}/{gcs_path}")
                return url
            except Exception as e:
                print(f"[WARNING] Failed to generate GCS signed URL: {str(e)}")
        
        # Fallback to local file path
        local_file_path = self._get_local_path(gcs_path)
        if local_file_path.exists():
            print(f"[LOCAL] Using local file path: {local_file_path}")
            return f"file://{str(local_file_path)}"
        else:
            raise FileNotFoundError(f"File not found: {gcs_path}")

# Singleton instance
_gcs_handler = None

def get_gcs_handler():
    """Get or create GCS handler singleton"""
    global _gcs_handler
    if _gcs_handler is None:
        _gcs_handler = GCSHandler()
    return _gcs_handler