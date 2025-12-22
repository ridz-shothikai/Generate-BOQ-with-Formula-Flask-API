"""
Google Cloud Storage Utility Module
Handles all GCS operations for file uploads and downloads
"""

from google.cloud import storage
import os
import tempfile
from pathlib import Path
from dotenv import load_dotenv
from datetime import timedelta

load_dotenv()

# GCS Configuration from environment variables
GCS_BUCKET_NAME = os.getenv('GCS_BUCKET_NAME')
GCS_PROJECT_ID = os.getenv('GCS_PROJECT_ID')
# GCS_CREDENTIALS_PATH = os.getenv('GCS_CREDENTIALS_PATH')

class GCSHandler:
    """Handle GCS operations for the project"""
    
    def __init__(self):
        # if GCS_CREDENTIALS_PATH and os.path.exists(GCS_CREDENTIALS_PATH):
        #     os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = GCS_CREDENTIALS_PATH
        
        self.client = storage.Client(project=GCS_PROJECT_ID)
        self.bucket = self.client.bucket(GCS_BUCKET_NAME)
    
    def upload_file(self, local_path, gcs_path):
        """Upload a file to GCS"""
        blob = self.bucket.blob(gcs_path)
        blob.upload_from_filename(local_path)
        print(f"[GCS] Uploaded {local_path} to gs://{GCS_BUCKET_NAME}/{gcs_path}")
    
    def download_file(self, gcs_path, local_path):
        """Download a file from GCS"""
        blob = self.bucket.blob(gcs_path)
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        blob.download_to_filename(local_path)
        print(f"[GCS] Downloaded gs://{GCS_BUCKET_NAME}/{gcs_path} to {local_path}")
    
    def download_to_temp(self, gcs_path, suffix=''):
        """Download a file from GCS to a temporary location"""
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        temp_path = temp_file.name
        temp_file.close()
        
        self.download_file(gcs_path, temp_path)
        return temp_path
    
    def file_exists(self, gcs_path):
        """Check if a file exists in GCS"""
        blob = self.bucket.blob(gcs_path)
        return blob.exists()
    
    def list_files(self, prefix=''):
        """List files in GCS bucket with given prefix"""
        blobs = self.bucket.list_blobs(prefix=prefix)
        return [blob.name for blob in blobs]
    
    def delete_file(self, gcs_path):
        """Delete a file from GCS"""
        blob = self.bucket.blob(gcs_path)
        blob.delete()
        print(f"[GCS] Deleted gs://{GCS_BUCKET_NAME}/{gcs_path}")
    
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
        
        Args:
            gcs_path: Path to the file in GCS
            expires_in_seconds: URL expiration time in seconds (default: 10 minutes)
            response_disposition: Content-Disposition header value (e.g., 'attachment; filename="file.xlsx"')
            response_type: Content-Type header value (e.g., 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        Returns:
            Signed URL string for direct download
        """
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
        
        print(f"[GCS] Generated signed URL for: gs://{GCS_BUCKET_NAME}/{gcs_path} (expires in {expires_in_seconds}s)")
        return url

# Singleton instance
_gcs_handler = None

def get_gcs_handler():
    """Get or create GCS handler singleton"""
    global _gcs_handler
    if _gcs_handler is None:
        _gcs_handler = GCSHandler()
    return _gcs_handler