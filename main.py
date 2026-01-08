"""
Flask API for BOQ Generation with Session Management
"""
import os
import sys
import subprocess
import threading
from pathlib import Path
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import traceback
import shutil
from datetime import datetime, timezone
import zipfile
from dotenv import load_dotenv
import tempfile
# Import session manager
from api.session_manager import SessionManager

# Import GCS handler
# sys.path.insert(0, str(Path(__file__).parent / 'src' / 'internal'))
from src.utils.gcs_utils import get_gcs_handler

load_dotenv()
# Initialize Flask app
app = Flask(__name__)
CORS(app)

# Configuration
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100 MB per file

# Get project root directory
PROJECT_ROOT = Path(__file__).parent
DATA_DIR = PROJECT_ROOT / 'data'
OUTPUT_DIR = PROJECT_ROOT / 'output'
TEMPLATE_DIR = PROJECT_ROOT / 'template'
SRC_DIR = PROJECT_ROOT / 'src'

# Required file names
REQUIRED_FILES = {
    'tcs_schedule': 'TCS Schedule.xlsx',
    'tcs_input': 'TCS Input.xlsx',
    'emb_height': 'Emb Height.xlsx',
    'pavement_input': 'Pavement Input.xlsx'
}

# Template files
TEMPLATE_FILE = TEMPLATE_DIR / 'main_carriageway_and_boq.xlsx'
TEMPLATE_FILE_SINGLE = TEMPLATE_DIR / 'main_carriageway.xlsx'

# Sequential script path
SEQUENTIAL_SCRIPT = SRC_DIR / 'sequential.py'

PORT = os.getenv('PORT', 5000)

def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def ensure_directories():
    """Ensure required directories exist"""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)

def validate_template():
    """Validate that template files exist"""
    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError(f"Template file not found: {TEMPLATE_FILE}")
    if not TEMPLATE_FILE_SINGLE.exists():
        raise FileNotFoundError(f"Template file not found: {TEMPLATE_FILE_SINGLE}")

def test_gcs_connection():
    """
    Test GCS connection and permissions
    Returns: (success: bool, message: str)
    """
    try:
        print("\n" + "="*80)
        print("TESTING GCS CONNECTION")
        print("="*80)
        
        # Get GCS configuration from environment
        bucket_name = os.getenv('GCS_BUCKET_NAME')
        project_id = os.getenv('GCS_PROJECT_ID')
        credentials_path = os.getenv('GCS_CREDENTIALS_PATH')
        
        # Check if environment variables are set
        if not bucket_name:
            return False, "GCS_BUCKET_NAME not set in environment variables"
        if not project_id:
            return False, "GCS_PROJECT_ID not set in environment variables"
        if not credentials_path:
            return False, "GCS_CREDENTIALS_PATH not set in environment variables"
        
        print(f"✓ Bucket Name: {bucket_name}")
        print(f"✓ Project ID: {project_id}")
        print(f"✓ Credentials Path: {credentials_path}")
        
        # Check if credentials file exists
        if not os.path.exists(credentials_path):
            return False, f"Credentials file not found at: {credentials_path}"
        print(f"✓ Credentials file exists")
        
        # Initialize GCS handler
        print("\n[TEST 1] Initializing GCS client...")
        gcs = get_gcs_handler()
        print("✓ GCS client initialized successfully")
        
        # Test 2: Check if bucket exists and is accessible
        print("\n[TEST 2] Checking bucket access...")
        if gcs.bucket.exists():
            print(f"✓ SUCCESS: Connected to bucket '{bucket_name}'")
        else:
            return False, f"Bucket '{bucket_name}' does not exist or is not accessible"
        
        # Test 3: Check read permissions
        print("\n[TEST 3] Checking read permissions...")
        try:
            blobs = list(gcs.bucket.list_blobs(max_results=1))
            print(f"✓ SUCCESS: Can read from bucket")
        except Exception as e:
            return False, f"Cannot read from bucket: {str(e)}"
        
        # Test 4: Check write permissions (create and delete a test file)
        print("\n[TEST 4] Checking write permissions...")
        try:
            test_file_path = '_test_connection.txt'
            test_blob = gcs.bucket.blob(test_file_path)
            test_content = f'Connection test at {datetime.now(timezone.utc).isoformat()}'
            test_blob.upload_from_string(test_content)
            print(f"✓ SUCCESS: Can write to bucket")
            
            # Clean up test file
            test_blob.delete()
            print(f"✓ SUCCESS: Can delete from bucket")
        except Exception as e:
            return False, f"Cannot write/delete from bucket: {str(e)}"
        
        # Test 5: Check session directory structure
        print("\n[TEST 5] Checking session directory structure...")
        try:
            test_session_id = '_test_session'
            test_paths = [
                f'sessions/{test_session_id}/data/test.txt',
                f'sessions/{test_session_id}/output/test.txt'
            ]
            
            for test_path in test_paths:
                blob = gcs.bucket.blob(test_path)
                blob.upload_from_string('test')
                blob.delete()
            
            print(f"✓ SUCCESS: Can create/delete in session directories")
        except Exception as e:
            return False, f"Cannot manage session directories: {str(e)}"
        
        print("\n" + "="*80)
        print("GCS CONNECTION TEST: ALL TESTS PASSED ✓")
        print("="*80)
        
        return True, "GCS connection successful"
        
    except Exception as e:
        error_msg = f"GCS connection test failed: {str(e)}"
        print(f"\n✗ ERROR: {error_msg}")
        print(traceback.format_exc())
        return False, error_msg

def run_session_processing(session_id, session_data_dir, session_output_file, is_merged=True):
    """
    Run the sequential processing for a specific session with progress tracking
    
    Args:
        session_id: Session identifier
        session_data_dir: Directory containing session data files
        session_output_file: Output file path
        is_merged: If True, process merged file (main_carriageway_and_boq) with BOQ. 
                   If False, process single file (main_carriageway) without BOQ.
    """
    session_manager = SessionManager()
    
    try:
        # Update session status to processing IMMEDIATELY
        session_manager.update_session_status(session_id, 'processing')
        session_manager.set_processing_started(session_id)
        
        # Define processing steps with user-friendly names and timeouts
        processing_steps = [
            {'script': 'tcs_schedule', 'name': 'TCS Schedule Processing', 'message': 'Processing TCS schedule data...', 'timeout': 600},
            {'script': 'tcs_input', 'name': 'TCS Input Processing', 'message': 'Processing TCS input specifications...', 'timeout': 600},
            {'script': 'emb_height', 'name': 'Embankment Height Processing', 'message': 'Processing embankment height data...', 'timeout': 600},
            {'script': 'pavement_input', 'name': 'Pavement Input Processing', 'message': 'Processing pavement layer specifications...', 'timeout': 600},
            {'script': 'constant_fill', 'name': 'Constant Values Processing', 'message': 'Applying constant values...', 'timeout': 600},
            {'script': 'formula_applier', 'name': 'Formula Application', 'message': 'Applying calculation formulas...', 'timeout': 600},
            {'script': 'pavement_input_with_internal', 'name': 'Geogrid Calculation', 'message': 'Calculating geogrid requirements...', 'timeout': 600},
            {'script': 'final_sum_applier', 'name': 'Final Summary', 'message': 'Generating final summary...', 'timeout': 600},
            # {'script': 'calculator', 'name': 'Formula Calculation', 'message': 'Calculating formula values...'},
        ]
        
        # Always write collected data once after calculations
        processing_steps.append({'script': 'collected_writer', 'name': 'Writing Collected Data', 'message': 'Writing all collected data to Excel...', 'timeout': 600})
        # Only add BOQ generation for merged flow
        if is_merged:
            processing_steps.append({'script': 'boq_populator', 'name': 'BOQ Generation', 'message': 'Generating BOQ template...', 'timeout': 600})
        
        total_steps = len(processing_steps)
        
        # Initialize progress - starting first step (0% complete)
        session_manager.update_progress(session_id, {
            'current_step': processing_steps[0]['name'],
            'current_step_number': 1,
            'total_steps': total_steps,
            'percentage': 0,
            'message': processing_steps[0]['message'],
            'completed_steps': 0,
            'started_at': datetime.now(timezone.utc).isoformat()
        })
        
        # Change to project root to ensure relative imports work
        original_cwd = os.getcwd()
        
        try:
            os.chdir(PROJECT_ROOT)
            
            # Set environment variables for session-specific paths
            env = os.environ.copy()
            env['SESSION_DATA_DIR'] = str(session_data_dir)
            env['SESSION_OUTPUT_FILE'] = str(session_output_file)
            env['SESSION_ID'] = session_id
            env['IS_MERGED'] = 'True' if is_merged else 'False'
            
            # Run each script sequentially with proper progress tracking
            for step_num, step_info in enumerate(processing_steps, 1):
                script_name = step_info['script']
                step_name = step_info['name']
                step_message = step_info['message']
                step_timeout = step_info.get('timeout', 300)  # Default to 5 minutes if not specified
                
                print(f"Executing step {step_num}/{total_steps}: {step_name} (timeout: {step_timeout}s)")
                
                # Update progress to show current step (before execution)
                if step_num > 1:  # For steps 2-8, update current step
                    session_manager.update_progress(session_id, {
                        'current_step': step_name,
                        'current_step_number': step_num,
                        'total_steps': total_steps,
                        'percentage': int(((step_num - 1) / total_steps) * 100),  # Previous step completed
                        'message': step_message,
                        'completed_steps': step_num - 1,
                        'last_completed_at': datetime.now(timezone.utc).isoformat()
                    })
                
                # Determine script path based on name
                if script_name in ['tcs_schedule', 'tcs_input', 'emb_height', 'pavement_input', 'constant_fill', 'calculator', 'boq_populator']:
                    script_path = SRC_DIR / 'processor' / f'{script_name}.py'
                else:
                    script_path = SRC_DIR / 'internal' / f'{script_name}.py'
                
                # Run the script - output goes directly to console for real-time visibility
                try:
                    result = subprocess.run(
                        [sys.executable, str(script_path)],
                        capture_output=False,  # Allow real-time output to console
                        text=True,
                        timeout=step_timeout,  # Use step-specific timeout
                        env=env,
                        cwd=PROJECT_ROOT
                    )
                    
                    if result.returncode != 0:
                        error_msg = f"Script {step_name} failed with exit code {result.returncode}"
                        raise RuntimeError(error_msg)
                        
                except subprocess.TimeoutExpired:
                    error_msg = f"{step_name} timed out after {step_timeout} seconds"
                    print(f"✗ {error_msg}")
                    raise RuntimeError(error_msg)
                
                # Update progress AFTER step completion
                progress_percentage = int((step_num / total_steps) * 100)
                completion_message = f'Completed: {step_name}'
                
                # If this is the last step, show final message
                if step_num == total_steps:
                    completion_message = 'All calculations completed successfully!'
                
                session_manager.update_progress(session_id, {
                    'current_step': step_name if step_num == total_steps else processing_steps[step_num]['name'] if step_num < total_steps else step_name,
                    'current_step_number': step_num if step_num == total_steps else step_num + 1,
                    'total_steps': total_steps,
                    'percentage': progress_percentage,
                    'message': completion_message,
                    'completed_steps': step_num,
                    'last_completed_at': datetime.now(timezone.utc).isoformat()
                })
                
                print(f"✓ Completed step {step_num}/{total_steps}: {step_name}")
            
            # Final completion update
            session_manager.update_progress(session_id, {
                'current_step': 'All steps completed',
                'current_step_number': total_steps,
                'total_steps': total_steps,
                'percentage': 100,
                'message': 'All calculations completed successfully!',
                'completed_steps': total_steps,
                'completed_at': datetime.now(timezone.utc).isoformat()
            })
            
            # Calculate execution time and update session
            session = session_manager.get_session(session_id)
            if session and session['processing_info']['started_at']:
                started_at = datetime.fromisoformat(session['processing_info']['started_at'])
                execution_time = (datetime.now(timezone.utc) - started_at).total_seconds()
                
                # Upload final file to GCS (single upload after all processing)
                if session_output_file.exists():
                    from src.utils.gcs_utils import get_gcs_handler
                    gcs = get_gcs_handler()
                    output_filename = session_output_file.name
                    output_gcs_path = gcs.get_gcs_path(session_id, output_filename, 'output')
                    gcs.upload_file(str(session_output_file), output_gcs_path)
                    print(f"[GCS] Final upload completed: gs://{gcs.bucket.name}/{output_gcs_path}")
                
                # Update session with output file info
                output_info = {
                    'filename': session_output_file.name,
                    'file_path': str(session_output_file),
                    'generated_at': datetime.now(timezone.utc).isoformat(),
                    'download_count': 0
                }
                
                session_manager.set_output_file(session_id, output_info)
                session_manager.sessions.update_one(
                    {'session_id': session_id},
                    {'$set': {'processing_info.execution_time_seconds': execution_time}}
                )
                
                # Add BOQ file info to session
                session_output_dir = session_output_file.parent
                boq_output_path = session_output_dir / f"{session_id}_BOQ.xlsx"
                if boq_output_path.exists():
                    boq_info = {
                        'filename': f"{session_id}_BOQ.xlsx",
                        'file_path': str(boq_output_path),
                        'generated_at': datetime.now(timezone.utc).isoformat(),
                        'download_count': 0
                    }
                    session_manager.sessions.update_one(
                        {'session_id': session_id},
                        {'$set': {'boq_file': boq_info}}
                    )
            
            print(f"✓ Session {session_id} processing completed successfully")
            
        finally:
            os.chdir(original_cwd)
            
    except Exception as e:
        print(f"✗ Session {session_id} processing failed: {str(e)}")
        session_manager.set_error(session_id, str(e), traceback.format_exc(), "sequential_processing")

@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    try:
        # Check if template files exist
        template_exists = TEMPLATE_FILE.exists() and TEMPLATE_FILE_SINGLE.exists()
        
        # Check MongoDB connection
        mongodb_healthy = False
        mongodb_message = ""
        try:
            session_manager = SessionManager()
            session_manager.client.admin.command('ping')
            mongodb_healthy = True
            mongodb_message = "Connected"
        except Exception as e:
            mongodb_message = str(e)
        
        # Check GCS connection
        gcs_healthy = False
        gcs_message = ""
        try:
            gcs_success, gcs_msg = test_gcs_connection()
            gcs_healthy = gcs_success
            gcs_message = gcs_msg
        except Exception as e:
            gcs_message = str(e)
        
        # Determine overall health
        all_healthy = template_exists and mongodb_healthy and gcs_healthy
        
        return jsonify({
            'status': 'healthy' if all_healthy else 'degraded',
            'timestamp': datetime.now(timezone.utc).isoformat(),
            'checks': {
                'template_file': {
                    'status': 'ok' if template_exists else 'error',
                    'merged_template': {
                        'path': str(TEMPLATE_FILE),
                        'exists': TEMPLATE_FILE.exists()
                    },
                    'single_template': {
                        'path': str(TEMPLATE_FILE_SINGLE),
                        'exists': TEMPLATE_FILE_SINGLE.exists()
                    }
                },
                'mongodb': {
                    'status': 'ok' if mongodb_healthy else 'error',
                    'message': mongodb_message
                },
                'gcs': {
                    'status': 'ok' if gcs_healthy else 'error',
                    'message': gcs_message,
                    'bucket': os.getenv('GCS_BUCKET_NAME', 'not configured')
                }
            }
        }), 200
    
    except Exception as e:
        # Catch any unexpected errors during health check
        return jsonify({
            'status': 'error',
            'timestamp': datetime.now(timezone.utc).isoformat(),
            'error': 'Health check failed',
            'message': str(e),
            'traceback': traceback.format_exc()
        }), 500

@app.route('/api/test-gcs-connection', methods=['GET'])
def test_gcs_connection_api():
    """Test GCS connection and return detailed status"""
    try:
        print("\n" + "="*80)
        print("TESTING GCS CONNECTION - API ENDPOINT")
        print("="*80)
        
        # Get GCS configuration from environment
        bucket_name = os.getenv('GCS_BUCKET_NAME')
        project_id = os.getenv('GCS_PROJECT_ID')
        credentials_path = os.getenv('GCS_CREDENTIALS_PATH')
        
        # Check if environment variables are set
        if not bucket_name:
            return jsonify({
                'status': 'error',
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'message': 'GCS_BUCKET_NAME not set in environment variables',
                'configuration': {
                    'bucket_name': None,
                    'project_id': project_id,
                    'credentials_path': credentials_path
                }
            }), 400
        
        if not project_id:
            return jsonify({
                'status': 'error',
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'message': 'GCS_PROJECT_ID not set in environment variables',
                'configuration': {
                    'bucket_name': bucket_name,
                    'project_id': None,
                    'credentials_path': credentials_path
                }
            }), 400
        
        if not credentials_path:
            return jsonify({
                'status': 'error',
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'message': 'GCS_CREDENTIALS_PATH not set in environment variables',
                'configuration': {
                    'bucket_name': bucket_name,
                    'project_id': project_id,
                    'credentials_path': None
                }
            }), 400
        
        # Check if credentials file exists
        if not os.path.exists(credentials_path):
            return jsonify({
                'status': 'error',
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'message': f'Credentials file not found at: {credentials_path}',
                'configuration': {
                    'bucket_name': bucket_name,
                    'project_id': project_id,
                    'credentials_path': credentials_path,
                    'credentials_file_exists': False
                }
            }), 400
        
        print(f"✓ Bucket Name: {bucket_name}")
        print(f"✓ Project ID: {project_id}")
        print(f"✓ Credentials Path: {credentials_path}")
        print(f"✓ Credentials file exists: True")
        
        # Initialize GCS handler
        print("\n[TEST 1] Initializing GCS client...")
        gcs = get_gcs_handler()
        print("✓ GCS client initialized successfully")
        
        # Test 2: Check if bucket exists and is accessible
        print("\n[TEST 2] Checking bucket access...")
        if gcs.bucket.exists():
            print(f"✓ SUCCESS: Connected to bucket '{bucket_name}'")
        else:
            return jsonify({
                'status': 'error',
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'message': f'Bucket "{bucket_name}" does not exist or is not accessible',
                'configuration': {
                    'bucket_name': bucket_name,
                    'project_id': project_id,
                    'credentials_path': credentials_path,
                    'credentials_file_exists': True
                }
            }), 500
        
        # Test 3: Check read permissions
        print("\n[TEST 3] Checking read permissions...")
        try:
            blobs = list(gcs.bucket.list_blobs(max_results=5))
            print(f"✓ SUCCESS: Can read from bucket (found {len(blobs)} objects)")
        except Exception as e:
            return jsonify({
                'status': 'error',
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'message': f'Cannot read from bucket: {str(e)}',
                'configuration': {
                    'bucket_name': bucket_name,
                    'project_id': project_id,
                    'credentials_path': credentials_path,
                    'credentials_file_exists': True
                }
            }), 500
        
        # Test 4: Check write permissions (create and delete a test file)
        print("\n[TEST 4] Checking write permissions...")
        try:
            test_file_path = '_test_connection.txt'
            test_blob = gcs.bucket.blob(test_file_path)
            test_content = f'Connection test at {datetime.now(timezone.utc).isoformat()}'
            test_blob.upload_from_string(test_content)
            print(f"✓ SUCCESS: Can write to bucket")
            
            # Clean up test file
            test_blob.delete()
            print(f"✓ SUCCESS: Can delete from bucket")
        except Exception as e:
            return jsonify({
                'status': 'error',
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'message': f'Cannot write/delete from bucket: {str(e)}',
                'configuration': {
                    'bucket_name': bucket_name,
                    'project_id': project_id,
                    'credentials_path': credentials_path,
                    'credentials_file_exists': True
                }
            }), 500
        
        # Test 5: Check session directory structure
        print("\n[TEST 5] Checking session directory structure...")
        try:
            test_session_id = '_test_session'
            test_paths = [
                f'sessions/{test_session_id}/data/test.txt',
                f'sessions/{test_session_id}/output/test.txt'
            ]
            
            for test_path in test_paths:
                blob = gcs.bucket.blob(test_path)
                blob.upload_from_string('test')
                blob.delete()
            
            print(f"✓ SUCCESS: Can create/delete in session directories")
        except Exception as e:
            return jsonify({
                'status': 'error',
                'timestamp': datetime.now(timezone.utc).isoformat(),
                'message': f'Cannot manage session directories: {str(e)}',
                'configuration': {
                    'bucket_name': bucket_name,
                    'project_id': project_id,
                    'credentials_path': credentials_path,
                    'credentials_file_exists': True
                }
            }), 500
        
        print("\n" + "="*80)
        print("GCS CONNECTION TEST: ALL TESTS PASSED ✓")
        print("="*80)
        
        return jsonify({
            'status': 'success',
            'timestamp': datetime.now(timezone.utc).isoformat(),
            'message': 'GCS connection successful',
            'configuration': {
                'bucket_name': bucket_name,
                'project_id': project_id,
                'credentials_path': credentials_path,
                'credentials_file_exists': True,
                'bucket_accessible': True,
                'read_permission': True,
                'write_permission': True,
                'session_directories_permission': True
            },
            'tests': [
                'GCS client initialization: PASSED',
                'Bucket access: PASSED',
                'Read permissions: PASSED',
                'Write permissions: PASSED',
                'Session directory structure: PASSED'
            ]
        }), 200
        
    except Exception as e:
        error_msg = f'GCS connection test failed: {str(e)}'
        print(f'\n✗ ERROR: {error_msg}')
        print(traceback.format_exc())
        return jsonify({
            'status': 'error',
            'timestamp': datetime.now(timezone.utc).isoformat(),
            'message': error_msg,
            'traceback': traceback.format_exc()
        }), 500

@app.route('/api/upload-files', methods=['POST'])
def upload_files():
    """
    Upload 4 required Excel files to GCS and create a new session
    Files are stored in GCS session-specific directories
    """
    try:
        # Validate all required files are present
        missing_files = []
        for key, filename in REQUIRED_FILES.items():
            if key not in request.files:
                missing_files.append(f"{key} ({filename})")
        
        if missing_files:
            return jsonify({
                'error': 'Missing required files',
                'missing_files': missing_files,
                'required_files': REQUIRED_FILES
            }), 400
        
        # Validate file extensions
        invalid_files = []
        for key, expected_filename in REQUIRED_FILES.items():
            file = request.files[key]
            if file.filename == '':
                invalid_files.append(f"{key}: No file selected")
            elif not allowed_file(file.filename):
                invalid_files.append(f"{key}: Invalid file type (must be .xlsx or .xls)")
        
        if invalid_files:
            return jsonify({
                'error': 'Invalid files',
                'details': invalid_files
            }), 400
        
        # Create new session
        session_manager = SessionManager()
        
        # Get optional project details
        project_name = request.form.get('project_name')
        project_id_val = request.form.get('project_id')
        
        session_data = session_manager.create_session(
            project_name=project_name, 
            project_id=project_id_val
        )
        session_id = session_data['session_id']
        
        print(f"\n{'='*80}")
        print(f"NEW SESSION CREATED: {session_id}")
        print(f"{'='*80}\n")
        
        # Initialize GCS handler
        gcs = get_gcs_handler()
        
        # Upload files to GCS
        uploaded_files = {}
        upload_errors = []
        
        try:
            for key, expected_filename in REQUIRED_FILES.items():
                file = request.files[key]
                
                print(f"Uploading {expected_filename}...")
                
                # Create temp file
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                temp_path = temp_file.name
                file.save(temp_path)
                temp_file.close()
                
                # Upload to GCS
                gcs_path = gcs.get_gcs_path(session_id, expected_filename, 'data')
                gcs.upload_file(temp_path, gcs_path)
                
                # Clean up temp file
                os.remove(temp_path)
                
                uploaded_files[key] = {
                    'filename': expected_filename,
                    'gcs_path': gcs_path,
                    'uploaded_at': datetime.now(timezone.utc).isoformat()
                }
                
                print(f"✓ Uploaded: {expected_filename} → gs://{gcs.bucket.name}/{gcs_path}")
        
        except Exception as upload_error:
            # If upload fails, clean up session and uploaded files
            upload_errors.append(str(upload_error))
            
            # Try to delete partially uploaded files
            for key, file_info in uploaded_files.items():
                try:
                    gcs.delete_file(file_info['gcs_path'])
                except:
                    pass
            
            # Delete session
            session_manager.delete_session(session_id)
            
            return jsonify({
                'error': 'File upload to GCS failed',
                'details': upload_errors,
                'session_id': session_id,
                'message': 'Session and files have been cleaned up'
            }), 500
        
        # Update session with GCS file information
        session_manager.update_session_data(session_id, {
            'uploaded_files': uploaded_files,
            'gcs_bucket': os.getenv('GCS_BUCKET_NAME'),
            'storage_type': 'gcs'
        })
        
        print(f"\n{'='*80}")
        print(f"ALL FILES UPLOADED TO GCS SUCCESSFULLY")
        print(f"Session ID: {session_id}")
        print(f"Bucket: {os.getenv('GCS_BUCKET_NAME')}")
        print(f"{'='*80}\n")
        
        return jsonify({
            'message': 'Files uploaded successfully to GCS',
            'session_id': session_id,
            'project_name': project_name,
            'project_id': project_id_val,
            'uploaded_files': uploaded_files,
            'gcs_bucket': os.getenv('GCS_BUCKET_NAME'),
            'next_step': 'Call /api/execute-calculation with this session_id to start processing'
        }), 201
    
    except Exception as e:
        print(f"Error in upload_files: {str(e)}")
        print(traceback.format_exc())
        return jsonify({
            'error': 'Upload failed',
            'message': str(e),
            'traceback': traceback.format_exc()
        }), 500

@app.route('/api/execute-calculation', methods=['POST'])
def execute_calculation():
    """Execute calculations for main_carriageway only (starts background thread)"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'No JSON data provided'}), 400
            
        session_id = data.get('session_id')
        
        if not session_id:
            return jsonify({'error': 'session_id is required'}), 400
        
        print(f"Looking for session: {session_id}")
        
        session_manager = SessionManager()
        session = session_manager.get_session(session_id)
        
        if not session:
            print(f"Session not found in database: {session_id}")
            return jsonify({'error': 'Session not found'}), 404
        
        print(f"Found session: {session_id}, status: {session.get('status')}")
        
        if session['status'] != 'uploaded':
            return jsonify({
                'error': f"Cannot start calculation",
                'message': f"Session status is '{session['status']}', must be 'uploaded'"
            }), 400
        
        # Create session output directory
        session_output_dir = OUTPUT_DIR / 'sessions' / session_id
        session_output_dir.mkdir(parents=True, exist_ok=True)
        output_filename = f"{session_id}_main_carriageway.xlsx"
        session_output_file = session_output_dir / output_filename
        
        print(f"Creating output file: {session_output_file}")
        
        # Copy template to session output directory
        if not TEMPLATE_FILE_SINGLE.exists():
            return jsonify({'error': f'Template file not found: {TEMPLATE_FILE_SINGLE}'}), 500
            
        shutil.copy2(TEMPLATE_FILE_SINGLE, session_output_file)
        print(f"Template copied to session output directory")
        
        # Update status to processing immediately
        session_manager.update_session_status(session_id, 'processing')
        print(f"Session status updated to 'processing'")
        
        # Start processing in background thread with is_merged=False
        thread = threading.Thread(
            target=run_session_processing,
            args=(session_id, DATA_DIR / 'sessions' / session_id, session_output_file, False)
        )
        thread.daemon = True
        thread.start()
        
        print(f"Background processing started for session: {session_id} (main_carriageway only)")
        
        return jsonify({
            'status': 'success',
            'session_id': session_id,
            'message': 'Calculation started in background',
            'output_file': output_filename
        }), 200
        
    except Exception as e:
        print(f"Error in execute_calculation: {str(e)}")
        print(traceback.format_exc())
        return jsonify({
            'error': 'Execution failed',
            'message': str(e),
            'traceback': traceback.format_exc()
        }), 500

@app.route('/api/execute-calculation-merged', methods=['POST'])
def execute_calculation_merged():
    """Execute calculations for main_carriageway_and_boq (starts background thread)"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'No JSON data provided'}), 400
            
        session_id = data.get('session_id')
        
        if not session_id:
            return jsonify({'error': 'session_id is required'}), 400
        
        print(f"Looking for session: {session_id}")
        
        session_manager = SessionManager()
        session = session_manager.get_session(session_id)
        
        if not session:
            print(f"Session not found in database: {session_id}")
            return jsonify({'error': 'Session not found'}), 404
        
        print(f"Found session: {session_id}, status: {session.get('status')}")
        
        if session['status'] != 'uploaded':
            return jsonify({
                'error': f"Cannot start calculation",
                'message': f"Session status is '{session['status']}', must be 'uploaded'"
            }), 400
        
        # Create session output directory
        session_output_dir = OUTPUT_DIR / 'sessions' / session_id
        session_output_dir.mkdir(parents=True, exist_ok=True)
        output_filename = f"{session_id}_main_carriageway_and_boq.xlsx"
        session_output_file = session_output_dir / output_filename
        
        print(f"Creating output file: {session_output_file}")
        
        # Copy template to session output directory
        if not TEMPLATE_FILE.exists():
            return jsonify({'error': f'Template file not found: {TEMPLATE_FILE}'}), 500
            
        shutil.copy2(TEMPLATE_FILE, session_output_file)
        print(f"Template copied to session output directory")
        
        # Update status to processing immediately
        session_manager.update_session_status(session_id, 'processing')
        print(f"Session status updated to 'processing'")
        
        # Start processing in background thread with is_merged=True
        thread = threading.Thread(
            target=run_session_processing,
            args=(session_id, DATA_DIR / 'sessions' / session_id, session_output_file, True)
        )
        thread.daemon = True
        thread.start()
        
        print(f"Background processing started for session: {session_id} (main_carriageway_and_boq)")
        
        return jsonify({
            'status': 'success',
            'session_id': session_id,
            'message': 'Calculation started in background',
            'output_file': output_filename
        }), 200
        
    except Exception as e:
        print(f"Error in execute_calculation_merged: {str(e)}")
        print(traceback.format_exc())
        return jsonify({
            'error': 'Execution failed',
            'message': str(e),
            'traceback': traceback.format_exc()
        }), 500

@app.route('/api/execute-calculation-sync', methods=['POST'])
def execute_calculation_sync():
    """Execute calculations for main_carriageway only (synchronous - waits for completion)"""
    try:
        data = request.get_json()
        session_id = data.get('session_id')
        
        if not session_id:
            return jsonify({'error': 'session_id is required'}), 400
        
        session_manager = SessionManager()
        session = session_manager.get_session(session_id)
        
        if not session:
            return jsonify({'error': 'Session not found'}), 404
        
        if session['status'] != 'uploaded':
            return jsonify({
                'error': f"Cannot start calculation",
                'message': f"Session status is '{session['status']}', must be 'uploaded'"
            }), 400
        
        # Create session output file
        session_output_dir = OUTPUT_DIR / 'sessions' / session_id
        session_output_dir.mkdir(parents=True, exist_ok=True)
        output_filename = f"{session_id}_main_carriageway.xlsx"
        session_output_file = session_output_dir / output_filename
        
        # Copy template to session output directory
        if not TEMPLATE_FILE_SINGLE.exists():
            return jsonify({'error': f'Template file not found: {TEMPLATE_FILE_SINGLE}'}), 500
        
        shutil.copy2(TEMPLATE_FILE_SINGLE, session_output_file)
        
        # Run processing synchronously (blocks until completion) with is_merged=False
        try:
            run_session_processing(session_id, DATA_DIR / 'sessions' / session_id, session_output_file, False)
            
            # Check if processing was successful
            updated_session = session_manager.get_session(session_id)
            if updated_session['status'] == 'completed':
                return jsonify({
                    'status': 'success',
                    'session_id': session_id,
                    'message': 'Calculation completed successfully',
                    'output_file': output_filename,
                    'execution_time_seconds': updated_session['processing_info'].get('execution_time_seconds')
                }), 200
            else:
                return jsonify({
                    'error': 'Calculation failed',
                    'session_id': session_id,
                    'message': updated_session['error_info'].get('error_message', 'Unknown error'),
                    'failed_script': updated_session['error_info'].get('failed_script')
                }), 500
                
        except Exception as processing_error:
            return jsonify({
                'error': 'Processing error',
                'session_id': session_id,
                'message': str(processing_error)
            }), 500
        
    except Exception as e:
        return jsonify({
            'error': 'Execution failed',
            'message': str(e)
        }), 500

@app.route('/api/execute-calculation-sync-merged', methods=['POST'])
def execute_calculation_sync_merged():
    """Execute calculations for main_carriageway_and_boq (synchronous - waits for completion)"""
    try:
        data = request.get_json()
        session_id = data.get('session_id')
        
        if not session_id:
            return jsonify({'error': 'session_id is required'}), 400
        
        session_manager = SessionManager()
        session = session_manager.get_session(session_id)
        
        if not session:
            return jsonify({'error': 'Session not found'}), 404
        
        if session['status'] != 'uploaded':
            return jsonify({
                'error': f"Cannot start calculation",
                'message': f"Session status is '{session['status']}', must be 'uploaded'"
            }), 400
        
        # Create session output file
        session_output_dir = OUTPUT_DIR / 'sessions' / session_id
        session_output_dir.mkdir(parents=True, exist_ok=True)
        output_filename = f"{session_id}_main_carriageway_and_boq.xlsx"
        session_output_file = session_output_dir / output_filename
        
        # Copy template to session output directory
        if not TEMPLATE_FILE.exists():
            return jsonify({'error': f'Template file not found: {TEMPLATE_FILE}'}), 500
        
        shutil.copy2(TEMPLATE_FILE, session_output_file)
        
        # Run processing synchronously (blocks until completion) with is_merged=True
        try:
            run_session_processing(session_id, DATA_DIR / 'sessions' / session_id, session_output_file, True)
            
            # Check if processing was successful
            updated_session = session_manager.get_session(session_id)
            if updated_session['status'] == 'completed':
                return jsonify({
                    'status': 'success',
                    'session_id': session_id,
                    'message': 'Calculation completed successfully',
                    'output_file': output_filename,
                    'execution_time_seconds': updated_session['processing_info'].get('execution_time_seconds')
                }), 200
            else:
                return jsonify({
                    'error': 'Calculation failed',
                    'session_id': session_id,
                    'message': updated_session['error_info'].get('error_message', 'Unknown error'),
                    'failed_script': updated_session['error_info'].get('failed_script')
                }), 500
                
        except Exception as processing_error:
            return jsonify({
                'error': 'Processing error',
                'session_id': session_id,
                'message': str(processing_error)
            }), 500
        
    except Exception as e:
        return jsonify({
            'error': 'Execution failed',
            'message': str(e)
        }), 500

@app.route('/api/session/<session_id>', methods=['DELETE'])
def delete_session(session_id):
    """
    Delete a session and all its associated data (files in GCS, database record, local files)
    """
    try:
        session_manager = SessionManager()
        session = session_manager.get_session(session_id)
        
        if not session:
            return jsonify({'error': 'Session not found'}), 404
            
        print(f"\n{'='*80}")
        print(f"DELETING SESSION: {session_id}")
        print(f"{'='*80}")
        
        # 1. Delete files from GCS
        try:
            gcs = get_gcs_handler()
            # List all objects with session prefix
            session_prefix = f"sessions/{session_id}/"
            blobs = gcs.list_files(prefix=session_prefix)
            
            deleted_count = 0
            for blob_name in blobs:
                try:
                    gcs.delete_file(blob_name)
                    deleted_count += 1
                except Exception as e:
                    print(f"Warning: Failed to delete GCS file {blob_name}: {str(e)}")
            
            print(f"✓ Deleted {deleted_count} files from GCS")
            
        except Exception as e:
            print(f"Warning: Error during GCS cleanup: {str(e)}")
            # Continue with deletion even if GCS cleanup fails partially
            
        # 2. Delete local files if they exist
        try:
            local_deleted = False
            
            # Check data dir
            local_data_dir = DATA_DIR / 'sessions' / session_id
            if local_data_dir.exists():
                shutil.rmtree(local_data_dir)
                local_deleted = True
                
            # Check output dir
            local_output_dir = OUTPUT_DIR / 'sessions' / session_id
            if local_output_dir.exists():
                shutil.rmtree(local_output_dir)
                local_deleted = True
                
            if local_deleted:
                print(f"✓ Deleted local session directories")
                
        except Exception as e:
            print(f"Warning: Error during local file cleanup: {str(e)}")
            
        # 3. Delete from database
        if session_manager.delete_session(session_id):
            print(f"✓ Deleted session record from database")
            
            return jsonify({
                'status': 'success',
                'session_id': session_id,
                'message': 'Session and all associated data deleted successfully'
            }), 200
        else:
            return jsonify({
                'error': 'Database deletion failed',
                'session_id': session_id
            }), 500
            
    except Exception as e:
        print(f"Error deleting session: {str(e)}")
        print(traceback.format_exc())
        return jsonify({
            'error': 'Delete failed',
            'message': str(e),
            'traceback': traceback.format_exc()
        }), 500

@app.route('/api/session-status/<session_id>', methods=['GET'])
def get_session_status(session_id):
    """Get session status with progress"""
    try:
        session_manager = SessionManager()
        session = session_manager.get_session(session_id)
        
        if not session:
            return jsonify({'error': 'Session not found'}), 404
        
        # Generate signed URLs for uploaded files
        uploaded_files_with_urls = {}
        if session.get('uploaded_files'):
            gcs = get_gcs_handler()
            for key, file_info in session['uploaded_files'].items():
                gcs_path = file_info.get('gcs_path')
                if gcs_path:
                    signed_url = gcs.generate_signed_url(
                        gcs_path,
                        expires_in_seconds=3600,
                        response_disposition=f'attachment; filename="{file_info.get("filename", key)}"',
                        response_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                    uploaded_files_with_urls[key] = {
                        **file_info,
                        'download_url': signed_url
                    }
        
        response_data = {
            'session_id': session_id,
            'project_name': session.get('project_name'),
            'project_id': session.get('project_id'),
            'status': session['status'],
            'created_at': session.get('created_at'),
            'updated_at': session.get('updated_at'),
            'has_error': session['error_info']['has_error'],
            'error_message': session['error_info']['error_message'],
            'input_files_count': len(session['input_files']),
            'uploaded_files': uploaded_files_with_urls,
            'output_file': session['output_file'],
            'boq_file': session.get('boq_file'),
            'zip_available': True,  # Add this
            'files_in_zip': [
                f"{session_id}_main_carriageway.xlsx",
                f"{session_id}_BOQ.xlsx"
            ],
            'progress': session.get('progress', {})  # Make sure this line includes progress
        }
        
        # Add processing info if available
        if session['processing_info']:
            response_data['processing_info'] = session['processing_info']
        
        return jsonify(response_data), 200
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/download-file/<session_id>', methods=['GET'])
def download_file(session_id):
    """
    Download the generated output file from GCS
    Returns main_carriageway_and_boq.xlsx if it exists, otherwise main_carriageway.xlsx
    """
    try:
        session_manager = SessionManager()
        session = session_manager.get_session(session_id)
        
        if not session:
            return jsonify({'error': 'Session not found'}), 404
        
        if session.get('status') != 'completed':
            return jsonify({
                'error': 'File not ready',
                'message': 'Processing not completed yet',
                'current_status': session.get('status')
            }), 400
        
        # Initialize GCS
        gcs = get_gcs_handler()
        
        # Check for both files, prioritize main_carriageway_and_boq if both exist
        merged_filename = f"{session_id}_main_carriageway_and_boq.xlsx"
        single_filename = f"{session_id}_main_carriageway.xlsx"
        
        merged_gcs_path = gcs.get_gcs_path(session_id, merged_filename, 'output')
        single_gcs_path = gcs.get_gcs_path(session_id, single_filename, 'output')
        
        # Determine which file to download
        output_filename = None
        gcs_path = None
        
        if gcs.file_exists(merged_gcs_path):
            output_filename = merged_filename
            gcs_path = merged_gcs_path
        elif gcs.file_exists(single_gcs_path):
            output_filename = single_filename
            gcs_path = single_gcs_path
        else:
            return jsonify({
                'error': 'File not found in GCS',
                'message': 'Neither main_carriageway.xlsx nor main_carriageway_and_boq.xlsx found',
                'checked_paths': {
                    'merged': merged_gcs_path,
                    'single': single_gcs_path
                }
            }), 404
        
        # Generate signed URL for direct download from GCS
        # URL expires in 1 hour (3600 seconds)
        response_disposition = f'attachment; filename="{output_filename}"'
        response_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        
        signed_url = gcs.generate_signed_url(
            gcs_path,
            expires_in_seconds=3600,
            response_disposition=response_disposition,
            response_type=response_type
        )
        
        # Return signed URL as JSON response
        return jsonify({
            'status': 'success',
            'download_url': signed_url,
            'filename': output_filename,
            'expires_in_seconds': 3600,
            'message': 'Download URL generated successfully'
        }), 200
    
    except Exception as e:
        print(f"Error downloading file: {str(e)}")
        print(traceback.format_exc())
        return jsonify({
            'error': 'Download failed',
            'message': str(e)
        }), 500
    
@app.route('/api/preview-file/<session_id>', methods=['GET'])
def preview_file(session_id):
    """
    Get a preview URL for the generated output file from GCS
    Returns a signed URL that opens the file inline in the browser for preview
    """
    try:
        session_manager = SessionManager()
        session = session_manager.get_session(session_id)
        
        if not session:
            return jsonify({'error': 'Session not found'}), 404
        
        if session.get('status') != 'completed':
            return jsonify({
                'error': 'File not ready',
                'message': 'Processing not completed yet',
                'current_status': session.get('status')
            }), 400
        
        # Initialize GCS
        gcs = get_gcs_handler()
        
        # Check for both files, prioritize main_carriageway_and_boq if both exist
        merged_filename = f"{session_id}_main_carriageway_and_boq.xlsx"
        single_filename = f"{session_id}_main_carriageway.xlsx"
        
        merged_gcs_path = gcs.get_gcs_path(session_id, merged_filename, 'output')
        single_gcs_path = gcs.get_gcs_path(session_id, single_filename, 'output')
        
        # Determine which file to preview
        output_filename = None
        gcs_path = None
        
        if gcs.file_exists(merged_gcs_path):
            output_filename = merged_filename
            gcs_path = merged_gcs_path
        elif gcs.file_exists(single_gcs_path):
            output_filename = single_filename
            gcs_path = single_gcs_path
        else:
            return jsonify({
                'error': 'File not found in GCS',
                'message': 'Neither main_carriageway.xlsx nor main_carriageway_and_boq.xlsx found',
                'checked_paths': {
                    'merged': merged_gcs_path,
                    'single': single_gcs_path
                }
            }), 404
        
        # Generate signed URL for inline preview from GCS
        # URL expires in 1 hour (3600 seconds)
        response_disposition = f'inline; filename="{output_filename}"'
        response_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        
        signed_url = gcs.generate_signed_url(
            gcs_path,
            expires_in_seconds=3600,
            response_disposition=response_disposition,
            response_type=response_type
        )
        
        # Return signed URL as JSON response
        return jsonify({
            'status': 'success',
            'preview_url': signed_url,
            'filename': output_filename,
            'expires_in_seconds': 3600,
            'message': 'Preview URL generated successfully'
        }), 200
    
    except Exception as e:
        print(f"Error generating preview URL: {str(e)}")
        print(traceback.format_exc())
        return jsonify({
            'error': 'Preview URL generation failed',
            'message': str(e)
        }), 500


@app.route('/api/download-boq/<session_id>', methods=['GET'])
def download_boq(session_id):
    """Download BOQ file for session"""
    try:
        session_manager = SessionManager()
        session = session_manager.get_session(session_id)
        
        if not session:
            return jsonify({'error': 'Session not found'}), 404
        
        if not session.get('boq_file'):
            return jsonify({
                'error': 'BOQ file not generated',
                'message': 'BOQ file has not been generated for this session'
            }), 404
        
        boq_file_path = session['boq_file']['file_path']
        
        if not os.path.exists(boq_file_path):
            return jsonify({'error': 'BOQ file not found'}), 404
        
        # Increment download count
        session_manager.sessions.update_one(
            {'session_id': session_id},
            {'$inc': {'boq_file.download_count': 1}}
        )
        
        return send_file(
            boq_file_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=session['boq_file']['filename']
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@app.route('/api/download-session/<session_id>', methods=['GET'])
def download_session_zip(session_id):
    """Download main carriageway file as ZIP from GCS"""
    try:
        session_manager = SessionManager()
        session = session_manager.get_session(session_id)
        
        if not session:
            return jsonify({'error': 'Session not found'}), 404
        
        if session.get('status') != 'completed':
            return jsonify({
                'error': 'Files not ready',
                'message': 'Processing not completed yet',
                'current_status': session.get('status')
            }), 400
        
        # Initialize GCS
        gcs = get_gcs_handler()
        
        # Create temporary directory
        temp_dir = tempfile.mkdtemp()
        zip_path = os.path.join(temp_dir, f'{session_id}_files.zip')
        
        try:
            # Download file from GCS
            main_filename = f"{session_id}_main_carriageway_and_boq.xlsx"
            main_gcs_path = gcs.get_gcs_path(session_id, main_filename, 'output')
            
            if not gcs.file_exists(main_gcs_path):
                return jsonify({
                    'error': 'File not found in GCS',
                    'gcs_path': main_gcs_path
                }), 404
            
            temp_main_file = gcs.download_to_temp(main_gcs_path, suffix='.xlsx')
            
            # Create ZIP with just the Excel file
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                zipf.write(temp_main_file, main_filename)
            
            # Clean up
            os.remove(temp_main_file)
            
            # Track download
            session_manager.increment_download_count(session_id)
            
            # Send ZIP
            response = send_file(
                zip_path,
                as_attachment=True,
                download_name=f'{session_id}_BOQ_files.zip',
                mimetype='application/zip'
            )
            
            # Cleanup after sending
            @response.call_on_close
            def cleanup():
                try:
                    if os.path.exists(zip_path):
                        os.remove(zip_path)
                    if os.path.exists(temp_dir):
                        os.rmdir(temp_dir)
                except:
                    pass
            
            return response
        
        except Exception as inner_error:
            try:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
            except:
                pass
            raise inner_error
    
    except Exception as e:
        print(f"Error creating ZIP: {str(e)}")
        return jsonify({
            'error': 'ZIP creation failed',
            'message': str(e)
        }), 500

@app.route('/api/output-file-paths/<session_id>', methods=['GET'])
def get_output_file_paths(session_id):
    """Get output file paths for a session"""
    try:
        session_manager = SessionManager()
        session = session_manager.get_session(session_id)
        
        if not session:
            return jsonify({'error': 'Session not found'}), 404
        
        if session.get('status') != 'completed':
            return jsonify({
                'error': 'Files not ready',
                'message': 'Processing not completed yet',
                'current_status': session.get('status')
            }), 400
        
        # Initialize GCS
        gcs = get_gcs_handler()
        
        # Get file paths
        file_paths = {}
        
        # Check for merged file (main_carriageway_and_boq)
        merged_filename = f"{session_id}_main_carriageway_and_boq.xlsx"
        merged_gcs_path = gcs.get_gcs_path(session_id, merged_filename, 'output')
        
        if gcs.file_exists(merged_gcs_path):
            file_paths['main_carriageway_and_boq'] = {
                'filename': merged_filename,
                'gcs_path': merged_gcs_path,
                'gcs_uri': f"gs://{gcs.bucket.name}/{merged_gcs_path}",
                'download_url': f"/api/download-file/{session_id}",
                'preview_url': f"/api/preview-file/{session_id}",
                'file_type': 'main_output_merged'
            }
        
        # Check for single file (main_carriageway)
        single_filename = f"{session_id}_main_carriageway.xlsx"
        single_gcs_path = gcs.get_gcs_path(session_id, single_filename, 'output')
        
        if gcs.file_exists(single_gcs_path):
            file_paths['main_carriageway'] = {
                'filename': single_filename,
                'gcs_path': single_gcs_path,
                'gcs_uri': f"gs://{gcs.bucket.name}/{single_gcs_path}",
                'download_url': f"/api/download-file/{session_id}",
                'preview_url': f"/api/preview-file/{session_id}",
                'file_type': 'main_output_single'
            }
        
        # BOQ file (if exists)
        boq_filename = f"{session_id}_BOQ.xlsx"
        boq_gcs_path = gcs.get_gcs_path(session_id, boq_filename, 'output')
        
        if gcs.file_exists(boq_gcs_path):
            file_paths['boq'] = {
                'filename': boq_filename,
                'gcs_path': boq_gcs_path,
                'gcs_uri': f"gs://{gcs.bucket.name}/{boq_gcs_path}",
                'download_url': f"/api/download-boq/{session_id}",
                'file_type': 'boq'
            }
        
        # Session ZIP file
        file_paths['session_zip'] = {
            'filename': f"{session_id}_BOQ_files.zip",
            'download_url': f"/api/download-session/{session_id}",
            'file_type': 'zip_archive',
            'description': 'Contains all generated files for the session'
        }
        
        # Add session metadata
        response_data = {
            'session_id': session_id,
            'status': session.get('status'),
            'created_at': session.get('created_at').isoformat() if session.get('created_at') else None,
            'completed_at': session.get('processing_info', {}).get('completed_at').isoformat() if session.get('processing_info', {}).get('completed_at') else None,
            'execution_time_seconds': session.get('processing_info', {}).get('execution_time_seconds'),
            'gcs_bucket': gcs.bucket.name,
            'files': file_paths,
            'available_downloads': len([f for f in file_paths.values() if f.get('download_url')]),
            'file_count': len([f for f in file_paths.values() if f.get('gcs_path')])  # Only count actual files in GCS
        }
        
        # Add output file info from session if available
        if session.get('output_file'):
            response_data['output_file_info'] = session['output_file']
        
        # Add BOQ file info from session if available
        if session.get('boq_file'):
            response_data['boq_file_info'] = session['boq_file']
        
        return jsonify(response_data), 200
        
    except Exception as e:
        print(f"Error getting output file paths: {str(e)}")
        print(traceback.format_exc())
        return jsonify({
            'error': 'Failed to get output file paths',
            'message': str(e),
            'traceback': traceback.format_exc()
        }), 500

@app.route('/api/sessions', methods=['GET'])
def get_all_sessions():
    """Get all sessions with pagination and filtering"""
    try:
        session_manager = SessionManager()
        
        # Get query parameters with defaults
        page = request.args.get('page', 1, type=int)
        limit = request.args.get('limit', 10, type=int)
        status = request.args.get('status', None)
        search = request.args.get('search', None)
        project_id = request.args.get('project_id', None)
        
        # Validate pagination
        if page < 1:
            page = 1
        if limit < 1 or limit > 100:
            limit = 10
        
        # Calculate skip for pagination
        skip = (page - 1) * limit
        
        # Get sessions using SessionManager with filters
        sessions, total_sessions = session_manager.get_all_sessions(
            limit=limit, 
            skip=skip, 
            status_filter=status, 
            search_filter=search,
            project_id_filter=project_id
        )
        total_pages = (total_sessions + limit - 1) // limit  # Ceiling division

        return jsonify({
            'sessions': sessions,
            'total_sessions': total_sessions,
            'total_pages': total_pages,
            'current_page': page,
            'limit': limit,
            'filters': {
                'status': status,
                'search': search,
                'project_id': project_id
            }
        }), 200
        
    except Exception as e:
        print(f"Error in get_all_sessions: {str(e)}")
        print(traceback.format_exc())
        return jsonify({
            'error': 'Failed to fetch sessions',
            'message': str(e),
            'traceback': traceback.format_exc()
        }), 500

@app.route('/api/analytics', methods=['GET'])
def get_analytics():
    """Get analytics for BOQ reports - total counts by status"""
    try:
        session_manager = SessionManager()
        
        # Use MongoDB aggregation for efficient counting
        pipeline = [
            {
                '$group': {
                    '_id': '$status',
                    'count': {'$sum': 1}
                }
            }
        ]
        
        # Execute aggregation
        aggregation_result = list(session_manager.sessions.aggregate(pipeline))
        
        # Initialize counts
        status_counts = {
            'total_reports_generated': 0,  # completed
            'total_reports_processing': 0,  # processing
            'total_reports_uploaded': 0,    # uploaded
            'total_reports_error': 0,       # error
            'total_sessions': 0
        }
        
        # Process aggregation results
        for result in aggregation_result:
            status = result['_id']
            count = result['count']
            status_counts['total_sessions'] += count
            
            if status == 'completed':
                status_counts['total_reports_generated'] = count
            elif status == 'processing':
                status_counts['total_reports_processing'] = count
            elif status == 'uploaded':
                status_counts['total_reports_uploaded'] = count
            elif status == 'error':
                status_counts['total_reports_error'] = count
        
        return jsonify({
            'analytics': status_counts,
            'timestamp': datetime.now(timezone.utc).isoformat(),
        }), 200
        
    except Exception as e:
        print(f"Error in get_analytics: {str(e)}")
        print(traceback.format_exc())
        return jsonify({
            'error': 'Failed to fetch analytics',
            'message': str(e),
            'traceback': traceback.format_exc()
        }), 500

@app.route('/api/stats', methods=['GET'])
def get_stats():
    """
    Get processing statistics - total sessions, completed, processing, and failed
    """
    try:
        session_manager = SessionManager()
        
        # Use MongoDB aggregation for efficient counting
        pipeline = [
            {
                '$group': {
                    '_id': '$status',
                    'count': {'$sum': 1}
                }
            }
        ]
        
        # Execute aggregation
        aggregation_result = list(session_manager.sessions.aggregate(pipeline))
        
        # Initialize counts
        stats = {
            'total_sessions': 0,
            'completed': 0,
            'processing': 0,
            'failed': 0
        }
        
        # Process aggregation results
        for result in aggregation_result:
            status = result['_id']
            count = result['count']
            stats['total_sessions'] += count
            
            if status == 'completed':
                stats['completed'] = count
            elif status == 'processing':
                stats['processing'] = count
            elif status == 'failed':
                stats['failed'] = count
        
        # Add timestamp
        stats['timestamp'] = datetime.now(timezone.utc).isoformat()
        
        # Calculate pending (uploaded sessions that haven't started processing yet)
        stats['pending'] = stats['total_sessions'] - (stats['processing'] + stats['completed'] + stats['failed'])
        
        return jsonify(stats), 200
        
    except Exception as e:
        print(f"Error in get_stats: {str(e)}")
        print(traceback.format_exc())
        return jsonify({
            'error': 'Failed to fetch stats',
            'message': str(e)
        }), 500

@app.route('/', methods=['GET'])
@app.route('/api', methods=['GET'])
def root():
    """API Root - Get available endpoints and basic usage information"""
    return jsonify({
        'title': 'BOQ Generation API',
        'version': '1.0.0',
        'description': 'Flask API for BOQ (Bill of Quantities) Generation with Session Management',
        'endpoints': {
            'health': {
                'method': 'GET',
                'path': '/health',
                'description': 'Health check endpoint',
                'usage': 'curl -X GET http://localhost:5000/health'
            },
            'upload_files': {
                'method': 'POST',
                'path': '/api/upload-files',
                'description': 'Upload 4 Excel files and create a new session',
                'required_files': [
                    'tcs_schedule (TCS Schedule.xlsx)',
                    'tcs_input (TCS Input.xlsx)',
                    'emb_height (Emb Height.xlsx)',
                    'pavement_input (Pavement Input.xlsx)'
                ],
                'usage': 'curl -X POST -F "tcs_schedule=@TCS_Schedule.xlsx" -F "tcs_input=@TCS_Input.xlsx" -F "emb_height=@Emb_Height.xlsx" -F "pavement_input=@Pavement_Input.xlsx" http://localhost:5000/api/upload-files'
            },
            'execute_calculation': {
                'method': 'POST',
                'path': '/api/execute-calculation',
                'description': 'Start calculation for main_carriageway only (runs in background)',
                'request_body': {
                    'session_id': 'string (required)'
                },
                'usage': 'curl -X POST -H "Content-Type: application/json" -d \'{"session_id": "your_session_id"}\' http://localhost:5000/api/execute-calculation'
            },
            'execute_calculation_merged': {
                'method': 'POST',
                'path': '/api/execute-calculation-merged',
                'description': 'Start calculation for main_carriageway_and_boq (runs in background)',
                'request_body': {
                    'session_id': 'string (required)'
                },
                'usage': 'curl -X POST -H "Content-Type: application/json" -d \'{"session_id": "your_session_id"}\' http://localhost:5000/api/execute-calculation-merged'
            },
            'execute_calculation_sync': {
                'method': 'POST',
                'path': '/api/execute-calculation-sync',
                'description': 'Start calculation for main_carriageway only (waits for completion)',
                'request_body': {
                    'session_id': 'string (required)'
                },
                'usage': 'curl -X POST -H "Content-Type: application/json" -d \'{"session_id": "your_session_id"}\' http://localhost:5000/api/execute-calculation-sync'
            },
            'execute_calculation_sync_merged': {
                'method': 'POST',
                'path': '/api/execute-calculation-sync-merged',
                'description': 'Start calculation for main_carriageway_and_boq (waits for completion)',
                'request_body': {
                    'session_id': 'string (required)'
                },
                'usage': 'curl -X POST -H "Content-Type: application/json" -d \'{"session_id": "your_session_id"}\' http://localhost:5000/api/execute-calculation-sync-merged'
            },
            'test_gcs_connection': {
                'method': 'GET',
                'path': '/api/test-gcs-connection',
                'description': 'Test GCS connection and permissions',
                'usage': 'curl -X GET http://localhost:5000/api/test-gcs-connection'
            },
            'session_list': {
                'method': 'GET',
                'path': '/api/sessions',
                'description': 'Get list of all sessions (paginated) with GCS download URLs for uploaded files',
                'query_parameters': {
                    'page': 'integer (optional, default: 1)',
                    'limit': 'integer (optional, default: 10, max: 100)'
                },
                'usage': 'curl -X GET "http://localhost:5000/api/sessions?page=1&limit=10"'
            },
            'session_status': {
                'method': 'GET',
                'path': '/api/session-status/<session_id>',
                'description': 'Check the status of a processing session',
                'usage': 'curl -X GET http://localhost:5000/api/session-status/your_session_id'
            },
            'delete_session': {
                'method': 'DELETE',
                'path': '/api/session/<session_id>',
                'description': 'Delete a session and all associated files',
                'response': {
                    'status': 'success',
                    'session_id': 'session_id',
                    'message': 'Session and all associated data deleted successfully'
                },
                'usage': 'curl -X DELETE http://localhost:5000/api/session/your_session_id'
            },
            'download_file': {
                'method': 'GET',
                'path': '/api/download-file/<session_id>',
                'description': 'Get signed URL for downloading the generated file (only available after completion). Returns JSON with download_url that expires in 1 hour.',
                'response': {
                    'status': 'success',
                    'download_url': 'https://storage.googleapis.com/... (signed URL)',
                    'filename': 'session_id_main_carriageway_and_boq.xlsx',
                    'expires_in_seconds': 3600,
                    'message': 'Download URL generated successfully'
                },
                'usage': 'curl -X GET http://localhost:5000/api/download-file/your_session_id'
            },
            'preview_file': {
                'method': 'GET',
                'path': '/api/preview-file/<session_id>',
                'description': 'Get signed URL for previewing the generated file inline in browser (only available after completion). Returns JSON with preview_url that expires in 1 hour.',
                'response': {
                    'status': 'success',
                    'preview_url': 'https://storage.googleapis.com/... (signed URL)',
                    'filename': 'session_id_main_carriageway_and_boq.xlsx',
                    'expires_in_seconds': 3600,
                    'message': 'Preview URL generated successfully'
                },
                'usage': 'curl -X GET http://localhost:5000/api/preview-file/your_session_id'
            },
            'download_boq_file': {
                'method': 'GET',
                'path': '/api/download-boq/<session_id>',
                'description': 'Download the generated BOQ file',
                'usage': 'curl -X GET -O http://localhost:5000/api/download-boq/your_session_id'
            },
            'download_session_zip': {
                'method': 'GET',
                'path': '/api/download-session/<session_id>',
                'description': 'Download both main carriageway and BOQ files as ZIP',
                'usage': 'curl -X GET -O http://localhost:5000/api/download-session/your_session_id'
            },
            'output_file_paths': {
                'method': 'GET',
                'path': '/api/output-file-paths/<session_id>',
                'description': 'Get output file paths and download URLs for a session',
                'usage': 'curl -X GET http://localhost:5000/api/output-file-paths/your_session_id'
            },
            'analytics': {
                'method': 'GET',
                'path': '/api/analytics',
                'description': 'Get analytics for BOQ reports - total counts by status',
                'response': {
                    'analytics': {
                        'total_reports_generated': 5,
                        'total_reports_processing': 2,
                        'total_reports_uploaded': 1,
                        'total_reports_error': 0,
                        'total_sessions': 8
                    },
                    'timestamp': '2025-12-18T10:30:00.000000',
                    'description': 'BOQ report analytics - session counts by status'
                },
                'usage': 'curl -X GET http://localhost:5000/api/analytics'
            },
            'stats': {
                'method': 'GET',
                'path': '/api/stats',
                'description': 'Get processing statistics - total sessions, completed, processing, and failed',
                'response': {
                    'total_sessions': 8,
                    'completed': 5,
                    'processing': 2,
                    'failed': 1,
                    'pending': 0,
                    'timestamp': '2025-12-27T12:05:00.000000+00:00'
                },
                'usage': 'curl -X GET http://localhost:5000/api/stats'
            }
        },
        'workflow': {
            'steps': [
                '1. Upload 4 required Excel files using /api/upload-files',
                '2. Get session_id from upload response',
                '3. Start calculation using /api/execute-calculation (main_carriageway only) or /api/execute-calculation-merged (main_carriageway_and_boq) with session_id',
                '4. Monitor progress using /api/session-status/<session_id>',
                '5. Preview result using /api/preview-file/<session_id> or download using /api/download-file/<session_id> when status is "completed"',
                '6. Download BOQ file using /api/download-boq/<session_id> (only for merged calculation)',
                '7. Download session ZIP file using /api/download-session/<session_id>'
            ]
        },
        'session_states': {
            'uploaded': 'Files uploaded, ready for calculation',
            'processing': 'Calculation in progress',
            'completed': 'Calculation completed, file ready for download',
            'error': 'Processing failed, check error message'
        },
        'notes': [
            'Maximum file size: 100 MB per file',
            'Allowed file types: .xlsx, .xls',
            'Session timeout: 10 minutes for processing',
            'All files are stored temporarily and cleaned up automatically',
            'For best results, download the ZIP file containing both files',
            'Extract both files to the same folder before opening the BOQ file',
            'Excel may show security prompts about external links - click "Enable"',
            'Keep both files in the same directory for formulas to work properly'
        ]
    }), 200

if __name__ == '__main__':
    # Load environment variables
    
    
    # Print MongoDB configuration
    mongodb_uri = os.getenv('MONGO_URI')
    db_name = os.getenv('MONGO_DB_NAME')
    print(f"✓ MongoDB URI: {mongodb_uri}")
    print(f"✓ Database: {db_name}")
    
    # Test MongoDB connection
    try:
        session_manager = SessionManager()
        # Ping the MongoDB server
        session_manager.client.admin.command('ping')
        print("✓ MongoDB connection test: SUCCESS")
    except Exception as e:
        print(f"⚠ MongoDB connection test: FAILED - {str(e)}")
    
    # Test GCS connection
    print("\n" + "-"*80)
    gcs_success, gcs_message = test_gcs_connection()
    if gcs_success:
        print(f"✓ GCS connection test: SUCCESS")
    else:
        print(f"⚠ GCS connection test: FAILED - {gcs_message}")
        print("WARNING: File operations will fail without GCS access!")
    print("-"*80 + "\n")
    
    # Ensure directories exist on startup
    ensure_directories()
    
    # Validate templates exist
    try:
        validate_template()
        print(f"✓ Merged template file found: {TEMPLATE_FILE}")
        print(f"✓ Single template file found: {TEMPLATE_FILE_SINGLE}")
    except FileNotFoundError as e:
        print(f"⚠ WARNING: {e}")
    
    # Print startup information
    print("\n" + "="*80)
    print("BOQ Generation Flask API with Session Management")
    print("="*80)
    print(f"Project root: {PROJECT_ROOT}")
    print(f"Data directory: {DATA_DIR}")
    print(f"Output directory: {OUTPUT_DIR}")
    print(f"Template directory: {TEMPLATE_DIR}")
    print("="*80)
    print("\nStarting Flask server...")
    print("API Endpoints:")
    print("  GET  /                                    - API root")
    print("  GET  /health                              - Health check")
    print("  POST /api/upload-files                    - Upload files & create session")
    print("  POST /api/execute-calculation             - Start calculation (main_carriageway only)")
    print("  POST /api/execute-calculation-merged       - Start calculation (main_carriageway_and_boq)")
    print("  POST /api/execute-calculation-sync         - Start calculation (main_carriageway only, synchronous)")
    print("  POST /api/execute-calculation-sync-merged - Start calculation (main_carriageway_and_boq, synchronous)")
    print("  GET  /api/test-gcs-connection              - Test GCS connection and permissions")
    print("  GET  /api/session-status/<id>             - Check session status")
    print("  DELETE /api/session/<id>                  - Delete session")
    print("  GET  /api/preview-file/<id>               - Get preview URL for result")
    print("  GET  /api/download-file/<id>              - Download result")
    print("  GET  /api/download-boq/<id>               - Download BOQ file")
    print("  GET  /api/download-session/<id>           - Download session ZIP file")
    print("  GET  /api/output-file-paths/<id>          - Get output file paths")
    print("  GET  /api/analytics                       - Get BOQ report analytics")
    print("  GET  /api/stats                           - Get processing statistics")
    print("  GET  /api/sessions                        - Get all sessions (with pagination & filtering)")
    print("    Query params: page, limit, status, search")
    print("="*80 + "\n")
    
    # Run Flask app
    app.run(debug=True, host='0.0.0.0', port=PORT)
