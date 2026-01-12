import pandas as pd
import os
import sys
import io
import shutil
from dotenv import load_dotenv
import tempfile

# Add project root to Python path
project_root = os.path.join(os.path.dirname(__file__), '..', '..')
sys.path.append(project_root)

from src.utils.gcs_utils import get_gcs_handler
from src.utils.excel_writer_utils import get_workbook_manager

load_dotenv()

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# NEW CODE:
script_dir = os.path.dirname(os.path.abspath(__file__))
session_id = os.getenv('SESSION_ID', 'default')
is_merged = os.getenv('IS_MERGED', 'True').lower() == 'true'

# Use SESSION_OUTPUT_FILE if available (local file), otherwise fallback to GCS
output_file = os.getenv('SESSION_OUTPUT_FILE', '')
if not output_file or not os.path.exists(output_file):
    # Fallback: download from GCS (for backward compatibility)
    gcs = get_gcs_handler()
    if is_merged:
        output_filename = f"{session_id}_main_carriageway_and_boq.xlsx"
    else:
        output_filename = f"{session_id}_main_carriageway.xlsx"
    output_gcs_path = gcs.get_gcs_path(session_id, output_filename, 'output')
    output_file = gcs.download_to_temp(output_gcs_path, suffix='.xlsx')
    print(f"[GCS] Downloaded output file from GCS: {output_file}")
else:
    print(f"Using local output file: {output_file}")

# Initialize GCS for input files only
gcs = get_gcs_handler()

# Download input from GCS
input_gcs_path = gcs.get_gcs_path(session_id, 'TCS Schedule.xlsx', 'data')
input_file = gcs.download_to_temp(input_gcs_path, suffix='.xlsx')

# Read columns B to E from the Excel file starting from 3rd row (row index 2)
# header=None means don't treat any row as header, just read raw data
df = pd.read_excel(input_file, sheet_name='TCS', skiprows=2, usecols='B:E', header=None)

print(f"DataFrame shape: {df.shape}")
print(f"DataFrame columns: {df.columns.tolist()}")
print(f"\nRaw data preview:")
print(df.head())

# Check if dataframe is empty
if df.empty:
    print("\nWarning: No data found after row 3 in columns B:E")
    sys.exit(1)

# Reset column names to ensure they are 0, 1, 2, 3
df.columns = range(len(df.columns))

# Convert numeric columns (first 3 columns are From, To, Length)
df[0] = pd.to_numeric(df[0], errors='coerce')  # From (column B)
df[1] = pd.to_numeric(df[1], errors='coerce')  # To (column C)
df[2] = pd.to_numeric(df[2], errors='coerce')  # Length (column D)
# df[3] is C/S Type (column E) - keep as string

# Remove any rows where all values are NaN
df_output = df.dropna(how='all')

# Round columns A and B (df[0] and df[1]) to 6 decimal places
df_output[0] = df_output[0].round(6)
df_output[1] = df_output[1].round(6)

print(f"\nData after cleaning:")
print(df_output.head())
print(f"Total rows to write: {len(df_output)}")

# Use optimized Excel writer (keep-alive pattern)
print("\n[OPTIMIZATION] Using direct openpyxl writing...")
manager = get_workbook_manager(output_file)
if not manager.is_open:
    manager.open()

manager.write_dataframe(df_output, 'Quantity', start_row=7, start_col=0, include_header=False)

print(f"\nSuccessfully wrote data to {output_file}")
print(f"Sheet: Quantity")
print(f"Starting from row: 7, column: A")
print(f"Total data rows written: {len(df_output)}")

# CRITICAL: Save and close because we run in subprocess
manager.close()
print("[OPTIMIZATION] Saved and closed (subprocess mode)")

# Note: File will be saved at the end of all processing in sequential.py
# Cleanup temp input file only
os.remove(input_file)