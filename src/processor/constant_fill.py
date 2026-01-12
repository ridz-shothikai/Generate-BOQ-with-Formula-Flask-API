"""
Constant Fill Processor
Fills specific columns in main_carriageway_and_boq.xlsx with constant values
"""

import pandas as pd
import os
import sys
import io
from dotenv import load_dotenv

# Add project root to Python path
project_root = os.path.join(os.path.dirname(__file__), '..', '..')
sys.path.append(project_root)

from src.utils.gcs_utils import get_gcs_handler
from src.utils.excel_writer_utils import get_workbook_manager

load_dotenv()
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


# ============================================================================
# FILE PATHS - Update these to match your folder structure
# ============================================================================

# NEW CODE:
script_dir = os.path.dirname(os.path.abspath(__file__))
session_id = os.getenv('SESSION_ID', 'default')
is_merged = os.getenv('IS_MERGED', 'True').lower() == 'true'

# Use SESSION_OUTPUT_FILE if available (local file), otherwise fallback to GCS
output_file = os.getenv('SESSION_OUTPUT_FILE', '')
if output_file and os.path.exists(output_file):
    MAIN_CARRIAGEWAY_FILE = output_file
    OUTPUT_EXCEL = output_file
    print(f"Using local output file: {output_file}")
else:
    # Fallback: download from GCS (for backward compatibility)
    gcs = get_gcs_handler()
    if is_merged:
        output_filename = f"{session_id}_main_carriageway_and_boq.xlsx"
    else:
        output_filename = f"{session_id}_main_carriageway.xlsx"
    output_gcs_path = gcs.get_gcs_path(session_id, output_filename, 'output')
    temp_file = gcs.download_to_temp(output_gcs_path, suffix='.xlsx')
    MAIN_CARRIAGEWAY_FILE = temp_file
    OUTPUT_EXCEL = temp_file
    print(f"[GCS] Downloaded output file from GCS: {temp_file}")


# ============================================================================
# CONSTANTS CONFIGURATION
# ============================================================================

CONSTANTS = [
    {
        'col_index': 48,  # Column AW
        'col_name': 'SUBGRADE',
        'value': 0.5
    },
    {
        'col_index': 62,  # Column BK
        'col_name': 'LHS_Subgrade_Thickness',
        'value': 0.5
    }
]


# ============================================================================
# MAIN PROCESSING
# ============================================================================

def fill_constant_columns(main_carriageway_file, constants, output_file):
    """
    Reads main_carriageway_and_boq.xlsx and fills specified columns with constant values
    """
    print("="*80)
    print("CONSTANT FILL PROCESSOR")
    print("="*80)
    
    # Read the main carriageway file
    df = pd.read_excel(main_carriageway_file, sheet_name='Quantity', skiprows=6, header=None)
    print("[OK] Read main_carriageway_and_boq.xlsx:", len(df), "rows")
    print("  Current columns:", len(df.columns))
    
    # Remove empty rows
    df = df.dropna(how='all')
    
    # Process each constant column
    for const in constants:
        col_idx = const['col_index']
        col_name = const['col_name']
        value = const['value']
        
        print(f"\n[OK] Processing Column {chr(65 + col_idx//26 - 1) if col_idx >= 26 else ''}{chr(65 + col_idx%26)} (index {col_idx})")
        
        # Ensure column exists
        if len(df.columns) <= col_idx:
            # Add empty columns up to target
            while len(df.columns) < col_idx:
                df[f'Empty_{len(df.columns)}'] = None
            # Add target column
            df.insert(col_idx, f'Col_{col_idx}', value)
            print(f"  [OK] Created new column with value {value}")
        else:
            # Column exists, update it
            df.iloc[:, col_idx] = value
            print(f"  [OK] Filled column with value {value}")
        
        print(f"  [OK] All {len(df)} rows set to {value}")
    
    # Save to Excel using optimized writer
    print(f"\n[OPTIMIZATION] Using direct openpyxl writing...")
    print(f"[OK] Saving to {output_file}...")
    
    manager = get_workbook_manager(output_file)
    if not manager.is_open:
        manager.open()
    
    manager.write_dataframe(df, 'Quantity', start_row=7, start_col=0, include_header=False)
    
    print("[OK] Saved!")
    
    # CRITICAL: Save and close because we run in subprocess
    manager.close()
    print("[OPTIMIZATION] Saved and closed (subprocess mode)")
    
    print(f"  Total columns: {len(df.columns)}")
    
    return df


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Main function to execute all steps"""
    print("\n" + "="*80)
    print("CONSTANT FILL PROCESSOR")
    print("="*80)
    print("Configuration:")
    for const in CONSTANTS:
        col_letter = chr(65 + const['col_index']//26 - 1) if const['col_index'] >= 26 else ''
        col_letter += chr(65 + const['col_index']%26)
        print(f"  • Column {col_letter} ({const['col_name']}): {const['value']}")
    print("="*80 + "\n")
    
    try:
        # Fill constant columns
        df = fill_constant_columns(MAIN_CARRIAGEWAY_FILE, CONSTANTS, OUTPUT_EXCEL)
        
        # Success summary
        print("\n" + "="*80)
        print("SUCCESS! [OK]")
        print("="*80)
        print("Output file:", OUTPUT_EXCEL)
        print("Total rows:", len(df))
        print("Total columns:", len(df.columns))
        
        # Show sample values
        print("\n" + "="*80)
        print("SAMPLE OUTPUT:")
        print("-"*80)
        for const in CONSTANTS:
            col_idx = const['col_index']
            col_name = const['col_name']
            if len(df.columns) > col_idx:
                sample_vals = df.iloc[:3, col_idx].tolist()
                print(f"\nColumn {col_name} (index {col_idx}):")
                for i, val in enumerate(sample_vals):
                    print(f"  Row {i + 2}: {val}")
        
        # Note: File will be uploaded to GCS at the end of all processing in main.py
        # No need to upload here for efficiency
        
        # Cleanup - only remove temp files, not the local output file
        # Only remove OUTPUT_EXCEL if it's a temp file (not the local SESSION_OUTPUT_FILE)
        if OUTPUT_EXCEL != os.getenv('SESSION_OUTPUT_FILE', ''):
            try:
                os.remove(OUTPUT_EXCEL)
            except:
                pass
        
    except FileNotFoundError as e:
        print("\n[ERROR] File not found")
        print(" ", e)
        print("\nPlease check:")
        print("  1. File exists in the data folder")
        print("  2. File name matches exactly")
    except Exception as e:
        print("\n[ERROR]:", e)
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
