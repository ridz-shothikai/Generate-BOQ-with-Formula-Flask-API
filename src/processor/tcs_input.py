"""
TCS Specification Populator - Extended Version
================================================
Reads columns D-V and AA-AS from TCS_Input.xlsx
PLUS column W (LEFT PAVED SHOULDER) mapped to AS, AT, AU, AV in output

Author: Auto-generated
Date: 2025
"""

import pandas as pd
import json
import os
import sys
import io
from dotenv import load_dotenv

# Add project root to Python path
project_root = os.path.join(os.path.dirname(__file__), '..', '..')
sys.path.append(project_root)

from src.utils.gcs_utils import get_gcs_handler
from src.utils.data_collector import get_collector

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
    temp_main_file = gcs.download_to_temp(output_gcs_path, suffix='.xlsx')
    MAIN_CARRIAGEWAY_FILE = temp_main_file
    OUTPUT_EXCEL = temp_main_file
    print(f"[GCS] Downloaded output file from GCS: {temp_main_file}")

# Initialize GCS for input files only
gcs = get_gcs_handler()

# Download input files from GCS
input_gcs_path = gcs.get_gcs_path(session_id, 'TCS Input.xlsx', 'data')
TCS_INPUT_FILE = gcs.download_to_temp(input_gcs_path, suffix='.xlsx')



# ============================================================================
# COLUMN RANGES TO EXTRACT
# ============================================================================

# Excel columns D to V (0-indexed: 3 to 21)
RANGE_1_START = 3  # Column D
RANGE_1_END = 21   # Column V (inclusive)

# Excel columns AA to AS (0-indexed: 26 to 44)
RANGE_2_START = 26  # Column AA
RANGE_2_END = 44    # Column AS (inclusive)

# Additional column W (0-indexed: 22) - LEFT PAVED SHOULDER
COLUMN_W_INDEX = 22  # Column W


# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def idx_to_excel_col(idx):
    """Convert 0-based column index to Excel column letter"""
    col = ''
    idx += 1
    while idx > 0:
        idx -= 1
        col = chr(65 + (idx % 26)) + col
        idx //= 26
    return col


# ============================================================================
# STEP 1: Read TCS_Input.xlsx and Create Dictionary
# ============================================================================

def create_tcs_dictionary(tcs_input_file):
    """
    Reads TCS_Input.xlsx and creates a dictionary with TCS type as key
    Extracts columns D-V, AA-AS, and W
    Returns: dictionary with TCS specifications
    """
    print("="*80)
    print("STEP 1: Creating TCS Dictionary (columns D-V, AA-AS, and W)")
    print("="*80)
    
    # Read the Excel file - first row for headers
    df_headers = pd.read_excel(tcs_input_file, sheet_name='Input', nrows=1)
    raw_headers = df_headers.iloc[0].tolist()
    
    # Read data starting from row 2 (skip header row)
    df = pd.read_excel(tcs_input_file, sheet_name='Input', skiprows=1, header=None)
    
    # Extract column ranges
    range_1_indices = list(range(RANGE_1_START, RANGE_1_END + 1))
    range_2_indices = list(range(RANGE_2_START, RANGE_2_END + 1))
    selected_indices = range_1_indices + range_2_indices
    
    print(f"[OK] Extracting columns:")
    print(f"  Range 1: D to V (indices {RANGE_1_START} to {RANGE_1_END}) = {len(range_1_indices)} columns")
    print(f"  Range 2: AA to AS (indices {RANGE_2_START} to {RANGE_2_END}) = {len(range_2_indices)} columns")
    print(f"  Column W (index {COLUMN_W_INDEX})")
    print(f"  Total columns to extract: {len(selected_indices) + 1}")
    
    # Get headers and handle duplicates
    temp_headers = []
    for idx in selected_indices:
        if idx < len(raw_headers):
            header = raw_headers[idx]
            if header and str(header) != 'nan' and pd.notna(header):
                temp_headers.append(str(header))
            else:
                temp_headers.append(None)
        else:
            temp_headers.append(None)
    
    # Check for duplicates
    from collections import Counter
    header_counts = Counter([h for h in temp_headers if h])
    
    # Add suffixes to duplicates
    selected_headers = []
    seen_counts = {}
    
    for i, idx in enumerate(selected_indices):
        header = temp_headers[i]
        
        if header:
            if header_counts[header] > 1:
                if header not in seen_counts:
                    seen_counts[header] = 0
                seen_counts[header] += 1
                
                # Add _LHS or _RHS suffix
                if seen_counts[header] == 1:
                    suffixed_header = f"{header}_LHS"
                else:
                    suffixed_header = f"{header}_RHS"
                
                selected_headers.append((idx, suffixed_header))
            else:
                selected_headers.append((idx, header))
        else:
            selected_headers.append((idx, None))
    
    print(f"\n[OK] Selected {len(selected_headers)} specification columns")
    
    # Create dictionary with TCS type as key
    cs_type_idx = 2  # Column C
    tcs_dict = {}
    
    for idx, row in df.iterrows():
        cs_type = row.iloc[cs_type_idx]
        
        if pd.notna(cs_type):
            row_dict = {}
            
            # Add specification columns
            for col_idx, header in selected_headers:
                if header:
                    value = row.iloc[col_idx]
                    if pd.isna(value):
                        row_dict[header] = None
                    elif isinstance(value, (int, float)):
                        row_dict[header] = value
                    else:
                        row_dict[header] = str(value)
            
            # Add column W value
            w_value = row.iloc[COLUMN_W_INDEX]
            if pd.isna(w_value):
                row_dict['COLUMN_W_VALUE'] = None
            elif isinstance(w_value, (int, float)):
                row_dict['COLUMN_W_VALUE'] = w_value
            else:
                row_dict['COLUMN_W_VALUE'] = str(w_value)
            
            tcs_dict[cs_type] = row_dict
    
    print(f"\n[OK] Created dictionary with {len(tcs_dict)} TCS types")
    print(f"[OK] Each type has {len(selected_headers) + 1} specifications (including column W)")
    
    return tcs_dict


# ============================================================================
# STEP 2: Populate main_carriageway_and_boq.xlsx with Specifications
# ============================================================================

def populate_specifications(main_carriageway_file, tcs_dict, output_file):
    """
    Reads collected data from tcs_schedule.py and adds TCS specifications
    """
    print("\n" + "="*80)
    print("STEP 2: Populating TCS Input Specifications")
    print("="*80)
    
    # Read from collected CSV (tcs_schedule has the latest data)
    import os as _os
    from pathlib import Path as _P
    
    session_data_dir = _os.getenv('SESSION_DATA_DIR')
    if session_data_dir:
        collect_base = _P(session_data_dir) / 'collected'
    else:
        session_id = _os.getenv('SESSION_ID', 'default')
        collect_base = _P(_os.getcwd()) / 'data' / 'sessions' / session_id / 'collected'
    
    # Try to read from tcs_schedule collected CSV
    tcs_schedule_csv = collect_base / 'tcs_schedule.csv'
    if tcs_schedule_csv.exists():
        df = pd.read_csv(tcs_schedule_csv)
        print(f"[OK] Read tcs_schedule.csv from collector: {len(df)} rows")
    else:
        print(f"[ERROR] tcs_schedule.csv not found at {tcs_schedule_csv}")
        print(f"[DEBUG] Available files: {list(collect_base.glob('*.csv')) if collect_base.exists() else 'directory does not exist'}")
        raise FileNotFoundError(f"tcs_schedule.csv not found in {collect_base}")
    
    # KEEP ONLY FIRST 4 COLUMNS (A, B, C, D)
    df = df.iloc[:, :4].copy()
    print(f"[OK] Kept first 4 columns (A-D): from, to, length, type_of_cross_section")
    
    # Create case-insensitive lookup dictionary
    tcs_dict_lower = {k.lower(): (k, v) for k, v in tcs_dict.items()}
    
    # Get specification column names (excluding COLUMN_W_VALUE)
    spec_columns = []
    for specs in tcs_dict.values():
        if specs:
            spec_columns = [k for k in specs.keys() if k != 'COLUMN_W_VALUE']
            break
    
    print(f"[OK] Will add {len(spec_columns)} specification columns starting from column E")
    print(f"[OK] Will add 2 placeholder columns")
    print(f"[OK] Will add 4 columns with column W values")
    
    # Create new columns for specifications
    spec_data = {col: [None] * len(df) for col in spec_columns}
    spec_data['AQ_PLACEHOLDER'] = [None] * len(df)
    spec_data['AR_PLACEHOLDER'] = [None] * len(df)
    spec_data['AS'] = [None] * len(df)
    spec_data['AT'] = [None] * len(df)
    spec_data['AU'] = [None] * len(df)
    spec_data['AV'] = [None] * len(df)
    
    # Fill in specifications
    print(f"\nFilling specifications...")
    matched_count = 0
    unmatched_types = set()
    type_counts = {}
    
    for idx in range(len(df)):
        cs_type = df.iloc[idx, 3]  # Column D (4th column, index 3)
        
        if pd.isna(cs_type):
            continue
            
        if cs_type not in type_counts:
            type_counts[cs_type] = 0
        type_counts[cs_type] += 1
        
        # Try exact match first, then case-insensitive
        if cs_type in tcs_dict:
            specs = tcs_dict[cs_type]
            matched_count += 1
        elif cs_type.lower() in tcs_dict_lower:
            original_key, specs = tcs_dict_lower[cs_type.lower()]
            matched_count += 1
        else:
            unmatched_types.add(cs_type)
            continue
        
        # Fill specification columns
        for col in spec_columns:
            if col in specs:
                spec_data[col][idx] = specs[col]
        
        # Fill column W values to AS, AT, AU, AV
        if 'COLUMN_W_VALUE' in specs:
            w_value = specs['COLUMN_W_VALUE']
            spec_data['AS'][idx] = w_value
            spec_data['AT'][idx] = w_value
            spec_data['AU'][idx] = w_value
            spec_data['AV'][idx] = w_value
        
        # Progress indicator
        if (idx + 1) % 200 == 0:
            print(f"  Processed {idx + 1}/{len(df)} rows...")
    
    print(f"\n[OK] Matched specifications for {matched_count}/{len(df)} rows")
    
    # Add specification columns to dataframe starting from column E (index 4)
    col_idx = 4  # Start from column E
    for col_name, col_data in spec_data.items():
        df[col_idx] = col_data
        col_idx += 1
    
    # Show summary
    print(f"\nDictionary lookup statistics:")
    for cs_type, count in sorted(type_counts.items(), key=lambda x: x[1], reverse=True)[:10]:
        status = "[OK]" if cs_type in tcs_dict or cs_type.lower() in tcs_dict_lower else "[ERROR]"
        print(f"  {status} {cs_type}: {count} rows")
    
    if unmatched_types:
        print(f"\n[WARNING] No specifications found for:")
        for cs_type in sorted(unmatched_types):
            print(f"    - {cs_type}")
    
    print(f"\n[OK] Final column count: {len(df.columns)}")
    print(f"  Columns A-D: Core data")
    print(f"  Columns E onwards: {len(spec_data)} specification columns")
    
    # MODIFIED: Store data in collector instead of writing to Excel
    print(f"\n[OK] Storing data for later writing (Quantity sheet, starting row 7, column A)...")
    
    collector = get_collector()
    collector.add_data('tcs_input', {
        'sheet_name': 'Quantity',
        'dataframe': df,
        'start_row': 6,  # Row 7 (0-indexed)
        'start_col': 0,  # Column A
        'write_header': False,
        'write_index': False
    })
    
    print(f"[OK] Successfully stored {len(df)} rows for writing")
    
    return df


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Main function to execute all steps"""
    print("\n" + "="*80)
    print("TCS SPECIFICATION POPULATOR - EXTENDED VERSION")
    print("Extracts columns D-V and AA-AS from TCS_Input.xlsx")
    print("PLUS column W mapped to Excel columns AS, AT, AU, AV")
    print("Writes to: output/main_carriageway.xlsx (Quantity sheet, row 7+)")
    print("="*80 + "\n")
    
    try:
        # Step 1: Create TCS dictionary
        tcs_dict = create_tcs_dictionary(TCS_INPUT_FILE)
        
        
        # Step 2: Populate main_carriageway.xlsx
        df = populate_specifications(MAIN_CARRIAGEWAY_FILE, tcs_dict, OUTPUT_EXCEL)
        
        # Success summary
        print("\n" + "="*80)
        print("SUCCESS! [OK]")
        print("="*80)
        print(f"Output file: {OUTPUT_EXCEL}")
        print(f"Sheet: Quantity")
        print(f"Starting row: 7")
        print(f"Total data rows: {len(df)}")
        print(f"Total columns: {len(df.columns)}")
        print(f"\nColumn layout:")
        print(f"  A-D: Core columns (from, to, length, type_of_cross_section)")
        print(f"  E onwards: Specification columns")
        
        # Note: File will be uploaded to GCS at the end of all processing in main.py
        # No need to upload here for efficiency
        
        # Cleanup - only remove temp files, not the local output file
        os.remove(TCS_INPUT_FILE)
        # Only remove OUTPUT_EXCEL if it's a temp file (not the local SESSION_OUTPUT_FILE)
        if OUTPUT_EXCEL != os.getenv('SESSION_OUTPUT_FILE', ''):
            try:
                os.remove(OUTPUT_EXCEL)
            except:
                pass
        
    except FileNotFoundError as e:
        print(f"\n[ERROR] File not found - {e}")
        print("Please check your file paths!")
    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
