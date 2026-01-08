"""
Embankment Height Processor
Reads Emb_Height.xlsx from row 5, columns A, E, F
Populates main_carriageway.xlsx Quantity sheet columns AQ and AR from row 7 onwards
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
emb_height_gcs = gcs.get_gcs_path(session_id, 'Emb Height.xlsx', 'data')
EMB_HEIGHT_FILE = gcs.download_to_temp(emb_height_gcs, suffix='.xlsx')


# ============================================================================
# STEP 1: Read Emb_Height.xlsx and Create Dictionary
# ============================================================================

def create_emb_height_dictionary(emb_height_file):
    """
    Reads Emb_Height.xlsx starting from Excel row 5
    Creates dictionary with Column A as key, Columns E and F as values
    Returns: dictionary {chainage: {'left': value_E, 'right': value_F}}
    """
    print("="*80)
    print("STEP 1: Creating Embankment Height Dictionary")
    print("="*80)
    
    # Read the Excel file
    df = pd.read_excel(emb_height_file)
    
    print("[OK] Read Emb_Height.xlsx:", len(df), "total rows")
    
    # Start from Excel row 5 (pandas index 4)
    data_start_row = 3
    
    # Extract data from row 5 onwards
    data = df.iloc[data_start_row:]
    
    print("[OK] Extracting data from Excel row 5 (pandas index", data_start_row, ")")
    print("[OK] Total data rows:", len(data))
    
    # Create dictionary with Column A as key
    emb_dict = {}
    skipped = 0
    
    for idx, row in data.iterrows():
        # Column A (index 0) = Key (chainage)
        key = row.iloc[0]
        
        # Column E (index 4) = Left height
        value_e = row.iloc[4]
        
        # Column F (index 5) = Right height
        value_f = row.iloc[5]
        
        # Skip if key is NaN
        if pd.notna(key):
            # Round key to 6 decimal places for consistent matching
            key_float = round(float(key), 6)
            emb_dict[key_float] = {
                'left': float(value_e) if pd.notna(value_e) else None,
                'right': float(value_f) if pd.notna(value_f) else None
            }
        else:
            skipped += 1
    
    print("[OK] Created dictionary with", len(emb_dict), "entries")
    if skipped > 0:
        print("  (Skipped", skipped, "rows with NaN keys)")
    
    # Show range
    if emb_dict:
        keys = sorted(emb_dict.keys())
        print("[OK] Key (chainage) range: %.3f to %.3f" % (keys[0], keys[-1]))
        print("\n  Sample entries:")
        for i, key in enumerate(keys[:3]):
            print("    %s: Left=%s, Right=%s" % (key, emb_dict[key]['left'], emb_dict[key]['right']))
    
    return emb_dict


# ============================================================================
# STEP 2: Populate main_carriageway_and_boq.xlsx with Embankment Heights
# ============================================================================

def populate_embankment_heights(main_carriageway_file, emb_dict, output_file):
    """
    Reads collected data from tcs_input.py and adds embankment height columns
    Matches Column A with dict keys
    Populates columns AQ (index 42) and AR (index 43) with embankment heights
    """
    print("\n" + "="*80)
    print("STEP 2: Populating Embankment Height Data")
    print("="*80)
    
    # Read from collected CSV (tcs_input has the latest data)
    import os as _os
    from pathlib import Path as _P
    
    session_data_dir = _os.getenv('SESSION_DATA_DIR')
    if session_data_dir:
        collect_base = _P(session_data_dir) / 'collected'
    else:
        session_id = _os.getenv('SESSION_ID', 'default')
        collect_base = _P(_os.getcwd()) / 'data' / 'sessions' / session_id / 'collected'
    
    # Try to read from tcs_input collected CSV
    tcs_input_csv = collect_base / 'tcs_input.csv'
    if tcs_input_csv.exists():
        df = pd.read_csv(tcs_input_csv)
        print(f"[OK] Read tcs_input.csv from collector: {len(df)} rows")
    else:
        print(f"[WARNING] tcs_input.csv not found at {tcs_input_csv}")
        print(f"[FALLBACK] Reading from tcs_schedule.csv...")
        tcs_schedule_csv = collect_base / 'tcs_schedule.csv'
        if tcs_schedule_csv.exists():
            df = pd.read_csv(tcs_schedule_csv)
            print(f"[OK] Read tcs_schedule.csv from collector: {len(df)} rows")
        else:
            print(f"[ERROR] No collected data found in {collect_base}")
            print(f"[DEBUG] Available files: {list(collect_base.glob('*.csv')) if collect_base.exists() else 'directory does not exist'}")
            raise FileNotFoundError(f"No collected CSV data found in {collect_base}")
    
    print("  Current columns:", len(df.columns))
    
    # Column AQ = index 42 (43rd column, 0-indexed)
    # Column AR = index 43 (44th column, 0-indexed)
    AQ_COL_INDEX = 42
    AR_COL_INDEX = 43
    
    # Ensure we have enough columns
    while len(df.columns) <= AR_COL_INDEX:
        df[len(df.columns)] = None
    
    print("\n[OK] Columns AQ (42) and AR (43) ready")
    print("  Total columns:", len(df.columns))
    
    print("\n[OK] Matching rows starting from index 0")
    
    # Match and populate
    matched = 0
    unmatched = 0
    
    for idx in range(len(df)):
        # Get the key from Column A (index 0)
        key = df.iloc[idx, 0]
        
        # Try to match in dictionary
        if pd.notna(key) and float(key) in emb_dict:
            heights = emb_dict[float(key)]
            # Write to column index 42 (AQ) and 43 (AR)
            df.iloc[idx, AQ_COL_INDEX] = heights['left']
            df.iloc[idx, AR_COL_INDEX] = heights['right']
            matched += 1
        else:
            # Fill with value from exactly upper cell (previous row)
            if idx > 0:
                df.iloc[idx, AQ_COL_INDEX] = df.iloc[idx - 1, AQ_COL_INDEX]
                df.iloc[idx, AR_COL_INDEX] = df.iloc[idx - 1, AR_COL_INDEX]
            else:
                # For the first row, set to None if unmatched
                df.iloc[idx, AQ_COL_INDEX] = None
                df.iloc[idx, AR_COL_INDEX] = None
            unmatched += 1
        
        # Progress indicator
        if (idx + 1) % 200 == 0:
            print("  Processed %d/%d rows..." % (idx + 1, len(df)))
    
    print("\n[OK] Matching complete:")
    print("  Matched:", matched)
    print("  Unmatched:", unmatched)
    
    if matched > 0:
        match_pct = matched / (matched + unmatched) * 100
        print("  Match rate: %.1f%%" % match_pct)
    
    # MODIFIED: Store data in collector instead of writing to Excel
    print("\n[OK] Storing data for later writing (Quantity sheet, starting row 7)...")
    
    collector = get_collector()
    collector.add_data('emb_height', {
        'sheet_name': 'Quantity',
        'dataframe': df,
        'start_row': 6,  # Row 7 (0-indexed)
        'start_col': 0,  # Column A
        'write_header': False,
        'write_index': False
    })
    
    print("[OK] Successfully stored", len(df), "rows for writing")
    
    return df, matched, unmatched


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Main function to execute all steps"""
    print("\n" + "="*80)
    print("EMBANKMENT HEIGHT PROCESSOR")
    print("="*80)
    print("Configuration:")
    print("  • Emb_Height.xlsx: Read from Excel row 5 (Col A, E, F)")
    print("  • main_carriageway_and_boq.xlsx: Quantity sheet, starting row 7")
    print("  • Output: Populate columns AQ and AR")
    print("="*80 + "\n")
    
    # Define column indices at the top
    AQ_COL_INDEX = 42
    AR_COL_INDEX = 43
    
    try:
        # Step 1: Create embankment height dictionary
        emb_dict = create_emb_height_dictionary(EMB_HEIGHT_FILE)
        
        # Step 2: Populate main_carriageway.xlsx
        df, matched, unmatched = populate_embankment_heights(
            MAIN_CARRIAGEWAY_FILE, emb_dict, OUTPUT_EXCEL
        )
        
        # Success summary
        print("\n" + "="*80)
        print("SUCCESS! [OK]")
        print("="*80)
        print("Output file:", OUTPUT_EXCEL)
        print("Sheet: Quantity")
        print("Starting row: 7")
        print("Total data rows:", len(df))
        print("Total columns:", len(df.columns))
        
        # Show sample output
        print("\n" + "="*80)
        print("SAMPLE OUTPUT:")
        print("-"*80)
        
        # Find rows with embankment heights populated
        rows_with_heights = df[df.iloc[:, AQ_COL_INDEX].notna()]
        
        if len(rows_with_heights) > 0:
            print("\nFound", len(rows_with_heights), "rows with embankment heights")
            print("\nFirst 3 rows with heights:")
            
            for i in range(min(3, len(rows_with_heights))):
                idx = rows_with_heights.index[i]
                row = rows_with_heights.iloc[i]
                print("\n  Data row %d (Excel row %d):" % (i + 1, idx + 7))
                print("    Column A (from):", row.iloc[0])
                print("    Column D (type):", row.iloc[3] if len(row) > 3 else 'N/A')
                print("    Column AQ (Emb_Height_Left):", row.iloc[AQ_COL_INDEX])
                print("    Column AR (Emb_Height_Right):", row.iloc[AR_COL_INDEX])
        else:
            print("\n[WARNING] No matching chainages found")
            print("Possible reasons:")
            print("  1. Chainage ranges don't overlap")
            print("  2. Key values don't match exactly")
            print("\nColumns AQ and AR have been added but remain empty")
        
        # Column layout
        print("\n" + "="*80)
        print("FINAL COLUMN LAYOUT:")
        print("-"*80)
        print("  Columns A-AP: First 42 columns")
        print("  Column AQ (index 42): Emb_Height_Left")
        print("  Column AR (index 43): Emb_Height_Right")
        print("  Total columns:", len(df.columns))
        
        # Note: File will be uploaded to GCS at the end of all processing in main.py
        # No need to upload here for efficiency
        
        # Cleanup - only remove temp files, not the local output file
        os.remove(EMB_HEIGHT_FILE)
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
        print("  1. Files exist in the correct folders")
        print("  2. File names match exactly")
    except Exception as e:
        print("\n[ERROR]", e)
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
