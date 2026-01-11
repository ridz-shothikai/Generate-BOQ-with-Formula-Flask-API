"""
Geogrid Calculator
Reads Pavement_Input.xlsx to check for Geogrid conditions
Calculates columns KY, KZ, LA, LB in main_carriageway.xlsx based on formulas
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
pavement_gcs = gcs.get_gcs_path(session_id, 'Pavement Input.xlsx', 'data')
PAVEMENT_INPUT_FILE = gcs.download_to_temp(pavement_gcs, suffix='.xlsx')


# ============================================================================
# STEP 1: Check Pavement_Input.xlsx for Geogrid Conditions
# ============================================================================

def check_geogrid_conditions(pavement_input_file):
    """
    Reads Pavement_Input.xlsx and checks if:
    - E9 contains "Geogrid Reinforced GSB"
    - E10 contains "Geogrid Reinforced WMM"
    - B9 contains "Geogrid Reinforced GSB"
    - B10 contains "Geogrid Reinforced WMM"
    Returns: dictionary with boolean flags
    """
    print("="*80)
    print("STEP 1: Checking Geogrid Conditions in Pavement Input")
    print("="*80)
    
    # Read the Excel file without headers
    df = pd.read_excel(pavement_input_file, header=None)
    
    print("[OK] Read Pavement_Input.xlsx:", len(df), "total rows")
    
    # Initialize conditions
    conditions = {
        'e9_geogrid_gsb': False,
        'e10_geogrid_wmm': False,
        'b9_geogrid_gsb': False,
        'b10_geogrid_wmm': False
    }
    
    # Check E9 (row index 8, column 4)
    if len(df) > 8:
        e9_value = df.iloc[8, 4]  # E9
        if pd.notna(e9_value) and "Geogrid Reinforced GSB" in str(e9_value):
            conditions['e9_geogrid_gsb'] = True
            print(f"[OK] E9 contains 'Geogrid Reinforced GSB': {e9_value}")
        else:
            print(f"[OK] E9 value: {e9_value} (not Geogrid Reinforced GSB)")
    
    # Check E10 (row index 9, column 4)
    if len(df) > 9:
        e10_value = df.iloc[9, 4]  # E10
        if pd.notna(e10_value) and "Geogrid Reinforced WMM" in str(e10_value):
            conditions['e10_geogrid_wmm'] = True
            print(f"[OK] E10 contains 'Geogrid Reinforced WMM': {e10_value}")
        else:
            print(f"[OK] E10 value: {e10_value} (not Geogrid Reinforced WMM)")
    
    # Check B9 (row index 8, column 1)
    if len(df) > 8:
        b9_value = df.iloc[8, 1]  # B9
        if pd.notna(b9_value) and "Geogrid Reinforced GSB" in str(b9_value):
            conditions['b9_geogrid_gsb'] = True
            print(f"[OK] B9 contains 'Geogrid Reinforced GSB': {b9_value}")
        else:
            print(f"[OK] B9 value: {b9_value} (not Geogrid Reinforced GSB)")
    
    # Check B10 (row index 9, column 1)
    if len(df) > 9:
        b10_value = df.iloc[9, 1]  # B10
        if pd.notna(b10_value) and "Geogrid Reinforced WMM" in str(b10_value):
            conditions['b10_geogrid_wmm'] = True
            print(f"[OK] B10 contains 'Geogrid Reinforced WMM': {b10_value}")
        else:
            print(f"[OK] B10 value: {b10_value} (not Geogrid Reinforced WMM)")
    
    print("\n[OK] Geogrid Conditions Summary:")
    print(f"  E9 Geogrid GSB: {conditions['e9_geogrid_gsb']}")
    print(f"  E10 Geogrid WMM: {conditions['e10_geogrid_wmm']}")
    print(f"  B9 Geogrid GSB: {conditions['b9_geogrid_gsb']}")
    print(f"  B10 Geogrid WMM: {conditions['b10_geogrid_wmm']}")
    
    return conditions


# ============================================================================
# STEP 2: Calculate Geogrid Columns
# ============================================================================

def find_last_row_with_data(ws, column_letter):
    """
    Find the last row with data in the specified column
    Args:
        ws: Worksheet object
        column_letter: Excel column letter (e.g., 'D')
    Returns:
        Last row number that contains data
    """
    from openpyxl.utils import column_index_from_string
    
    col_idx = column_index_from_string(column_letter)
    last_row = 0
    
    # Iterate from bottom to top to find last non-empty cell
    for row_idx in range(ws.max_row, 0, -1):
        cell_value = ws.cell(row_idx, col_idx).value
        if cell_value is not None and cell_value != '':
            last_row = row_idx
            break
    
    return last_row

def calculate_geogrid_columns(main_carriageway_file, conditions, output_file):
    """
    Reads main_carriageway_and_boq.xlsx and calculates geogrid columns based on conditions
    Updates only columns KY, KZ, LA, LB without touching other columns
    """
    print("\n" + "="*80)
    print("STEP 2: Calculating Geogrid Columns")
    print("="*80)
    
    # Prefer latest collected dataframe instead of reading workbook (single write design)
    import pandas as _pd
    from pathlib import Path as _P
    import os as _os
    session_data_dir = _os.getenv('SESSION_DATA_DIR')
    if session_data_dir:
        collect_base = _P(session_data_dir) / 'collected'
    else:
        collect_base = _P(_os.getcwd()) / 'data' / 'sessions' / _os.getenv('SESSION_ID', 'default') / 'collected'
    
    # CRITICAL: Load from formula_applier.csv which has all calculated columns (DL, DS, EE, EJ, FF, FD, FY, FS)
    # These columns are needed for geogrid calculations
    csv_candidates = [
        collect_base / 'formula_applier.csv',  # PRIMARY: has all formula results needed
        collect_base / 'constant_fill.csv',    # Fallback 1
        collect_base / 'pavement_input.csv',   # Fallback 2
        collect_base / 'tcs_input.csv',        # Fallback 3
    ]
    df = None
    loaded_csv = None
    for csv_path in csv_candidates:
        if csv_path.exists():
            try:
                df = _pd.read_csv(csv_path)
                loaded_csv = csv_path.name
                print(f"[GeogridCalculator] Loaded collected CSV: {csv_path.name} with shape {df.shape}")
                print(f"[GeogridCalculator] DataFrame has {len(df.columns)} columns")
                break
            except Exception as e:
                print(f"[GeogridCalculator] Error reading {csv_path.name}: {e}")
                continue
    
    if df is None:
        print(f"[GeogridCalculator] No collected CSV found in {collect_base}")
        if collect_base.exists():
            files = list(collect_base.glob('*.csv'))
            print(f"[GeogridCalculator] Available CSVs: {[f.name for f in files]}")
        raise FileNotFoundError("No collected CSV found for geogrid calculation")
    
    print(f"[OK] Loaded collected dataframe for geogrid from '{loaded_csv}': shape={df.shape}")
    last_row_with_data = len(df) + 6  # data begins at Excel row 7
    
    # Column letters (Excel columns, 1-indexed)
    LENGTH_COL = 3      # Column C
    DL_COL = 116        # Column DL
    DS_COL = 123        # Column DS
    EE_COL = 135        # Column EE
    EJ_COL = 140        # Column EJ
    FD_COL = 160        # Column FD
    FF_COL = 162        # Column FF
    FS_COL = 175        # Column FS
    FY_COL = 181        # Column FY
    
    KY_COL = 311        # Column KY
    KZ_COL = 312        # Column KZ
    LA_COL = 313        # Column LA
    LB_COL = 314        # Column LB
    
    # Check if required columns exist in the loaded DataFrame
    required_cols = [DL_COL, DS_COL, EE_COL, EJ_COL, FF_COL, FD_COL, FY_COL, FS_COL]
    missing_cols = [col for col in required_cols if col >= len(df.columns)]
    if missing_cols:
        print(f"[WARNING] DataFrame has only {len(df.columns)} columns, but need columns up to {max(required_cols)}")
        print(f"[WARNING] Missing columns at indices: {missing_cols}")
        print(f"[WARNING] This CSV may not have formula-calculated columns. Need formula_applier.csv!")
        print(f"[WARNING] Geogrid values will be 0 because source columns are missing.")
    
    # Data starts from row 7
    start_row = 7
    
    print(f"\n[OK] Calculating geogrid values from row {start_row} to row {last_row_with_data}...")
    print(f"[OK] Active conditions: E9_GSB={conditions['e9_geogrid_gsb']}, E10_WMM={conditions['e10_geogrid_wmm']}, B9_GSB={conditions['b9_geogrid_gsb']}, B10_WMM={conditions['b10_geogrid_wmm']}")
    print(f"[OK] DataFrame has {len(df.columns)} columns available for reading")
    
    row_count = 0
    modifications = []
    non_zero_count = 0
    
    for row_idx in range(start_row, last_row_with_data + 1):
        # Map Excel row to DataFrame index
        df_idx = row_idx - 7
        length = df.iloc[df_idx, LENGTH_COL] if 0 <= df_idx < len(df) else 0
        
        # Skip empty rows (if length column is empty, skip this row)
        if length == 0 or length is None or length == '':
            continue
        
        # Get column values from DataFrame (with column bounds checking)
        def _safe(v):
            """Convert value to float, return 0 if not possible"""
            if _pd.isna(v):
                return 0.0
            try:
                return float(v)
            except (ValueError, TypeError):
                return 0.0
        
        def _safe_col(row_idx, col_idx):
            """Safely get column value, return 0.0 if column doesn't exist"""
            if 0 <= row_idx < len(df) and 0 <= col_idx < len(df.columns):
                return _safe(df.iloc[row_idx, col_idx])
            return 0.0
        
        # Convert length to float
        try:
            length = float(length) if length not in [None, '', 0] else 0.0
        except (ValueError, TypeError):
            length = 0.0
        
        dl_val = _safe_col(df_idx, DL_COL)
        ds_val = _safe_col(df_idx, DS_COL)
        ee_val = _safe_col(df_idx, EE_COL)
        ej_val = _safe_col(df_idx, EJ_COL)
        ff_val = _safe_col(df_idx, FF_COL)
        fd_val = _safe_col(df_idx, FD_COL)
        fy_val = _safe_col(df_idx, FY_COL)
        fs_val = _safe_col(df_idx, FS_COL)
        
        # Calculate KY (ensure float result)
        ky_val = float(((dl_val if conditions['e9_geogrid_gsb'] else 0.0) + 
                        (ds_val if conditions['e10_geogrid_wmm'] else 0.0)) * length)
        modifications.append({'sheet': 'Quantity', 'row': row_idx, 'col': KY_COL, 'value': ky_val})
        
        # Calculate KZ
        kz_val = float(((ee_val if conditions['b9_geogrid_gsb'] else 0.0) + 
                        (ej_val if conditions['b10_geogrid_wmm'] else 0.0)) * length)
        modifications.append({'sheet': 'Quantity', 'row': row_idx, 'col': KZ_COL, 'value': kz_val})
        
        # Calculate LA
        la_val = float(((ff_val if conditions['b9_geogrid_gsb'] else 0.0) + 
                        (fd_val if conditions['b10_geogrid_wmm'] else 0.0)) * length)
        modifications.append({'sheet': 'Quantity', 'row': row_idx, 'col': LA_COL, 'value': la_val})
        
        # Calculate LB
        lb_val = float(((fy_val if conditions['e9_geogrid_gsb'] else 0.0) + 
                        (fs_val if conditions['e10_geogrid_wmm'] else 0.0)) * length)
        modifications.append({'sheet': 'Quantity', 'row': row_idx, 'col': LB_COL, 'value': lb_val})
        
        # Track non-zero values
        if ky_val != 0 or kz_val != 0 or la_val != 0 or lb_val != 0:
            non_zero_count += 1
            # Log first few non-zero calculations for debugging
            if non_zero_count <= 3:
                print(f"  Row {row_idx}: KY={ky_val:.2f}, KZ={kz_val:.2f}, LA={la_val:.2f}, LB={lb_val:.2f}")
                print(f"    Source: DL={dl_val:.2f}, DS={ds_val:.2f}, EE={ee_val:.2f}, EJ={ej_val:.2f}, FF={ff_val:.2f}, FD={fd_val:.2f}, FY={fy_val:.2f}, FS={fs_val:.2f}, Length={length:.2f}")
        
        row_count += 1
        
        # Progress indicator
        if row_count % 200 == 0:
            print(f"  Processed {row_count} rows...")
    
    print(f"[OK] Geogrid calculations completed for {row_count} rows")
    print(f"[OK] Non-zero geogrid values found in {non_zero_count} rows")
    
    if non_zero_count == 0:
        print("[WARNING] All geogrid values are 0! This may indicate:")
        print("  1. No geogrid conditions are active (all conditions are False)")
        print("  2. Source columns (DL, DS, EE, EJ, FF, FD, FY, FS) are empty/zero")
        print("  3. Wrong CSV loaded (need formula_applier.csv with calculated values)")
    
    # Store modifications via collector for single-write later
    collector = get_collector()
    collector.add_workbook_modifications('pavement_input_with_internal', modifications)
    print(f"[OK] Stored {len(modifications)} geogrid modifications for later writing")
    
    return row_count


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Main function to execute all steps"""
    print("\n" + "="*80)
    print("GEOGRID CALCULATOR")
    print("="*80)
    print("Configuration:")
    print("  • Pavement_Input.xlsx: Check for Geogrid conditions")
    print("  • main_carriageway_and_boq.xlsx: Calculate columns KY, KZ, LA, LB")
    print("="*80 + "\n")
    
    try:
        # Step 1: Check geogrid conditions
        conditions = check_geogrid_conditions(PAVEMENT_INPUT_FILE)
        
        # Step 2: Calculate geogrid columns
        row_count = calculate_geogrid_columns(MAIN_CARRIAGEWAY_FILE, conditions, OUTPUT_EXCEL)
        
        # Success summary
        print("\n" + "="*80)
        print("SUCCESS! [OK]")
        print("="*80)
        print("Output file:", OUTPUT_EXCEL)
        print("Total rows processed:", row_count)
        print("Columns updated: KY, KZ, LA, LB")
        
        # Note: File will be uploaded to GCS at the end of all processing in main.py
        # No need to upload here for efficiency
        
        # Cleanup - only remove temp files, not the local output file
        os.remove(PAVEMENT_INPUT_FILE)
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
        print("  1. Files exist in the data folder")
        print("  2. File names match exactly")
    except Exception as e:
        print("\n[ERROR]:", e)
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
