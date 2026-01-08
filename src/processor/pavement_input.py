"""
Pavement Input Processor
Reads Pavement_Input.xlsx from row 9, columns B, C, E, F
Populates main_carriageway_and_boq.xlsx column AX from row 1 onwards
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
# FILE PATHS - GCS Version
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
pavement_gcs_path = gcs.get_gcs_path(session_id, 'Pavement Input.xlsx', 'data')
PAVEMENT_INPUT_FILE = gcs.download_to_temp(pavement_gcs_path, suffix='.xlsx')


# ============================================================================
# STEP 1: Read Pavement_Input.xlsx and Create Dictionary
# ============================================================================

def create_pavement_dictionary(pavement_input_file):
    """
    Reads Pavement_Input.xlsx starting from Excel row 9
    Creates dictionary with:
      - Keys from Column B (with cell ref + value + B1 suffix)
      - Keys from Column E (with cell ref + value + E1 suffix)
      - Values from Column C (for B keys) and Column F (for E keys)
    Returns: dictionary
    """
    print("="*80)
    print("STEP 1: Creating Pavement Dictionary")
    print("="*80)
    
    # Read the Excel file without headers
    df = pd.read_excel(pavement_input_file, header=None)
    
    print("[OK] Read Pavement_Input.xlsx:", len(df), "total rows")
    
    # Get suffixes from row 1 (index 0)
    suffix_B = df.iloc[0, 1]  # B1 = "MCW"
    suffix_E = df.iloc[0, 4]  # E1 = "SR"
    
    print(f"[OK] Suffix from B1: '{suffix_B}'")
    print(f"[OK] Suffix from E1: '{suffix_E}'")
    
    # Create dictionary
    pavement_dict = {}
    
    # Process from row 9 onwards (index 8)
    data_start_row = 8
    
    print(f"\n[OK] Processing from Excel row 9 (index {data_start_row})")
    
    # Process Column B and C
    for i in range(data_start_row, len(df)):
        b_value = df.iloc[i, 1]  # Column B
        c_value = df.iloc[i, 2]  # Column C
        
        if pd.notna(b_value):
            excel_row = i + 1
            key = f"B{excel_row}_{b_value}_{suffix_B}"
            pavement_dict[key] = c_value if pd.notna(c_value) else 0
    
    # Process Column E and F
    for i in range(data_start_row, len(df)):
        e_value = df.iloc[i, 4]  # Column E
        f_value = df.iloc[i, 5]  # Column F
        
        if pd.notna(e_value):
            excel_row = i + 1
            key = f"E{excel_row}_{e_value}_{suffix_E}"
            pavement_dict[key] = f_value if pd.notna(f_value) else 0
    
    print(f"[OK] Created dictionary with {len(pavement_dict)} entries")
    
    # Show sample entries
    print("\n  Sample entries:")
    for i, (key, value) in enumerate(list(pavement_dict.items())[:5]):
        print(f"    {key}: {value}")
    
    return pavement_dict


# ============================================================================
# STEP 2: Populate main_carriageway_and_boq.xlsx Column AX
# ============================================================================

def populate_columns(main_carriageway_file, pavement_dict, output_file):
    """
    Reads collected data from emb_height.py and adds pavement columns
    Populates columns based on formula
    """
    print("\n" + "="*80)
    print("STEP 2: Populating Pavement Input Data")
    print("="*80)
    
    # Read from collected CSV (emb_height has the latest data)
    import os as _os
    from pathlib import Path as _P
    
    session_data_dir = _os.getenv('SESSION_DATA_DIR')
    if session_data_dir:
        collect_base = _P(session_data_dir) / 'collected'
    else:
        session_id = _os.getenv('SESSION_ID', 'default')
        collect_base = _P(_os.getcwd()) / 'data' / 'sessions' / session_id / 'collected'
    
    # Try to read from emb_height collected CSV
    emb_height_csv = collect_base / 'emb_height.csv'
    if emb_height_csv.exists():
        df = pd.read_csv(emb_height_csv)
        print(f"[OK] Read emb_height.csv from collector: {len(df)} rows")
    else:
        print(f"[WARNING] emb_height.csv not found at {emb_height_csv}")
        print(f"[FALLBACK] Reading from tcs_input.csv...")
        tcs_input_csv = collect_base / 'tcs_input.csv'
        if tcs_input_csv.exists():
            df = pd.read_csv(tcs_input_csv)
            print(f"[OK] Read tcs_input.csv from collector: {len(df)} rows")
        else:
            print(f"[ERROR] No collected data found in {collect_base}")
            print(f"[DEBUG] Available files: {list(collect_base.glob('*.csv')) if collect_base.exists() else 'directory does not exist'}")
            raise FileNotFoundError(f"No collected CSV data found in {collect_base}")
    print("[OK] Read collected data:", len(df), "rows")
    print("  Current columns:", len(df.columns))
    
    # Column AX = index 49
    AX_COL_INDEX = 49
    
    # Check if E9 contains "CTSB" (formula: IF E9="CTSB" THEN F9/1000 ELSE 0)
    ctsb_value = None
    ctsb_key = None
    
    for key, value in pavement_dict.items():
        if key.startswith('E9_CTSB_'):
            ctsb_value = value
            ctsb_key = key
            break
    
    if ctsb_value is not None:
        ax_value = ctsb_value / 1000
        print(f"[OK] Found CTSB in dictionary: {ctsb_key} = {ctsb_value}")
        print(f"  Column AX value will be: {ax_value}")
    else:
        ax_value = 0
        print("[OK] No CTSB found in dictionary")
        print("  Column AX value will be: 0")
        
    # Column AZ = index 50
    AZ_COL_INDEX = 51
    ctb_value = None
    ctb_key = None
    for key, value in pavement_dict.items():
        if key.startswith('E10_CTB_'):
            ctb_value = value
            ctb_key = key
            break
    if ctb_value is not None:
        az_value = ctb_value / 1000
        print(f"[OK] Found CTB in dictionary: {ctb_key} = {ctb_value}")
        print(f"  Column AZ value will be: {az_value}")
    else:
        az_value = 0
        print("[OK] No CTB found in dictionary")
        print("  Column AZ value will be: 0")

    # Column BA = Complex WMM formula
    # IF E10="WMM" OR E10="Geogrid Reinforced WMM" THEN F10/1000
    # ELSE IF E11="WMM" THEN F11/1000
    # ELSE 0
    BA_COL_INDEX = 52
    ba_value = 0
    ba_found = False
    
    # First check: E10 for "WMM" or "Geogrid Reinforced WMM"
    for key, value in pavement_dict.items():
        if key.startswith('E10_'):
            parts = key.split('_')
            if len(parts) >= 2:
                layer_name = parts[1]
                # Check if layer_name is exactly "WMM" or contains "Geogrid Reinforced WMM"
                if layer_name == "WMM" or layer_name == "Geogrid Reinforced WMM":
                    ba_value = value / 1000 if pd.notna(value) and value != 0 else 0
                    print(f"✓ Found E10 '{layer_name}' in dictionary: {key} = {value}, BA value = {ba_value}")
                    ba_found = True
                    break
    
    # Second check: If not found in E10, check E11 for "WMM"
    if not ba_found:
        for key, value in pavement_dict.items():
            if key.startswith('E11_'):
                parts = key.split('_')
                if len(parts) >= 2:
                    layer_name = parts[1]
                    if layer_name == "WMM":
                        ba_value = value / 1000 if pd.notna(value) and value != 0 else 0
                        print(f"✓ Found E11 'WMM' in dictionary: {key} = {value}, BA value = {ba_value}")
                        ba_found = True
                        break
    
    if not ba_found:
        print("✓ No WMM found in E10 or E11, BA value = 0")
        
    # Column BB = F11/1000 (Check E11 keys, as E column has F values)
    BB_COL_INDEX = 53
    bb_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('E11_'):
            bb_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found E11 in dictionary: {key} = {value}, BB value = {bb_value}")
            break
    if bb_value == 0:
        print("[OK] No E11 found in dictionary, BB value = 0")
    
    # Column BC = C23/1000 (Check B23 keys, as B column has C values)
    BC_COL_INDEX = 54
    bc_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B23_'):
            bc_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B23 in dictionary: {key} = {value}, BC value = {bc_value}")
            break
    if bc_value == 0:
        print("[OK] No B23 found in dictionary, BC value = 0")
    
    # Column BD = C24/1000 (Check B24 keys, as B column has C values)
    BD_COL_INDEX = 55
    bd_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B24_'):
            bd_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B24 in dictionary: {key} = {value}, BD value = {bd_value}")
            break
    if bd_value == 0:
        print("[OK] No B24 found in dictionary, BD value = 0")
        
    # Column BE = IF E10="RAP" THEN F10/1000 ELSE 0
    BE_COL_INDEX = 56
    be_value = 0
    
    for key, value in pavement_dict.items():
        if key.startswith('E10_RAP_'):
            be_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"✓ Found E10 RAP in dictionary: {key} = {value}, BE value = {be_value}")
            break
    
    if be_value == 0:
        print("✓ No RAP found in E10, BE value = 0")
        
    # Column BF = IF E12="BM" THEN F12/1000 ELSE 0
    BF_COL_INDEX = 57
    bf_value = 0
    
    for key, value in pavement_dict.items():
        if key.startswith('E12_BM_'):
            bf_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"✓ Found E12 BM in dictionary: {key} = {value}, BF value = {bf_value}")
            break
    
    if bf_value == 0:
        print("✓ No BM found in E12, BF value = 0")
    
    # Column BG = IF E13="DBM" THEN F13/1000 ELSE 0
    BG_COL_INDEX = 58
    bg_value = 0
    
    # Column BH = IF E14="PC&SC" THEN F14/1000 ELSE 0 (with error handling)
    BH_COL_INDEX = 59
    bh_value = 0
    
    for key, value in pavement_dict.items():
        if key.startswith('E14_PC&SC_'):
            bh_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"✓ Found E14 PC&SC in dictionary: {key} = {value}, BH value = {bh_value}")
            break
    
    if bh_value == 0:
        print("✓ No PC&SC found in E14, BH value = 0")
    
    # Column BI = IF E15="SDBC" THEN F15/1000 ELSE 0
    BI_COL_INDEX = 60
    bi_value = 0
    
    for key, value in pavement_dict.items():
        if key.startswith('E15_SDBC_'):
            bi_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"✓ Found E15 SDBC in dictionary: {key} = {value}, BI value = {bi_value}")
            break
    
    if bi_value == 0:
        print("✓ No SDBC found in E15, BI value = 0")
    
    # Column BJ = IF E16="BC" THEN F16/1000 ELSE 0
    BJ_COL_INDEX = 61
    bj_value = 0
    
    for key, value in pavement_dict.items():
        if key.startswith('E16_BC_'):
            bj_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"✓ Found E16 BC in dictionary: {key} = {value}, BJ value = {bj_value}")
            break
    
    if bj_value == 0:
        print("✓ No BC found in E16, BJ value = 0")
    
    for key, value in pavement_dict.items():
        if key.startswith('E13_DBM_'):
            bg_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"✓ Found E13 DBM in dictionary: {key} = {value}, BG value = {bg_value}")
            break
    
    if bg_value == 0:
        print("✓ No DBM found in E13, BG value = 0")
    
    # NEW CODE (without column naming) for AX
    if len(df.columns) <= AX_COL_INDEX:
        while len(df.columns) < AX_COL_INDEX:
            df[f'Empty_{len(df.columns)}'] = None
        df.insert(AX_COL_INDEX, f'Col_{AX_COL_INDEX}', ax_value)  # Simple column name
    else:
        df.iloc[:, AX_COL_INDEX] = ax_value
    
    print(f"\n[OK] Column AX (index {AX_COL_INDEX}) set to: {ax_value}")
    print(f"  Total columns: {len(df.columns)}")
    
    # Column AY = IF(BD>0, C22/1000, IF(E9="GSB" OR E9="Geogrid Reinforced GSB", F9/1000, 0))
    AY_COL_INDEX = 50
    ay_value = 0
    
    # First check: If BD (PQC) > 0, then use C22/1000
    if bd_value > 0:
        for key, value in pavement_dict.items():
            if key.startswith('B22_'):
                ay_value = value / 1000 if pd.notna(value) and value != 0 else 0
                print(f"[OK] BD > 0, Found B22 in dictionary: {key} = {value}, AY value = {ay_value}")
                break
        if ay_value == 0:
            print("[OK] BD > 0 but no B22 found in dictionary, AY value = 0")
    else:
        # Second check: If BD = 0, check E9 for "GSB" or "Geogrid Reinforced GSB"
        for key, value in pavement_dict.items():
            if key.startswith('E9_'):
                parts = key.split('_')
                if len(parts) >= 2:
                    layer_name = parts[1]
                    if layer_name == "GSB" or layer_name == "Geogrid Reinforced GSB":
                        ay_value = value / 1000 if pd.notna(value) and value != 0 else 0
                        print(f"[OK] BD = 0, Found E9 '{layer_name}' in dictionary: {key} = {value}, AY value = {ay_value}")
                        break
        if ay_value == 0:
            print("[OK] BD = 0 and no GSB found in E9, AY value = 0")
    
    # Ensure column AZ exists
    if len(df.columns) <= AZ_COL_INDEX:
        while len(df.columns) < AZ_COL_INDEX:
            df[f'Empty_{len(df.columns)}'] = None
        df.insert(AZ_COL_INDEX, f'Col_{AZ_COL_INDEX}', az_value)
    else:
        df.iloc[:, AZ_COL_INDEX] = az_value
    print(f"[OK] Column AZ (index {AZ_COL_INDEX}) set to: {az_value}")
    print(f"  Total columns: {len(df.columns)}")
    
    # Column BL = IF B9="CTSB" THEN C9/1000 ELSE 0
    BL_COL_INDEX = 63
    bl_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B9_CTSB_'):
            bl_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B9 CTSB in dictionary: {key} = {value}, BL value = {bl_value}")
            break
    if bl_value == 0:
        print("[OK] No CTSB found in B9, BL value = 0")
    
    # Column BN = IF B10="CTB" THEN C10/1000 ELSE 0
    BN_COL_INDEX = 65
    bn_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B10_CTB_'):
            bn_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B10 CTB in dictionary: {key} = {value}, BN value = {bn_value}")
            break
    if bn_value == 0:
        print("[OK] No CTB found in B10, BN value = 0")
    
    # Column BO = IF B10="WMM" OR B10="Geogrid Reinforced WMM" THEN C10/1000, ELSE IF B11="WMM" THEN C11/1000 ELSE 0
    BO_COL_INDEX = 66
    bo_value = 0
    bo_found = False
    
    # First check B10
    for key, value in pavement_dict.items():
        if key.startswith('B10_'):
            parts = key.split('_')
            if len(parts) >= 2:
                layer_name = parts[1]
                if layer_name == "WMM" or layer_name == "Geogrid Reinforced WMM":
                    bo_value = value / 1000 if pd.notna(value) and value != 0 else 0
                    print(f"[OK] Found B10 '{layer_name}' in dictionary: {key} = {value}, BO value = {bo_value}")
                    bo_found = True
                    break
    
    # If not found in B10, check B11
    if not bo_found:
        for key, value in pavement_dict.items():
            if key.startswith('B11_'):
                parts = key.split('_')
                if len(parts) >= 2:
                    layer_name = parts[1]
                    if layer_name == "WMM":
                        bo_value = value / 1000 if pd.notna(value) and value != 0 else 0
                        print(f"[OK] Found B11 'WMM' in dictionary: {key} = {value}, BO value = {bo_value}")
                        bo_found = True
                        break
    
    if not bo_found:
        print("[OK] No WMM found in B10 or B11, BO value = 0")
    
    # Column BP = C11/1000
    BP_COL_INDEX = 67
    bp_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B11_'):
            bp_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B11 in dictionary: {key} = {value}, BP value = {bp_value}")
            break
    if bp_value == 0:
        print("[OK] No B11 found in dictionary, BP value = 0")
    
    # Column BQ = C23/1000
    BQ_COL_INDEX = 68
    bq_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B23_'):
            bq_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B23 in dictionary: {key} = {value}, BQ value = {bq_value}")
            break
    if bq_value == 0:
        print("[OK] No B23 found in dictionary, BQ value = 0")
    
    # Column BR = C24/1000
    BR_COL_INDEX = 69
    br_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B24_'):
            br_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B24 in dictionary: {key} = {value}, BR value = {br_value}")
            break
    if br_value == 0:
        print("[OK] No B24 found in dictionary, BR value = 0")
    
    # Column BM = IF(BR>0, C22/1000, IF(B9="GSB" OR B9="Geogrid Reinforced GSB", C9/1000, 0))
    BM_COL_INDEX = 64
    bm_value = 0
    
    # First check: If BR (RHS_PQC) > 0, then use C22/1000
    if br_value > 0:
        for key, value in pavement_dict.items():
            if key.startswith('B22_'):
                bm_value = value / 1000 if pd.notna(value) and value != 0 else 0
                print(f"[OK] BR > 0, Found B22 in dictionary: {key} = {value}, BM value = {bm_value}")
                break
        if bm_value == 0:
            print("[OK] BR > 0 but no B22 found in dictionary, BM value = 0")
    else:
        # Second check: If BR = 0, check B9 for "GSB" or "Geogrid Reinforced GSB"
        for key, value in pavement_dict.items():
            if key.startswith('B9_'):
                parts = key.split('_')
                if len(parts) >= 2:
                    layer_name = parts[1]
                    if layer_name == "GSB" or layer_name == "Geogrid Reinforced GSB":
                        bm_value = value / 1000 if pd.notna(value) and value != 0 else 0
                        print(f"[OK] BR = 0, Found B9 '{layer_name}' in dictionary: {key} = {value}, BM value = {bm_value}")
                        break
        if bm_value == 0:
            print("[OK] BR = 0 and no GSB found in B9, BM value = 0")
    
    # Column BS = IF B10="RAP" THEN C10/1000 ELSE 0
    BS_COL_INDEX = 70
    bs_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B10_RAP_'):
            bs_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B10 RAP in dictionary: {key} = {value}, BS value = {bs_value}")
            break
    if bs_value == 0:
        print("[OK] No RAP found in B10, BS value = 0")
    
    # Column BT = IF B12="BM" THEN C12/1000 ELSE 0
    BT_COL_INDEX = 71
    bt_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B12_BM_'):
            bt_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B12 BM in dictionary: {key} = {value}, BT value = {bt_value}")
            break
    if bt_value == 0:
        print("[OK] No BM found in B12, BT value = 0")
    
    # Column BU = IF B13="DBM" THEN C13/1000 ELSE 0
    BU_COL_INDEX = 72
    bu_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B13_DBM_'):
            bu_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B13 DBM in dictionary: {key} = {value}, BU value = {bu_value}")
            break
    if bu_value == 0:
        print("[OK] No DBM found in B13, BU value = 0")
    
    # Column BV = IF B14="PC&SC" THEN C14/1000 ELSE 0
    BV_COL_INDEX = 73
    bv_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B14_PC&SC_'):
            bv_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B14 PC&SC in dictionary: {key} = {value}, BV value = {bv_value}")
            break
    if bv_value == 0:
        print("[OK] No PC&SC found in B14, BV value = 0")
    
    # Column BW = IF B15="SDBC" THEN C15/1000 ELSE 0
    BW_COL_INDEX = 74
    bw_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B15_SDBC_'):
            bw_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B15 SDBC in dictionary: {key} = {value}, BW value = {bw_value}")
            break
    if bw_value == 0:
        print("[OK] No SDBC found in B15, BW value = 0")
    
    # Column BX = IF B16="BC" THEN C16/1000 ELSE 0
    BX_COL_INDEX = 75
    bx_value = 0
    for key, value in pavement_dict.items():
        if key.startswith('B16_BC_'):
            bx_value = value / 1000 if pd.notna(value) and value != 0 else 0
            print(f"[OK] Found B16 BC in dictionary: {key} = {value}, BX value = {bx_value}")
            break
    if bx_value == 0:
        print("[OK] No BC found in B16, BX value = 0")
    
    # Set columns
    for col_idx, col_value, col_name in [
        (AY_COL_INDEX, ay_value, 'LHS_GSB_Thickness'),
        (BA_COL_INDEX, ba_value, 'LHS_WMM_Thickness'),
        (BB_COL_INDEX, bb_value, 'LHS_AIL_Thickness'),
        (BC_COL_INDEX, bc_value, 'LHS_DLC_Thickness'),
        (BD_COL_INDEX, bd_value, 'LHS_PQC_Thickness'),
        (BE_COL_INDEX, be_value, 'LHS_RAP_Thickness'),
        (BF_COL_INDEX, bf_value, 'LHS_BM_Thickness'),
        (BG_COL_INDEX, bg_value, 'LHS_DBM_Thickness'),
        (BH_COL_INDEX, bh_value, 'LHS_PC&SC_Thickness'),
        (BI_COL_INDEX, bi_value, 'LHS_SDBC_Thickness'),
        (BJ_COL_INDEX, bj_value, 'LHS_BC_Thickness'),
        (BL_COL_INDEX, bl_value, 'RHS_CTSB_Thickness'),
        (BM_COL_INDEX, bm_value, 'RHS_GSB_Thickness'),
        (BN_COL_INDEX, bn_value, 'RHS_CTB_Thickness'),
        (BO_COL_INDEX, bo_value, 'RHS_WMM_Thickness'),
        (BP_COL_INDEX, bp_value, 'RHS_AIL_Thickness'),
        (BQ_COL_INDEX, bq_value, 'RHS_DLC_Thickness'),
        (BR_COL_INDEX, br_value, 'RHS_PQC_Thickness'),
        (BS_COL_INDEX, bs_value, 'RHS_RAP_Thickness'),
        (BT_COL_INDEX, bt_value, 'RHS_BM_Thickness'),
        (BU_COL_INDEX, bu_value, 'RHS_DBM_Thickness'),
        (BV_COL_INDEX, bv_value, 'RHS_PC&SC_Thickness'),
        (BW_COL_INDEX, bw_value, 'RHS_SDBC_Thickness'),
        (BX_COL_INDEX, bx_value, 'RHS_BC_Thickness'),
    ]:
        if len(df.columns) <= col_idx:
            while len(df.columns) < col_idx:
                df[f'Empty_{len(df.columns)}'] = None
            df.insert(col_idx, f'Col_{col_idx}', col_value)
        else:
            df.iloc[:, col_idx] = col_value
        print(f"[OK] Column index {col_idx} ({col_name}) set to: {col_value}")
    
    # MODIFIED: Store data in collector instead of writing to Excel
    print(f"\n[OK] Storing data for later writing...")
    
    collector = get_collector()
    collector.add_data('pavement_input', {
        'sheet_name': 'Quantity',
        'dataframe': df,
        'start_row': 6,  # Row 7 (0-indexed)
        'start_col': 0,  # Column A
        'write_header': False,
        'write_index': False
    })
    
    print("[OK] Data stored successfully!")
    
    return df


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Main function to execute all steps"""
    print("\n" + "="*80)
    print("PAVEMENT INPUT PROCESSOR")
    print("="*80)
    print("Configuration:")
    print("  • Pavement_Input.xlsx: Read from Excel row 9 (Col B, C, E, F)")
    print("  • main_carriageway.xlsx: Populate columns")
    print("  • Output: Save to output folder")
    print("="*80 + "\n")
    
    try:
        # Step 1: Create pavement dictionary
        pavement_dict = create_pavement_dictionary(PAVEMENT_INPUT_FILE)
        
        # Step 2: Populate main_carriageway.xlsx
        df = populate_columns(MAIN_CARRIAGEWAY_FILE, pavement_dict, OUTPUT_EXCEL)
        
        # Success summary
        print("\n" + "="*80)
        print("SUCCESS! [OK]")
        print("="*80)
        print("Output file:", OUTPUT_EXCEL)
        print("Total rows:", len(df))
        print("Total columns:", len(df.columns))
        
        # Note: File will be uploaded to GCS at the end of all processing in main.py
        # No need to upload here for efficiency
        
        # Cleanup temp files - only remove temp files, not the local output file
        os.remove(PAVEMENT_INPUT_FILE)
        # Only remove OUTPUT_EXCEL if it's a temp file (not the local SESSION_OUTPUT_FILE)
        if OUTPUT_EXCEL != os.getenv('SESSION_OUTPUT_FILE', ''):
            try:
                os.remove(OUTPUT_EXCEL)
            except:
                pass
        
        # Show sample output
        print("\n" + "="*80)
        print("SAMPLE OUTPUT:")
        print("-"*80)
        print("First 3 rows, Column AX value:")
        for idx in range(min(3, len(df))):
            ax_val = df.iloc[idx, 49] if len(df.columns) > 49 else None
            print(f"  Row {idx + 2}: {ax_val}")
        
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