import json
import os
import sys
from pathlib import Path
from openpyxl import load_workbook
import shutil

# Add project root to Python path
project_root = os.path.join(os.path.dirname(__file__), '..', '..')
sys.path.append(project_root)

from src.utils.gcs_utils import get_gcs_handler

def log_debug(message):
    debug_file = Path(__file__).parent / 'boq_debug.log'
    with open(debug_file, 'a') as f:
        f.write(f"{message}\n")
    print(message)

# Use session directories from environment
session_id = os.getenv('SESSION_ID', 'default')
is_merged = os.getenv('IS_MERGED', 'True').lower() == 'true'

# Use SESSION_OUTPUT_FILE if available (local file), otherwise fallback to GCS
output_file = os.getenv('SESSION_OUTPUT_FILE', '')
if output_file and os.path.exists(output_file):
    boq_output_path = Path(output_file)
    log_debug(f"Using local output file: {output_file}")
else:
    # Fallback: download from GCS (for backward compatibility)
    gcs = get_gcs_handler()
    if is_merged:
        output_filename = f"{session_id}_main_carriageway_and_boq.xlsx"
    else:
        output_filename = f"{session_id}_main_carriageway.xlsx"
    output_gcs_path = gcs.get_gcs_path(session_id, output_filename, 'output')
    boq_output_path = Path(gcs.download_to_temp(output_gcs_path, suffix='.xlsx'))
    log_debug(f"[GCS] Downloaded output file from GCS: {boq_output_path}")

log_debug(f"=== BOQ POPULATOR - FORMULA WRITING TO TEMPLATE ===")

project_root = Path(__file__).parent.parent.parent
boq_formula_json_path = project_root / 'boq_formula_mapping.json'

try:
    # Load formula mapping from JSON
    if not boq_formula_json_path.exists():
        log_debug(f"ERROR: Formula mapping JSON not found at {boq_formula_json_path}")
        exit(1)
    
    with open(boq_formula_json_path, 'r', encoding='utf-8') as f:
        formula_mapping = json.load(f)
    
    log_debug(f"Loaded formula mapping with {len(formula_mapping)} items")

    # Check if the merged template file exists
    if not boq_output_path.exists():
        log_debug(f"ERROR: Merged template file not found at {boq_output_path}")
        exit(1)

    # Load the existing merged workbook
    wb = load_workbook(boq_output_path)
    sheet = wb['BOQ']
    
    # OPTIMIZATION: Batch-generate all formula assignments before writing
    log_debug(f"Batch-generating formulas for {len(formula_mapping)} items...")
    formula_batch = []
    
    for item_code, formula_data in formula_mapping.items():
        excel_row = formula_data['excel_row']
        formula_E = formula_data['column_E']
        formula_F = formula_data['column_F']
        
        if formula_E:
            formula_batch.append((f'E{excel_row}', formula_E))
        if formula_F:
            formula_batch.append((f'F{excel_row}', formula_F))
    
    # Batch-write all formulas
    log_debug(f"Writing {len(formula_batch)} formulas to BOQ sheet...")
    for cell_address, formula in formula_batch:
        sheet[cell_address] = formula
    
    populated_count = len(formula_mapping)

    # Save the workbook
    wb.save(boq_output_path)
    wb.close()
    
    log_debug(f"Written {populated_count} formulas to merged template")
    
    # Note: File will be uploaded to GCS at the end of all processing in main.py
    # No need to upload here for efficiency
    
    # Cleanup - only remove temp files, not the local output file
    # Only remove if it's a temp file (not the local SESSION_OUTPUT_FILE)
    if str(boq_output_path) != os.getenv('SESSION_OUTPUT_FILE', ''):
        try:
            os.remove(boq_output_path)
        except:
            pass

except Exception as e:
    log_debug(f"ERROR: {str(e)}")
    import traceback
    log_debug(f"TRACEBACK: {traceback.format_exc()}")

log_debug("=== FINISHED ===")
