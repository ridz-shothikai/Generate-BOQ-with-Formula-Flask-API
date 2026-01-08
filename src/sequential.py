"""
Master Script - Run All Processing Scripts
Executes all data processing scripts in the correct order
"""

import subprocess
import sys
import os
import io
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Get script directory
script_dir = os.path.dirname(os.path.abspath(__file__))

# Define scripts in execution order
SCRIPTS = [
    ('processor/tcs_schedule.py', 'TCS Schedule Processing'),
    ('processor/tcs_input.py', 'TCS Input Processing'),
    ('processor/emb_height.py', 'Embankment Height Processing'),
    ('processor/pavement_input.py', 'Pavement Input Processing'),
    ('processor/constant_fill.py', 'Constant Fill Processing'),
    ('internal/formula_applier.py', 'Formula Applier Processing'),
    # ('internal/recalc.py', 'Excel Formula Recalculation'),
    ('internal/pavement_input_with_internal.py', 'Pavement Input with Internal Processing'),
    ('internal/final_sum_applier.py', 'Final Sum Applier Processing'),
    ('internal/collected_writer.py', 'Write Collected Data to Excel'),
    # ('processor/calculator.py', 'Formula Calculation'),
    ('processor/boq_populator.py', 'BOQ Template Population'),
]

def run_script(script_name, description):
    """Run a Python script and return success status"""
    script_path = os.path.join(script_dir, script_name)
    
    print("\n" + "="*80)
    print(f"RUNNING: {description}")
    print(f"Script: {script_name}")
    print("="*80)
    
    try:
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            check=True
        )
        
        print(result.stdout)
        
        if result.stderr:
            print("STDERR:", result.stderr)
        
        print(f"\n✓ {script_name} completed successfully")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"\n✗ ERROR in {script_name}")
        print("STDOUT:", e.stdout)
        print("STDERR:", e.stderr)
        return False
    except FileNotFoundError:
        print(f"\n✗ ERROR: Script not found: {script_path}")
        return False
    except Exception as e:
        print(f"\n✗ ERROR: {e}")
        return False

def main():
    """Main execution function"""
    print("\n" + "="*80)
    print("MASTER SCRIPT - RUNNING ALL PROCESSING SCRIPTS")
    print("="*80)
    print(f"Total scripts to run: {len(SCRIPTS)}")
    print(f"Working directory: {script_dir}")
    print("="*80)
    
    results = []
    
    for script_name, description in SCRIPTS:
        success = run_script(script_name, description)
        results.append((script_name, success))
        
        if not success:
            print("\n" + "="*80)
            print(f"STOPPING: {script_name} failed")
            print("="*80)
            break
    
    # Final summary
    print("\n" + "="*80)
    print("EXECUTION SUMMARY")
    print("="*80)
    
    all_success = True
    for script_name, success in results:
        status = "✓ SUCCESS" if success else "✗ FAILED"
        print(f"{status}: {script_name}")
        if not success:
            all_success = False
    
    print("="*80)
    
    if all_success and len(results) == len(SCRIPTS):
        print("\n🎉 ALL SCRIPTS COMPLETED SUCCESSFULLY! 🎉\n")
        return 0
    else:
        print("\n⚠️  PROCESSING INCOMPLETE OR FAILED ⚠️\n")
        return 1

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)