"""
Collected Writer
Writes all collected data (CSV/JSON in session collected dir) into the session's output Excel file once.
"""

import os
import sys
import io
from dotenv import load_dotenv

# Add project root to Python path
import os as _os
from pathlib import Path
project_root = _os.path.join(_os.path.dirname(__file__), '..', '..')
sys.path.append(project_root)

from src.utils.excel_writer import write_all_collected_data

load_dotenv()
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def main():
    output_file = os.getenv('SESSION_OUTPUT_FILE')
    if not output_file or not os.path.exists(output_file):
        raise FileNotFoundError('SESSION_OUTPUT_FILE not set or does not exist')
    print("Collected Writer: writing all collected data to:", output_file)
    write_all_collected_data(output_file)
    print("Collected Writer: done (written and cleaned, unless KEEP_COLLECTED=true)")


if __name__ == "__main__":
    main()
