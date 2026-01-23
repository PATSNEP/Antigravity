from backend.ppt_processor import process_ppt
import os

import glob

# Auto-detect latest CSV
csv_files = glob.glob("*.csv") + glob.glob("uploads/*.csv")
if not csv_files:
    raise FileNotFoundError("No CSV files found in root or uploads/ folder.")

# Sort by modification time (newest first)
latest_csv = max(csv_files, key=os.path.getmtime)

print(f"--- Test Run ---")
print(f"Auto-selected Input File: {latest_csv}")

csv_file = latest_csv
output_dir = "backend/outputs"
os.makedirs(output_dir, exist_ok=True)

try:
    print("Running processor (Simulation)...")
    out_file = process_ppt(csv_file, output_dir)
    print(f"Success! Output: {out_file}")
    
    # Validation
    final_path = os.path.join(output_dir, out_file)
    if os.path.exists(final_path):
        print("File exists.")
        # Minimal size check
        if os.path.getsize(final_path) > 1000:
            print("File size looks valid.")
        else:
            print("Warning: File too small.")
            
except Exception as e:
    print(f"FAILED: {e}")
    import traceback
    traceback.print_exc()

