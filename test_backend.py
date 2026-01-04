from backend.ppt_processor import process_ppt
import os

# Test Setup
CSV_PATH = "UseCases-Table-2026-01-04T09_57_57.3469416Z.csv"
OUTPUT_DIR = "backend/outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

try:
    print("Running processor (Simulation)...")
    out_file = process_ppt(CSV_PATH, OUTPUT_DIR)
    print(f"Success! Output: {out_file}")
    
    # Validation
    final_path = os.path.join(OUTPUT_DIR, out_file)
    if os.path.exists(final_path):
        print("File exists.")
        # Minimal size check
        if os.path.getsize(final_path) > 1000:
            print("File size looks valid.")
        else:
            print("Warning: File too small.")
            
except Exception as e:
    print(f"FAILED: {e}")
