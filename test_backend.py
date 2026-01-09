from backend.ppt_processor import process_ppt
import os

# Updated filename
csv_file = "UseCases-Table-2026-01-09T19_54_13.7355551Z.csv"
output_dir = "backend/outputs"
os.makedirs(output_dir, exist_ok=True)

try:
    print("Running processor (Simulation)...")
    out_file = process_ppt(csv_file, output_dir)
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
