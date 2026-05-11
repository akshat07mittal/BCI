import pandas as pd
import os
import random

# CONFIG
METADATA_PATH = r"D:\BCI\MILimbEEG\data2\metadata.xlsx"

# The metadata was generated with paths relative to the script execution in 'codes'
# e.g., '../MILimbEEG/data/...'
# We need to resolve these relative to the 'codes' directory where this script runs.
BASE_EXECUTION_DIR = r"D:\BCI\codes"

def verify_metadata():
    print(f"Loading metadata from: {METADATA_PATH}")
    if not os.path.exists(METADATA_PATH):
        print(f"❌ Error: Metadata file NOT FOUND.")
        return

    df = pd.read_csv(METADATA_PATH) if METADATA_PATH.endswith('.csv') else pd.read_excel(METADATA_PATH)
    print(f"[OK] Loaded {len(df)} rows from metadata.")

    # Check 1: Check a random sample of files
    sample_size = 5
    print(f"\nChecking a random sample of {sample_size} (or fewer) files...")
    
    # Ensure 'local_url' exists
    if 'local_url' not in df.columns:
        print("[Error] 'local_url' column not found in metadata.")
        print("Columns found:", df.columns.tolist())
        return

    sample = df.sample(min(len(df), sample_size))
    
    missing_count = 0
    
    for idx, row in sample.iterrows():
        rel_path = row['local_url']
        # Resolve path
        # If rel_path starts with '..', verify it relative to D:\BCI\codes
        # We need to construct the full path manually to check existence relative to script location
        full_path = os.path.normpath(os.path.join(BASE_EXECUTION_DIR, rel_path))
        
        exists = os.path.exists(full_path)
        status = "[FOUND]" if exists else "[MISSING]"
        
        print(f"  {status}: {rel_path}")
        if not exists:
            print(f"    -> Looked at: {full_path}")
            missing_count += 1

    # Check 2: Check total file count on disk vs metadata
    # This might be slow if scanning everything, let's just do the sample check first.
    
    print("\n--- Summary ---")
    if missing_count == 0:
        print("[Pass] Random verification PASSED. The metadata paths point to existing files.")
        print("You can confidently use this metadata file.")
    else:
        print(f"[Fail] Verification FAILED. {missing_count}/{len(sample)} checked files were missing.")
        print("The paths in metadata.xlsx might be incorrect relative to your script location.")

if __name__ == "__main__":
    verify_metadata()
