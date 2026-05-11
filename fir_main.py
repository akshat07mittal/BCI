
import os
import numpy as np
import pandas as pd
from scipy.signal import firwin, filtfilt


RAW_DIR = 'D:/BCI/MILimbEEG/data'
OUTPUT_ROOT = 'D:/BCI/MILimbEEG'
FIR_DIR = os.path.join(OUTPUT_ROOT, "fir_dataset")

FS = 125  # sampling frequency (Hz)

# Frequency bands (Hz)
BANDS = {
    "a": (8, 12),    # Alpha
    "b": (12, 30),   # Beta
    "g": (30, 50)    # Gamma
}

FILTER_ORDER = 101  # odd -> linear-phase FIR

# =========================
# FIR FILTER FUNCTION
# =========================

def bandpass_fir(signal, lowcut, highcut, fs, order):
    """
    Linear-phase FIR bandpass filter applied in zero-phase manner.
    """
    nyq = 0.5 * fs
    taps = firwin(
        numtaps=order,
        cutoff=[lowcut / nyq, highcut / nyq],
        pass_zero=False
    )
    return filtfilt(taps, [1.0], signal)

# =========================
# FULL DATASET PROCESSING
# =========================

def main():
    print(f"Scanning raw data in: {RAW_DIR}")
    if not os.path.exists(RAW_DIR):
        print(f"Error: Raw directory not found at {RAW_DIR}")
        return

    # Create top-level FIR directory
    os.makedirs(FIR_DIR, exist_ok=True)
    print(f"Saving filtered data to: {FIR_DIR}")

    # 1. List all subject folders
    try:
        subject_folders = [f for f in os.listdir(RAW_DIR) if os.path.isdir(os.path.join(RAW_DIR, f)) and f.startswith('S')]
    except Exception as e:
         print(f"Error reading directory {RAW_DIR}: {e}")
         return

    subject_folders.sort()  # Process in order S1, S2, ...

    total_files = 0
    processed_count = 0

    for subject in subject_folders:
        subject_path = os.path.join(RAW_DIR, subject)
        
        # Create output folder for this subject
        output_subject_path = os.path.join(FIR_DIR, subject)
        os.makedirs(output_subject_path, exist_ok=True)

        # 2. List all CSV files in subject folder
        csv_files = [f for f in os.listdir(subject_path) if f.endswith('.csv')]
        csv_files.sort()
        
        print(f"Processing {subject} ({len(csv_files)} files)...")

        for csv_file in csv_files:
            input_path = os.path.join(subject_path, csv_file)
            output_filename = f"f_{csv_file}"
            output_path = os.path.join(output_subject_path, output_filename)

            # Check if exists (optional)
            if os.path.exists(output_path):
                processed_count += 1
                continue

            try:
                # Load raw CSV
                # Assuming standard format: col 0 is index/time, col 1-16 are channels
                df = pd.read_csv(input_path)

                # Prepare output dataframe
                # Keep the first column (often sample index)
                output_df = pd.DataFrame({
                    df.columns[0]: df.iloc[:, 0].values
                })

                # Process each channel (columns 1 to end)
                for col in df.columns[1:]:
                    # Ensure numeric
                    signal = pd.to_numeric(df[col], errors='coerce').fillna(0).values

                    for band_label, (low, high) in BANDS.items():
                        filtered_signal = bandpass_fir(
                            signal,
                            lowcut=low,
                            highcut=high,
                            fs=FS,
                            order=FILTER_ORDER
                        )
                        output_df[f"{col}_{band_label}"] = filtered_signal

                # Save
                output_df.to_csv(output_path, index=False)
                processed_count += 1

            except Exception as e:
                print(f"  Error processing {csv_file}: {e}")
        
    print(f"✅ Processing complete. {processed_count} files saved to {FIR_DIR}.")

if __name__ == "__main__":
    main()
