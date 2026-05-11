
import os
import requests
import zipfile
import shutil
import re
import pandas as pd
import openpyxl

# -----------------------------
# CONFIG
# -----------------------------
DATA_URL = 'https://prod-dcd-datasets-cache-zipfiles.s3.eu-west-1.amazonaws.com/x8psbz3f6x-2.zip'
ZIP_NAME = 'dataset.zip'

# Updated path as per user request
# Data is already downloaded in 'data' folder
DATA_SOURCE_DIR = '../MILimbEEG/data'
OUTPUT_DIR = '../MILimbEEG/data2'

# RAW_DIR points to where the data IS
RAW_DIR = DATA_SOURCE_DIR
# META_PATH points to where the user wants the result
META_PATH = os.path.join(OUTPUT_DIR, 'metadata.xlsx')

# -----------------------------
# LABEL MAPPING (unchanged)
# -----------------------------
label_map = {
    '1': 'BEO',
    '2': 'CLH',
    '3': 'CRH',
    '4': 'DLF',
    '5': 'PLF',
    '6': 'DRF',
    '7': 'PRF',
    '8': 'Rest',
}

label_encoding = {
    'BEO': 0,
    'CLH': 1,
    'CRH': 2,
    'DLF': 3,
    'PLF': 4,
    'DRF': 5,
    'PRF': 6,
    'Rest': 7,
}

# -----------------------------
# STEP 1: DOWNLOAD ZIP
# -----------------------------
# -----------------------------
# STEP 1: DOWNLOAD ZIP
# -----------------------------
# -----------------------------
# STEP 1: PREPARE OUTPUT
# -----------------------------
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -----------------------------
# STEP 2: VERIFY DATA (SKIP DOWNLOAD/EXTRACT)
# -----------------------------
if os.path.isdir(RAW_DIR) and any(d.startswith('S') for d in os.listdir(RAW_DIR)):
    print(f"Found existing data in {RAW_DIR}, proceeding with metadata generation.")
else:
    print(f"ERROR: Expected data in {RAW_DIR} but found none.")
    print("Please ensure the dataset is extracted in that folder (folders S1, S2, ...).")
    exit(1)

# -----------------------------
# STEP 4: METADATA GENERATION
# -----------------------------
rows = []
srno = 1

def patient_sort_key(folder):
    match = re.match(r'S(\d+)', folder)
    return int(match.group(1)) if match else float('inf')

for patient_folder in sorted(os.listdir(RAW_DIR), key=patient_sort_key):
    if not patient_folder.startswith('S'):
        continue

    match_patient = re.match(r'S(\d+)', patient_folder)
    if not match_patient:
        continue

    patient_num = int(match_patient.group(1))
    patient_path = os.path.join(RAW_DIR, patient_folder)

    for file in sorted(os.listdir(patient_path)):
        if not file.endswith('.csv'):
            continue

        base = file[:-4]  # remove .csv
        # Example: S24R1I6_5
        match = re.match(r'S(\d+)R(\d+)([MI])(\d+)_(\d+)', base)
        if not match:
            continue

        subj_id, rep_num, task_type, label_num, task_rep_num = match.groups()
        local_url = os.path.join(patient_path, file)

        task_type_val = 1 if task_type == 'M' else 0
        label_str = label_map.get(label_num, 'Unknown')
        label_encoded = label_encoding.get(label_str, -1)

        rows.append([
            srno,
            patient_num,
            local_url,
            int(rep_num),
            task_type_val,
            int(task_rep_num),
            label_str,
            label_encoded
        ])
        srno += 1

# -----------------------------
# STEP 5: SAVE EXCEL
# -----------------------------
columns = [
    'srno',
    'patient_number',
    'local_url',
    'repetition_number',
    'task_type',
    'task_repetition_number',
    'task_label',
    'task_label_encoded'
]

df = pd.DataFrame(rows, columns=columns)
df.to_excel(META_PATH, index=False)

print(f'Metadata written to {META_PATH}')
