import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# =========================
# PATHS
# =========================
PROJECT_ROOT = r"D:\BCI"
DATA_DIR = os.path.join(PROJECT_ROOT, "MILimbEEG", "data2")
PLOT_DIR = os.path.join(PROJECT_ROOT, "MILimbEEG", "plots")

os.makedirs(PLOT_DIR, exist_ok=True)

METADATA_PATH = os.path.join(DATA_DIR, "metadata.xlsx")

OVERLAY_IMG = os.path.join(PLOT_DIR, "electrode9_decomposition_overlay.png")
STACKED_IMG = os.path.join(PLOT_DIR, "electrode9_decomposition_stacked.png")

FS = 125  # sampling frequency (Hz)
ELECTRODE = 9

# =========================
# LOAD METADATA
# =========================
metadata = pd.read_excel(METADATA_PATH)

# choose one movement instance (change index if needed)
row_idx = 64

local_path = metadata.loc[row_idx, "local_url"]

# Filtered data is in fir_dataset/{patient}/f_{filename}
# Let's derive it or look for it.
patient = f"S{metadata.loc[row_idx, 'patient_number']}"
filename = os.path.basename(local_path)
filtered_path = os.path.join(PROJECT_ROOT, "MILimbEEG", "fir_dataset", patient, f"f_{filename}")

# =========================
# LOAD EEG FILES
# =========================
raw_df = pd.read_csv(local_path)
filt_df = pd.read_csv(filtered_path)

# Drop first row (electrode numbers)
raw_df = raw_df.iloc[1:, :]
filt_df = filt_df.iloc[1:, :]

# Drop serial-number column
raw_df = raw_df.iloc[:, 1:]
filt_df = filt_df.iloc[:, 1:]

raw_df = raw_df.astype(float)
filt_df = filt_df.astype(float)

num_samples = len(raw_df)

# =========================
# TIME AXIS
# =========================
time_sec = np.arange(num_samples) / FS

# =========================
# EXTRACT SIGNALS (ELECTRODE 9)
# =========================
original = raw_df.iloc[:, ELECTRODE].values

alpha = filt_df[f"{ELECTRODE}_a"].values
beta = filt_df[f"{ELECTRODE}_b"].values
gamma = filt_df[f"{ELECTRODE}_g"].values

# =========================
# OVERLAY PLOT
# =========================
plt.figure(figsize=(14, 6))

plt.plot(time_sec, original, label="Original EEG", linewidth=1.2)
plt.plot(time_sec, alpha, label="Alpha (8–12 Hz)", linewidth=1.0)
plt.plot(time_sec, beta, label="Beta (12–30 Hz)", linewidth=1.0)
plt.plot(time_sec, gamma, label="Gamma (30–50 Hz)", linewidth=1.0)

plt.title(
    "EEG Decomposition (Overlay)\n"
    f"Electrode {ELECTRODE}"
)
plt.xlabel("Time (seconds)")
plt.ylabel("Amplitude (µV)")
plt.legend()
plt.grid(True)

plt.tight_layout()
plt.savefig(OVERLAY_IMG, dpi=300)
plt.close()

# =========================
# STACKED PLOTS
# =========================
fig, axes = plt.subplots(4, 1, figsize=(14, 10), sharex=True)

axes[0].plot(time_sec, original)
axes[0].set_title("Original EEG")
axes[0].set_ylabel("µV")

axes[1].plot(time_sec, alpha)
axes[1].set_title("Alpha (8–12 Hz)")
axes[1].set_ylabel("µV")

axes[2].plot(time_sec, beta)
axes[2].set_title("Beta (12–30 Hz)")
axes[2].set_ylabel("µV")

axes[3].plot(time_sec, gamma)
axes[3].set_title("Gamma (30–50 Hz)")
axes[3].set_ylabel("µV")
axes[3].set_xlabel("Time (seconds)")

fig.suptitle(
    "EEG Decomposition (Stacked)\n"
    f"Electrode {ELECTRODE}",
    fontsize=12
)

plt.tight_layout(rect=[0, 0, 1, 0.96])
plt.savefig(STACKED_IMG, dpi=300)
plt.close()

print("Saved plots:")
print(OVERLAY_IMG)
print(STACKED_IMG)
