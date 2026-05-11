import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.signal import welch

# Import the filter function from fir_main
from fir_main import bandpass_fir, BANDS, FS, FILTER_ORDER

# =========================
# CONFIGURATION
# =========================
VALIDATION_DIR = r"D:\BCI\MILimbEEG\validation_plots"
REAL_DATA_PATH = r"D:\BCI\MILimbEEG\fir_dataset\S1\f_S1R1I1_1.csv"

os.makedirs(VALIDATION_DIR, exist_ok=True)

def plot_psd(axis, signal, fs, label, color):
    f, Pxx = welch(signal, fs=fs, nperseg=256)
    axis.semilogy(f, Pxx, label=label, color=color)
    axis.set_ylabel('PSD (V^2/Hz)')
    axis.grid(True, which='both', linestyle='--', alpha=0.5)

# =========================
# 1. SYNTHETIC DATA TEST
# =========================
def run_synthetic_test():
    print("Running Synthetic Data Test...")
    t = np.arange(0, 4.0, 1/FS)
    
    # 10Hz (Alpha), 20Hz (Beta), 40Hz (Gamma)
    s10 = np.sin(2 * np.pi * 10 * t)
    s20 = np.sin(2 * np.pi * 20 * t)
    s40 = np.sin(2 * np.pi * 40 * t)
    synthetic_signal = s10 + s20 + s40
    
    filtered = {}
    for band, (low, high) in BANDS.items():
        filtered[band] = bandpass_fir(synthetic_signal, low, high, FS, FILTER_ORDER)
    
    # Plot Time Domain
    fig, axes = plt.subplots(4, 1, figsize=(12, 10), sharex=True)
    axes[0].plot(t, synthetic_signal, color='black')
    axes[0].set_title("Synthetic Signal (10Hz + 20Hz + 40Hz)")
    
    axes[1].plot(t, filtered['a'], color='blue')
    axes[1].set_title("Alpha Band (Expected: 10Hz)")
    
    axes[2].plot(t, filtered['b'], color='green')
    axes[2].set_title("Beta Band (Expected: 20Hz)")
    
    axes[3].plot(t, filtered['g'], color='red')
    axes[3].set_title("Gamma Band (Expected: 40Hz)")
    axes[3].set_xlabel("Time (s)")
    
    plt.tight_layout()
    plt.savefig(os.path.join(VALIDATION_DIR, "synthetic_validation_time.png"))
    plt.close()

    # Plot Frequency Domain
    fig, ax = plt.subplots(figsize=(10, 6))
    plot_psd(ax, synthetic_signal, FS, "Original", "black")
    plot_psd(ax, filtered['a'], FS, "Alpha", "blue")
    plot_psd(ax, filtered['b'], FS, "Beta", "green")
    plot_psd(ax, filtered['g'], FS, "Gamma", "red")
    ax.set_title("Synthetic Signal PSD Comparison")
    ax.legend()
    ax.set_xlabel("Frequency (Hz)")
    ax.set_xlim(0, 60)
    
    plt.savefig(os.path.join(VALIDATION_DIR, "synthetic_validation_psd.png"))
    plt.close()
    print("Synthetic test plots saved.")

# =========================
# 2. REAL DATA SPECTRAL ANALYSIS
# =========================
def run_real_data_test():
    print(f"Running Real Data Test on: {os.path.basename(REAL_DATA_PATH)}")
    if not os.path.exists(REAL_DATA_PATH):
        print(f"File not found: {REAL_DATA_PATH}")
        return
    
    df = pd.read_csv(REAL_DATA_PATH)
    
    # Find a base channel name (columns are like '0_a', '1_a', ...)
    cols = [c for c in df.columns if c.endswith('_a')]
    if not cols:
        print("No band-split columns found in file.")
        return
    
    # Extract channel prefix (e.g., '0' from '0_a')
    first_col = cols[0]
    base_channel = first_col.split('_')[0] 
    print(f"Analyzing channel {base_channel} bands...")
    
    fig, ax = plt.subplots(figsize=(12, 6))
    
    colors = {'a': 'blue', 'b': 'green', 'g': 'red'}
    labels = {'a': 'Alpha (8-12Hz)', 'b': 'Beta (12-30Hz)', 'g': 'Gamma (30-50Hz)'}
    
    for band in BANDS.keys():
        col_name = f"{base_channel}_{band}"
        if col_name in df.columns:
            plot_psd(ax, df[col_name].values, FS, labels[band], colors[band])
    
    ax.set_title(f"Real Data PSD Validation ({base_channel})")
    ax.legend()
    ax.set_xlabel("Frequency (Hz)")
    ax.set_xlim(0, 60)
    
    # Add band boundaries
    for band, (low, high) in BANDS.items():
        ax.axvspan(low, high, color=colors[band], alpha=0.1)
    
    plt.savefig(os.path.join(VALIDATION_DIR, "real_data_psd_validation.png"))
    plt.close()
    print("Real data test plot saved.")

if __name__ == "__main__":
    run_synthetic_test()
    run_real_data_test()
    print(f"Validation complete. Plots are in: {VALIDATION_DIR}")
