
"""
Post-Processing Analysis for Printed Dipole Antenna
Extracts and plots S11, radiation patterns, and calculates bandwidth
"""

import numpy as np
import matplotlib.pyplot as plt

def analyze_s11(frequencies, s11_db):
    """Analyze S11 data and find key parameters."""
    # Find resonance
    min_idx = np.argmin(s11_db)
    resonance_freq = frequencies[min_idx]
    min_s11 = s11_db[min_idx]

    # Find -10 dB bandwidth
    below_10db = np.where(s11_db < -10)[0]
    if len(below_10db) >= 2:
        bw_indices = below_10db[[0, -1]]
        bw_low = frequencies[bw_indices[0]]
        bw_high = frequencies[bw_indices[1]]
        bandwidth = bw_high - bw_low
        rel_bandwidth = (bandwidth / resonance_freq) * 100
    else:
        bw_low = bw_high = resonance_freq
        bandwidth = 0
        rel_bandwidth = 0

    return {
        "resonance_freq": resonance_freq,
        "min_s11": min_s11,
        "bw_low": bw_low,
        "bw_high": bw_high,
        "bandwidth": bandwidth,
        "relative_bandwidth": rel_bandwidth
    }

def plot_radiation_pattern(theta, gain_e, gain_h, title):
    """Plot E-plane and H-plane radiation patterns."""
    fig, ax = plt.subplots(subplot_kw='polar')
    ax.plot(np.radians(theta), gain_e, 'b-', label='E-plane (xz)', linewidth=2)
    ax.plot(np.radians(theta), gain_h, 'r--', label='H-plane (yz)', linewidth=2)
    ax.set_title(title, pad=20)
    ax.legend(loc='upper right')
    ax.set_theta_zero_location('N')
    ax.set_theta_direction(-1)
    plt.tight_layout()
    plt.show()

def plot_s11(frequencies, s11_db, resonance_freq, bandwidth_low, bandwidth_high):
    """Plot S11 reflection coefficient."""
    plt.figure(figsize=(10, 6))
    plt.plot(frequencies/1e9, s11_db, 'b-', linewidth=2)
    plt.axhline(y=-10, color='r', linestyle='--', label='-10 dB threshold')
    plt.axvline(x=resonance_freq/1e9, color='g', linestyle=':', label=f'Resonance: {resonance_freq/1e9:.3f} GHz')

    if bandwidth_low != bandwidth_high:
        plt.axvspan(bandwidth_low/1e9, bandwidth_high/1e9, alpha=0.2, color='green', label='-10 dB Bandwidth')

    plt.xlabel('Frequency (GHz)')
    plt.ylabel('S11 (dB)')
    plt.title('S11 Reflection Coefficient')
    plt.grid(True, alpha=0.3)
    plt.legend()
    plt.tight_layout()
    plt.show()

# Example usage with simulated data
if __name__ == "__main__":
    # Simulated S11 data (replace with actual AEDT data)
    freqs = np.linspace(1.5e9, 3.0e9, 151)
    s11_simulated = -10 * np.exp(-0.5 * ((freqs - 2.217e9) / 0.3e9)**2) - 2

    results = analyze_s11(freqs, s11_simulated)

    print("=" * 50)
    print("S11 Analysis Results")
    print("=" * 50)
    print(f"Resonance Frequency: {results['resonance_freq']/1e9:.4f} GHz")
    print(f"Minimum S11: {results['min_s11']:.2f} dB")
    print(f"-10 dB Bandwidth: {results['bw_low']/1e9:.3f} - {results['bw_high']/1e9:.3f} GHz")
    print(f"Bandwidth: {results['bandwidth']/1e6:.1f} MHz")
    print(f"Relative Bandwidth: {results['relative_bandwidth']:.2f}%")
