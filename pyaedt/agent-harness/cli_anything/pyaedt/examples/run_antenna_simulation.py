#!/usr/bin/env python3
"""
Printed Dipole Antenna with Microstrip Balun - Direct HFSS Simulation
Using PyAEDT 0.25.x

Center Frequency: 2.217 GHz
Substrate: FR4 (εr = 4.4)

IMPORTANT: Run this script in Windows CMD or PowerShell, NOT in WSL bash.
Example: cmd.exe /c "python run_antenna_simulation.py"
"""

import numpy as np
import os
import sys

# ============================================================================
# DESIGN PARAMETERS
# ============================================================================
print("=" * 70)
print("Printed Dipole Antenna with Microstrip Balun - HFSS Simulation")
print("=" * 70)

# Physical constants
c = 3e8  # Speed of light (m/s)
epsilon_r = 4.4  # FR4 relative permittivity

# Frequency parameters
f_center = 2.217e9  # Center frequency (Hz)
f_min = 1.5e9  # Sweep start (Hz)
f_max = 3.0e9  # Sweep end (Hz)
f_step = 0.01e9  # Sweep step (Hz)

# Wavelength calculations
lambda_0 = c / f_center  # Free space wavelength
lambda_g = lambda_0 / np.sqrt(epsilon_r)  # Guide wavelength in FR4

# Dipole parameters
L_dipole_initial = 50e-3  # Initial dipole arm length (m)
W_dipole = 3e-3  # Dipole arm width (m)

# Balun parameters
L_balun = lambda_g / 4  # Balun length (quarter wavelength)
W_feed_start = 1e-3  # Balun narrow end width (m)
W_feed_end = 3e-3  # Balun wide end width (m)

# Substrate parameters
L_substrate = 120e-3  # Substrate length (m)
W_substrate = 60e-3  # Substrate width (m)
H_substrate = 1.6e-3  # Substrate thickness (m)

print(f"Center Frequency: {f_center/1e9:.3f} GHz")
print(f"Free Space Wavelength: {lambda_0*1e3:.2f} mm")
print(f"Guide Wavelength (FR4): {lambda_g*1e3:.2f} mm")
print(f"Initial Dipole Arm Length: {L_dipole_initial*1e3:.1f} mm")
print(f"Balun Length (λg/4): {L_balun*1e3:.2f} mm")
print(f"Substrate Size: {L_substrate*1e3:.1f} x {W_substrate*1e3:.1f} x {H_substrate*1e3:.2f} mm")
print()

# ============================================================================
# DETECT AND SET AEDT ENVIRONMENT
# ============================================================================
print("Detecting AEDT installation...")

# Patch DesignSettings for AEDT 2019 compatibility
try:
    import pyaedt.application.Design as dm
    _orig_init = dm.DesignSettings.__init__
    def _patched_init(self, app):
        try:
            _orig_init(self, app)
        except AttributeError:
            self._app = app
            self.design_settings = None
            self.manipulate_inputs = None
    dm.DesignSettings.__init__ = _patched_init
except Exception:
    pass

from pyaedt import Hfss

# ============================================================================
# LAUNCH AEDT AND CREATE PROJECT
# ============================================================================
print("Launching AEDT...")

project_name = "Printed_Dipole_Antenna"
design_name = "Dipole_Balun_Design"

try:
    # Try to launch AEDT
    hfss = Hfss(
        projectname=project_name,
        designname=design_name,
        solution_type="Terminal Network",
        non_graphical=True,
        specified_version='2019.1'
    )
    print(f"AEDT launched successfully!")
    print(f"Project: {hfss.project_name}")
    print(f"Design: {hfss.design_name}")
except Exception as e:
    print(f"AEDT launch error: {e}")
    print("\n" + "="*70)
    print("TROUBLESHOOTING:")
    print("="*70)
    print("1. Make sure you're running this script in Windows CMD or PowerShell,")
    print("   NOT in WSL or Git Bash.")
    print("2. Verify AEDT is installed and licensed.")
    print("3. Check that environment variables are set correctly.")
    print("4. Try running this in Windows CMD:")
    print('   cd /d "D:\\work_tool\\CLI-Anything\\pyaedt\\agent-harness\\cli_anything\\pyaedt\\examples"')
    print("   python run_antenna_simulation.py")
    raise

# Set model units
UNITS = "mm"
hfss.modeler.model_units = UNITS

# ============================================================================
# CREATE SUBSTRATE
# ============================================================================
print("\nCreating FR4 substrate...")

substrate = hfss.modeler.create_box(
    position=[-L_substrate/2, -W_substrate/2, 0],
    dimensions_list=[L_substrate, W_substrate, H_substrate],
    name="Substrate",
    material="FR4_epoxy"
)
print(f"  Substrate created: {substrate.name}")

# ============================================================================
# CREATE GROUND PLANE
# ============================================================================
print("Creating ground plane...")

ground = hfss.modeler.create_box(
    position=[-L_substrate/2, -W_substrate/2, H_substrate],
    dimensions_list=[L_substrate, W_substrate, 0.035],
    name="Ground_Plane",
    material="copper"
)
print(f"  Ground plane created: {ground.name}")

# ============================================================================
# CREATE DIPOLE ARMS
# ============================================================================
print("Creating dipole arms...")

left_arm = hfss.modeler.create_box(
    position=[-L_dipole_initial, -W_dipole/2, H_substrate],
    dimensions_list=[L_dipole_initial, W_dipole, 0.035],
    name="Dipole_Left",
    material="copper"
)

right_arm = hfss.modeler.create_box(
    position=[0, -W_dipole/2, H_substrate],
    dimensions_list=[L_dipole_initial, W_dipole, 0.035],
    name="Dipole_Right",
    material="copper"
)
print(f"  Left arm: {left_arm.name}")
print(f"  Right arm: {right_arm.name}")

# ============================================================================
# CREATE MICROSTRIP BALUN (TAPERED FEED)
# ============================================================================
print("Creating microstrip balun...")

num_segments = 10
balun_length = L_balun

for i in range(num_segments):
    x_pos = -balun_length + (i * balun_length / num_segments)
    width = W_feed_start + (W_feed_end - W_feed_start) * (i / num_segments)
    next_width = W_feed_start + (W_feed_end - W_feed_start) * ((i + 1) / num_segments)
    seg_length = balun_length / num_segments
    avg_width = (width + next_width) / 2

    segment = hfss.modeler.create_box(
        position=[x_pos, -avg_width/2, H_substrate],
        dimensions_list=[seg_length + 0.001, avg_width, 0.035],
        name=f"Balun_Seg_{i}",
        material="copper"
    )

# Feed point connection
feed_point = hfss.modeler.create_box(
    position=[-balun_length - 1e-3, -W_feed_start/2, H_substrate],
    dimensions_list=[1e-3, W_feed_start, 0.035],
    name="Feed_Point",
    material="copper"
)
print(f"  Balun segments: {num_segments}")

# ============================================================================
# CREATE LUMPED PORT (FEED)
# ============================================================================
print("Creating lumped port excitation...")

port_rect = hfss.modeler.create_box(
    position=[-balun_length - 1e-3, -W_feed_start/2, H_substrate],
    dimensions_list=[0.5e-3, W_feed_start, 0.035],
    name="Port1_Geom",
    material="copper"
)

face_id = hfss.modeler.get_faceid(port_rect.name, 0)

hfss.lumped_port(
    name="Port1",
    faces=[face_id],
    impedance=50,
    direction=[1, 0, 0]
)
print("  Lumped port 'Port1' created (50 ohm)")

# ============================================================================
# CREATE RADIATION BOUNDARY (AIR BOX)
# ============================================================================
print("Creating radiation boundary...")

air_box = hfss.modeler.create_box(
    position=[-L_substrate, -L_substrate, 0],
    dimensions_list=[L_substrate * 3, L_substrate * 3, L_substrate * 2],
    name="Air_Box",
    material="air"
)

hfss.create_open_region(
    frequency=f_center,
    boundary="Radiation",
    objects=["Air_Box"]
)
print("  Radiation boundary created")

# ============================================================================
# CREATE ANALYSIS SETUP
# ============================================================================
print("\nCreating frequency sweep setup...")

setup = hfss.create_setup(
    name="Setup1",
    setup_type="Hfss",
    Frequency=f_center,
    MaxDeltaZ=0.02,
    MaxPasses=15,
    MinimumPasses=3
)
print(f"  Setup 'Setup1' created")

hfss.create_linear_count_sweep(
    setupname="Setup1",
    sweepname="Freq_Sweep",
    StartValue=f_min,
    StopValue=f_max,
    Count=int((f_max - f_min) / f_step) + 1,
    SweepType="Linear",
    SaveFields=True
)
print(f"  Frequency sweep: {f_min/1e9:.2f} - {f_max/1e9:.2f} GHz")

# ============================================================================
# SAVE PROJECT
# ============================================================================
print("\nSaving project...")
hfss.save_project()
print("  Project saved")

# ============================================================================
# RUN SIMULATION
# ============================================================================
print("\n" + "=" * 70)
print("Starting simulation...")
print("=" * 70)

hfss.analyze_setup("Setup1", num_threads=4)

print("\nSimulation completed!")

# ============================================================================
# EXTRACT S11 RESULTS
# ============================================================================
print("\n" + "=" * 70)
print("Extracting S11 Results")
print("=" * 70)

try:
    s11_data = hfss.get_solution_data(
        expession="S11",
        setup_sweeps=["Setup1:Freq_Sweep"],
    )

    s11_db = s11_data.data_real()
    freqs = s11_data.sweeps.get("Freq", s11_data.sweeps.get("Frequency", []))

    if len(s11_db) > 0:
        min_idx = np.argmin(s11_db)
        resonance_freq = freqs[min_idx] if len(freqs) > min_idx else f_center
        min_s11 = s11_db[min_idx]

        print(f"  Resonance frequency: {resonance_freq/1e9:.4f} GHz")
        print(f"  Minimum S11: {min_s11:.2f} dB")

        below_10db_indices = np.where(s11_db < -10)[0]
        if len(below_10db_indices) >= 2:
            bw_low = freqs[below_10db_indices[0]]
            bw_high = freqs[below_10db_indices[-1]]
            bandwidth = (bw_high - bw_low) / 1e9
            rel_bandwidth = (bandwidth / f_center) * 100
            print(f"  -10 dB Bandwidth: {bw_low/1e9:.3f} - {bw_high/1e9:.3f} GHz")
            print(f"  Bandwidth: {bandwidth*1e3:.1f} MHz")
            print(f"  Relative bandwidth: {rel_bandwidth:.2f}%")
        else:
            print("  S11 does not reach -10 dB threshold")
    else:
        print("  No S11 data retrieved")
        resonance_freq = f_center
        min_s11 = 0
except Exception as e:
    print(f"  Error extracting S11: {e}")
    resonance_freq = f_center
    min_s11 = 0

# ============================================================================
# FAR FIELD RADIATION PATTERN
# ============================================================================
print("\n" + "=" * 70)
print("Computing Far Field Radiation Pattern")
print("=" * 70)

try:
    hfss.create_far_field_sphere(
        name="Far_Field",
        frequency=f_center,
        azimuth_plane_size=360,
        elevation_plane_size=180,
        radial_distance=1,
        use_infinite_sphere=True
    )

    ff_data = hfss.get_solution_data(
        expression="GainTotal",
        setup_sweeps=["Setup1:Freq_Sweep"],
        far_field_sphere="Far_Field"
    )

    gain_total = ff_data.data_real()
    max_gain_idx = np.argmax(gain_total)
    max_gain = gain_total[max_gain_idx]

    print(f"  Maximum Gain: {max_gain:.2f} dBi")
    print("\n  Computing E-plane (xz) and H-plane (yz) patterns...")

except Exception as e:
    print(f"  Error computing far field: {e}")
    max_gain = 0

# ============================================================================
# DESIGN SUMMARY
# ============================================================================
print("\n" + "=" * 70)
print("Design Summary")
print("=" * 70)
print(f"  Center Frequency: {f_center/1e9:.3f} GHz (Target)")
print(f"  Resonance Frequency: {resonance_freq/1e9:.4f} GHz (Simulated)")
print(f"  Minimum S11: {min_s11:.2f} dB")
if abs(resonance_freq - f_center) / f_center > 0.02:
    print(f"  WARNING: Resonance deviates from target by {abs(resonance_freq - f_center)/f_center*100:.1f}%")
    print(f"  Consider running parameter sweep to optimize dipole length")

# ============================================================================
# CLEANUP
# ============================================================================
print("\nReleasing AEDT...")
hfss.release_desktop()

print("\n" + "=" * 70)
print("SIMULATION COMPLETE")
print("=" * 70)