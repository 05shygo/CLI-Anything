#!/usr/bin/env python3
"""
印刷偶极子天线（带微带巴伦馈线）- HFSS 仿真设计
Printed Dipole Antenna with Microstrip Balun - HFSS Simulation

AEDT 2019.1 + PyAEDT 0.8.11 兼容版

设计参数:
  - 天线类型：印刷偶极子 + 三角形微带巴伦
  - 中心频率：2.217 GHz
  - 介质板：FR4 (εr=4.4, 厚度1.6mm)
  - 偶极子臂长初始值：50 mm
  - 频率扫描：1.5 ~ 3.0 GHz，步进 0.01 GHz
"""

import numpy as np
import os
import sys
import time
import csv

# ============================================================================
# 0. 环境和兼容性配置
# ============================================================================
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'

# 日志输出到文件，避免 gbk 编码问题
PROJECT_DIR = r"D:\class_design"
os.makedirs(PROJECT_DIR, exist_ok=True)
LOG_FILE = os.path.join(PROJECT_DIR, "dipole_sim.log")

class DualLogger:
    def __init__(self, logpath):
        self.terminal = sys.stdout
        self.log = open(logpath, "w", encoding="utf-8")
    def write(self, msg):
        try:
            self.terminal.write(msg)
        except UnicodeEncodeError:
            self.terminal.write(msg.encode('ascii', 'replace').decode())
        self.log.write(msg)
        self.log.flush()
    def flush(self):
        self.terminal.flush()
        self.log.flush()

sys.stdout = DualLogger(LOG_FILE)
sys.stderr = sys.stdout

# Patch PyAEDT for AEDT 2019.1 compatibility
try:
    from pyaedt import desktop as _desktop_mod
    _origDesktop_init = _desktop_mod.Desktop.__init__
    def _patched_Desktop_init(self, *a, **kw):
        _origDesktop_init(self, *a, **kw)
        if not hasattr(self, 'student_version'):
            self.student_version = False
    _desktop_mod.Desktop.__init__ = _patched_Desktop_init
except Exception:
    pass

try:
    import pyaedt.application.Design as _design_mod
    _origDS = _design_mod.DesignSettings.__init__
    def _patchedDS(self, app):
        try:
            _origDS(self, app)
        except AttributeError:
            self._app = app
            self.design_settings = None
            self.manipulate_inputs = None
    _design_mod.DesignSettings.__init__ = _patchedDS
except Exception:
    pass

from pyaedt import Hfss

# ============================================================================
# 1. 设计参数计算
# ============================================================================
print("=" * 70)
print("STEP 1: Design Parameter Calculation")
print("=" * 70)

c = 3e8          # m/s
f0 = 2.217e9     # Hz
eps_r = 4.4

lambda_0 = c / f0                      # ~135.3 mm
lambda_g = lambda_0 / np.sqrt(eps_r)   # ~64.5  mm

# Substrate (mm)
sub_L = 120.0
sub_W = 60.0
sub_H = 1.6
cu_t  = 0.035

# Dipole (mm)
dipole_arm = 50.0   # initial arm length (one side)
dipole_w   = 3.0

# Balun (mm) - λg/4
balun_L = lambda_g * 1e3 / 4.0         # ~16.1 mm
balun_w_start = 1.0                    # narrow end
balun_w_end   = 3.0                    # wide end (matches dipole width)
balun_segs = 10

# Frequency sweep
f_min  = 1.5e9
f_max  = 3.0e9
f_step = 0.01e9

# Gap between dipole arms
gap = 1.0  # mm

print(f"  f0           = {f0/1e9:.3f} GHz")
print(f"  lambda_0     = {lambda_0*1e3:.2f} mm")
print(f"  lambda_g     = {lambda_g*1e3:.2f} mm")
print(f"  dipole arm   = {dipole_arm:.1f} mm")
print(f"  balun L      = {balun_L:.2f} mm (lambda_g/4)")
print(f"  substrate    = {sub_L} x {sub_W} x {sub_H} mm")
print(f"  sweep        = {f_min/1e9:.1f} ~ {f_max/1e9:.1f} GHz")
print()

# ============================================================================
# 2. Launch AEDT
# ============================================================================
print("=" * 70)
print("STEP 2: Launch AEDT and Create Project")
print("=" * 70)

PROJECT_NAME = "Printed_Dipole_v6"

hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, PROJECT_NAME),
    designname="Dipole_Balun",
    solution_type="DrivenModal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)

print(f"  Project: {hfss.project_name}")
print(f"  Design:  {hfss.design_name}")

hfss.modeler.model_units = "mm"

# Get COM handles
oEditor = hfss.odesign.SetActiveEditor("3D Modeler")
oDesign = hfss.odesign

# ============================================================================
# COM helper functions (verified format for AEDT 2019.1)
# ============================================================================
def create_box(name, x, y, z, dx, dy, dz, mat, solve_inside=None):
    if solve_inside is None:
        solve_inside = mat.lower() not in ("copper", "pec", "aluminum")
    oEditor.CreateBox(
        ["NAME:BoxParameters",
         "XPosition:=", str(x),
         "YPosition:=", str(y),
         "ZPosition:=", str(z),
         "XSize:=", str(dx),
         "YSize:=", str(dy),
         "ZSize:=", str(dz)],
        ["NAME:Attributes",
         "Name:=", name,
         "Flags:=", "",
         "Color:=", "(143 175 131)",
         "Transparency:=", 0,
         "PartCoordinateSystem:=", "Global",
         "UDMId:=", "",
         "MaterialValue:=", '"' + mat + '"',
         "SurfaceMaterialValue:=", '""',
         "SolveInside:=", solve_inside,
         "IsMaterialEditable:=", True,
         "UseMaterialAppearance:=", False,
         "IsLightweight:=", False])

def create_rectangle(name, x, y, z, w, h, axis="Y"):
    """Create a 2D rectangle sheet (for port face)."""
    oEditor.CreateRectangle(
        ["NAME:RectangleParameters",
         "IsCovered:=", True,
         "XStart:=", str(x),
         "YStart:=", str(y),
         "ZStart:=", str(z),
         "Width:=", str(w),
         "Height:=", str(h),
         "WhichAxis:=", axis],
        ["NAME:Attributes",
         "Name:=", name,
         "Flags:=", "",
         "Color:=", "(0 0 255)",
         "Transparency:=", 0,
         "PartCoordinateSystem:=", "Global",
         "UDMId:=", "",
         "MaterialValue:=", '""',
         "SurfaceMaterialValue:=", '""',
         "SolveInside:=", True,
         "IsMaterialEditable:=", True,
         "UseMaterialAppearance:=", False,
         "IsLightweight:=", False])

# ============================================================================
# 3. Create Geometry
# ============================================================================
print()
print("=" * 70)
print("STEP 3: Create Antenna Geometry")
print("=" * 70)

z_top = sub_H  # top surface of substrate

# 3.1 FR4 Substrate
print("  [3.1] FR4 Substrate ...")
create_box("Substrate",
           -sub_L/2, -sub_W/2, 0,
           sub_L, sub_W, sub_H,
           "FR4_epoxy")
print(f"        {sub_L} x {sub_W} x {sub_H} mm")

# 3.2 Ground plane (bottom)
print("  [3.2] Ground plane (bottom) ...")
create_box("Ground",
           -sub_L/2, -sub_W/2, -cu_t,
           sub_L, sub_W, cu_t,
           "copper")
print(f"        {sub_L} x {sub_W} x {cu_t} mm (z<0)")

# 3.3 Dipole arms (top surface)
print("  [3.3] Dipole arms ...")
left_x = -(gap/2 + dipole_arm)
create_box("Dipole_Left",
           left_x, -dipole_w/2, z_top,
           dipole_arm, dipole_w, cu_t,
           "copper")
print(f"        Left:  x=[{left_x:.1f}, {-gap/2:.1f}]")

right_x = gap / 2
create_box("Dipole_Right",
           right_x, -dipole_w/2, z_top,
           dipole_arm, dipole_w, cu_t,
           "copper")
print(f"        Right: x=[{right_x:.1f}, {right_x+dipole_arm:.1f}]")

# 3.4 Microstrip balun (tapered feed on top surface)
print("  [3.4] Microstrip balun (tapered) ...")
for i in range(balun_segs):
    frac_s = i / balun_segs
    frac_e = (i + 1) / balun_segs
    w_s = balun_w_end - (balun_w_end - balun_w_start) * frac_s
    w_e = balun_w_end - (balun_w_end - balun_w_start) * frac_e
    avg_w = (w_s + w_e) / 2.0
    seg_len = balun_L / balun_segs
    y_pos = -dipole_w / 2 - seg_len * (i + 1)
    create_box(f"Balun_{i}",
               -avg_w / 2, y_pos, z_top,
               avg_w, seg_len, cu_t,
               "copper")
print(f"        {balun_segs} segments, L={balun_L:.2f}mm, taper {balun_w_end}->{balun_w_start}mm")

# 3.5 Feed extension line
feed_ext = 5.0  # mm
feed_y_end = -dipole_w / 2 - balun_L - feed_ext
create_box("Feed_Line",
           -balun_w_start / 2, feed_y_end, z_top,
           balun_w_start, feed_ext, cu_t,
           "copper")
print(f"        Feed line ext: {feed_ext}mm, width={balun_w_start}mm")

# ============================================================================
# 4. Lumped Port
# ============================================================================
print("  [3.5] Lumped Port ...")

# Create a rectangle sheet at the feed line end (xz plane)
# Rectangle in xz-plane at y = feed_y_end, spanning from z=(-cu_t) to z=(z_top+cu_t)
port_w = balun_w_start   # width in x
port_h = z_top + cu_t + cu_t  # height in z (from ground bottom to top copper)

create_rectangle("Port_Rect",
                 -port_w / 2, feed_y_end, -cu_t,
                 port_w, port_h, "Y")

oModule_Bnd = oDesign.GetModule("BoundarySetup")
oModule_Bnd.AssignLumpedPort(
    ["NAME:Port1",
     "Objects:=", ["Port_Rect"],
     "RenormalizeAllTerminals:=", True,
     "DoDeembed:=", False,
     ["NAME:Modes",
      ["NAME:Mode1",
       "ModeNum:=", 1,
       "UseIntLine:=", True,
       ["NAME:IntLine",
        "Start:=", ["0", str(feed_y_end), str(-cu_t)],
        "End:=",   ["0", str(feed_y_end), str(z_top + cu_t)]],
       "CharImp:=", "Zpi"]]])
print(f"        Port1 (50 Ohm) at y={feed_y_end:.2f}")

# ============================================================================
# 5. Radiation Boundary (Air Box)
# ============================================================================
print("  [3.6] Radiation boundary ...")

pad = lambda_0 * 1e3 / 4.0  # lambda_0/4 padding
air_x = sub_L / 2 + pad
air_y_pos = sub_W / 2 + pad
air_y_neg = sub_W / 2 + pad + balun_L + feed_ext
air_z_top = pad
air_z_bot = pad

create_box("AirBox",
           -air_x, -air_y_neg, -air_z_bot,
           2 * air_x, air_y_pos + air_y_neg, sub_H + air_z_top + air_z_bot,
           "vacuum", solve_inside=True)

oModule_Bnd.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["AirBox"],
     "IsFssReference:=", False,
     "IsForPML:=", False])
print(f"        Padding = {pad:.1f} mm")

# ============================================================================
# 6. Analysis Setup + Sweep
# ============================================================================
print()
print("=" * 70)
print("STEP 4: Analysis Setup")
print("=" * 70)

oModule_Analysis = oDesign.GetModule("AnalysisSetup")
oModule_Analysis.InsertSetup("HfssDriven",
    ["NAME:Setup1",
     "Frequency:=", f"{f0/1e9}GHz",
     "MaxDeltaS:=", 0.02,
     "MaximumPasses:=", 15,
     "MinimumPasses:=", 3,
     "MinimumConvergedPasses:=", 2,
     "PercentRefinement:=", 30,
     "IsEnabled:=", True,
     "BasisOrder:=", 1,
     "UseIterativeSolver:=", False,
     "DoLambdaRefine:=", True,
     "DoMaterialLambdaRefine:=", True,
     "SetLambdaTarget:=", False,
     "Target:=", 0.3333])
print(f"  Setup1: f={f0/1e9:.3f}GHz, MaxDeltaS=0.02, MaxPasses=15")

n_pts = int((f_max - f_min) / f_step) + 1
oModule_Analysis.InsertFrequencySweep("Setup1",
    ["NAME:Sweep1",
     "IsEnabled:=", True,
     "SetupType:=", "LinearCount",
     "StartValue:=", f"{f_min/1e9}GHz",
     "StopValue:=",  f"{f_max/1e9}GHz",
     "Count:=", n_pts,
     "Type:=", "Discrete",
     "SaveFields:=", True,
     "SaveRadFields:=", True])
print(f"  Sweep1: {f_min/1e9:.1f}~{f_max/1e9:.1f}GHz, {n_pts} pts (Discrete)")

# ============================================================================
# 7. Save & Analyze
# ============================================================================
print()
print("  Saving project ...")
hfss.save_project()
print(f"  Saved: {hfss.project_path}")

print()
print("=" * 70)
print("STEP 5: Run Simulation")
print("=" * 70)

t0 = time.time()
print("  Running Setup1 ... (please wait)")
oDesign.Analyze("Setup1")
elapsed = time.time() - t0
print(f"  Done! Elapsed: {elapsed/60:.1f} min")

# ============================================================================
# 8. Extract S11
# ============================================================================
print()
print("=" * 70)
print("STEP 6: S11 Analysis")
print("=" * 70)

oModule_Report = oDesign.GetModule("ReportSetup")

oModule_Report.CreateReport(
    "S11_Report", "Modal Solution Data", "Rectangular Plot",
    "Setup1 : Sweep1",
    ["Domain:=", "Sweep"],
    ["Freq:=", ["All"]],
    ["X Component:=", "Freq",
     "Y Component:=", ["dB(S(Port1,Port1))"]])

s11_csv = os.path.join(PROJECT_DIR, "s11_data.csv")
oModule_Report.ExportToFile("S11_Report", s11_csv)
print(f"  S11 exported: {s11_csv}")

# Parse S11 CSV
freqs_ghz = []
s11_db = []
try:
    with open(s11_csv, 'r') as f:
        reader = csv.reader(f)
        header = next(reader)
        for row in reader:
            if len(row) >= 2:
                try:
                    freqs_ghz.append(float(row[0]))
                    s11_db.append(float(row[1]))
                except ValueError:
                    continue
    freqs_ghz = np.array(freqs_ghz)
    s11_db = np.array(s11_db)

    min_idx = np.argmin(s11_db)
    f_res = freqs_ghz[min_idx]
    s11_min = s11_db[min_idx]

    print(f"\n  -------- S11 Results --------")
    print(f"  Resonance:  {f_res:.4f} GHz")
    print(f"  S11 min:    {s11_min:.2f} dB")

    below_10 = np.where(s11_db < -10)[0]
    if len(below_10) >= 2:
        bw_lo = freqs_ghz[below_10[0]]
        bw_hi = freqs_ghz[below_10[-1]]
        bw_abs = bw_hi - bw_lo
        bw_rel = bw_abs / f_res * 100
        print(f"  -10dB BW:   {bw_lo:.3f} ~ {bw_hi:.3f} GHz")
        print(f"  Abs BW:     {bw_abs*1e3:.1f} MHz")
        print(f"  Rel BW:     {bw_rel:.2f}%")
    else:
        bw_rel = 0
        print("  S11 does not reach -10 dB")

    freq_err = abs(f_res - f0/1e9) / (f0/1e9) * 100
    need_sweep = freq_err > 2.0
    print(f"  Freq error: {freq_err:.2f}%  {'-> NEEDS OPTIMIZATION' if need_sweep else '-> OK'}")

except Exception as e:
    print(f"  S11 parse error: {e}")
    need_sweep = True
    f_res = f0 / 1e9
    s11_min = 0
    bw_rel = 0

# ============================================================================
# 9. S11 Plot
# ============================================================================
print("\n  Generating S11 plot ...")
try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(freqs_ghz, s11_db, 'b-', lw=2, label='S11')
    ax.axhline(-10, color='r', ls='--', lw=1, label='-10 dB')
    ax.axvline(f0/1e9, color='g', ls=':', lw=1, label=f'Target {f0/1e9:.3f} GHz')
    ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'Resonance {f_res:.3f} GHz')
    if len(below_10) >= 2:
        ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green', label=f'BW {bw_abs*1e3:.0f} MHz')
    ax.set_xlabel('Frequency (GHz)', fontsize=13)
    ax.set_ylabel('S11 (dB)', fontsize=13)
    ax.set_title('Printed Dipole Antenna - S11', fontsize=14)
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3)
    ax.set_xlim(f_min/1e9, f_max/1e9)
    s11_img = os.path.join(PROJECT_DIR, "s11_plot.png")
    fig.savefig(s11_img, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  S11 plot: {s11_img}")
except Exception as e:
    print(f"  Plot error: {e}")

# ============================================================================
# 10. Far-Field Radiation Pattern
# ============================================================================
print()
print("=" * 70)
print("STEP 7: Radiation Pattern & Gain")
print("=" * 70)

oModule_RadField = oDesign.GetModule("RadField")
try:
    oModule_RadField.InsertInfiniteSphereDef(
        ["NAME:InfSphere1",
         "UseCustomRadiationSurface:=", False,
         "ThetaStart:=", "-180deg",
         "ThetaStop:=", "180deg",
         "ThetaStep:=", "2deg",
         "PhiStart:=", "0deg",
         "PhiStop:=", "360deg",
         "PhiStep:=", "2deg",
         "UseLocalCS:=", False])
    print("  Infinite Sphere defined")
except Exception as e:
    print(f"  InfSphere note: {e}")

# E-plane (phi=0, xz plane)
try:
    oModule_Report.CreateReport(
        "E_Plane", "Far Fields", "Radiation Pattern",
        "Setup1 : LastAdaptive",
        ["Context:=", "InfSphere1"],
        ["Theta:=", ["All"], "Phi:=", ["0deg"], "Freq:=", [f"{f0/1e9}GHz"]],
        ["X Component:=", "Theta",
         "Y Component:=", ["GainTotal"]])
    e_csv = os.path.join(PROJECT_DIR, "e_plane.csv")
    oModule_Report.ExportToFile("E_Plane", e_csv)
    print(f"  E-plane exported: {e_csv}")
except Exception as e:
    print(f"  E-plane error: {e}")
    e_csv = None

# H-plane (phi=90, yz plane)
try:
    oModule_Report.CreateReport(
        "H_Plane", "Far Fields", "Radiation Pattern",
        "Setup1 : LastAdaptive",
        ["Context:=", "InfSphere1"],
        ["Theta:=", ["All"], "Phi:=", ["90deg"], "Freq:=", [f"{f0/1e9}GHz"]],
        ["X Component:=", "Theta",
         "Y Component:=", ["GainTotal"]])
    h_csv = os.path.join(PROJECT_DIR, "h_plane.csv")
    oModule_Report.ExportToFile("H_Plane", h_csv)
    print(f"  H-plane exported: {h_csv}")
except Exception as e:
    print(f"  H-plane error: {e}")
    h_csv = None

# Parse & plot patterns
max_gain = 0
try:
    def read_csv_pattern(path):
        th, g = [], []
        with open(path, 'r') as f:
            reader = csv.reader(f)
            next(reader)
            for row in reader:
                if len(row) >= 2:
                    try:
                        th.append(float(row[0]))
                        g.append(float(row[1]))
                    except ValueError:
                        continue
        return np.array(th), np.array(g)

    if e_csv and h_csv:
        th_e, g_e = read_csv_pattern(e_csv)
        th_h, g_h = read_csv_pattern(h_csv)

        max_gain_e = np.max(g_e) if len(g_e) > 0 else 0
        max_gain_h = np.max(g_h) if len(g_h) > 0 else 0
        max_gain = max(max_gain_e, max_gain_h)

        print(f"\n  -------- Radiation Results --------")
        print(f"  E-plane max gain: {max_gain_e:.2f} dBi")
        print(f"  H-plane max gain: {max_gain_h:.2f} dBi")
        print(f"  Antenna gain:     {max_gain:.2f} dBi")

        # Plot
        fig, (ax1, ax2) = plt.subplots(1, 2, subplot_kw={'projection': 'polar'},
                                        figsize=(14, 6))
        ax1.plot(np.radians(th_e), g_e, 'b-', lw=2)
        ax1.set_title(f'E-Plane (xz, phi=0)\nGain={max_gain_e:.2f} dBi', pad=20)
        ax1.set_theta_zero_location('N')
        ax1.set_theta_direction(-1)

        ax2.plot(np.radians(th_h), g_h, 'r-', lw=2)
        ax2.set_title(f'H-Plane (yz, phi=90)\nGain={max_gain_h:.2f} dBi', pad=20)
        ax2.set_theta_zero_location('N')
        ax2.set_theta_direction(-1)

        pat_img = os.path.join(PROJECT_DIR, "radiation_pattern.png")
        fig.savefig(pat_img, dpi=150, bbox_inches='tight')
        plt.close(fig)
        print(f"  Pattern plot: {pat_img}")

except Exception as e:
    print(f"  Pattern analysis error: {e}")

# ============================================================================
# 11. Parameter Sweep (if resonance deviates > 2%)
# ============================================================================
if need_sweep:
    print()
    print("=" * 70)
    print("STEP 8: Parameter Sweep Optimization")
    print("=" * 70)

    if f_res > f0 / 1e9:
        arm_min = dipole_arm
        arm_max = dipole_arm * 1.3
    else:
        arm_min = dipole_arm * 0.7
        arm_max = dipole_arm

    arm_step_val = 1.0
    print(f"  Sweep range: {arm_min:.0f} ~ {arm_max:.0f} mm, step {arm_step_val} mm")

    # Create design variable
    try:
        oDesign.ChangeProperty(
            ["NAME:AllTabs",
             ["NAME:LocalVariableTab",
              ["NAME:PropServers", "LocalVariables"],
              ["NAME:NewProps",
               ["NAME:arm_len",
                "PropType:=", "VariableProp",
                "UserDef:=", True,
                "Value:=", f"{dipole_arm}mm"]]]])

        oModule_Optim = oDesign.GetModule("Optimetrics")
        oModule_Optim.InsertSetup("OptiParametric",
            ["NAME:ParamSweep1",
             "IsEnabled:=", True,
             ["NAME:ProdOptiSetupDataV2",
              "SaveFields:=", True,
              "CopyMesh:=", False,
              "SolveWithCopiedMeshOnly:=", True],
             ["NAME:StartingPoint"],
             "Sim. Setups:=", ["Setup1"],
             ["NAME:Sweeps",
              ["NAME:SweepDefinition",
               "Variable:=", "arm_len",
               "Data:=", f"LIN {arm_min}mm {arm_max}mm {arm_step_val}mm",
               "OffsetF1:=", False,
               "Synchronize:=", 0]],
             ["NAME:Sweep Operations"],
             ["NAME:Goals"]])

        print("  Running parameter sweep ...")
        hfss.save_project()
        t2 = time.time()
        oModule_Optim.SolveSetup("ParamSweep1")
        t2e = time.time() - t2
        print(f"  Sweep done! {t2e/60:.1f} min")

        # Analyze sweep results
        best_arm = dipole_arm
        best_err = 999
        best_s11_val = 0
        arm_vals = np.arange(arm_min, arm_max + arm_step_val/2, arm_step_val)

        for av in arm_vals:
            try:
                rname = f"S11_arm{av:.0f}"
                sweep_csv = os.path.join(PROJECT_DIR, f"s11_arm_{av:.0f}.csv")
                oModule_Report.CreateReport(
                    rname, "Modal Solution Data", "Rectangular Plot",
                    "Setup1 : Sweep1",
                    ["Domain:=", "Sweep", "arm_len:=", [f"{av}mm"]],
                    ["Freq:=", ["All"]],
                    ["X Component:=", "Freq",
                     "Y Component:=", ["dB(S(Port1,Port1))"]])
                oModule_Report.ExportToFile(rname, sweep_csv)

                fl, sl = [], []
                with open(sweep_csv, 'r') as f:
                    rd = csv.reader(f)
                    next(rd)
                    for row in rd:
                        if len(row) >= 2:
                            try:
                                fl.append(float(row[0]))
                                sl.append(float(row[1]))
                            except ValueError:
                                continue
                if sl:
                    idx = np.argmin(sl)
                    fr = fl[idx]
                    err = abs(fr - f0/1e9)
                    print(f"    arm={av:.0f}mm: f_res={fr:.4f}GHz, S11={sl[idx]:.2f}dB, err={err*1e3:.1f}MHz")
                    if err < best_err:
                        best_err = err
                        best_arm = av
                        best_s11_val = sl[idx]
            except Exception as e:
                print(f"    arm={av:.0f}mm: error - {e}")

        print(f"\n  ======== Optimization Result ========")
        print(f"  Best arm length: {best_arm:.1f} mm")
        print(f"  Freq error:      {best_err*1e3:.1f} MHz")
        print(f"  S11:             {best_s11_val:.2f} dB")

    except Exception as e:
        print(f"  Sweep error: {e}")
        import traceback
        traceback.print_exc()
else:
    print("\n  Frequency error < 2%, no optimization needed")

# ============================================================================
# 12. Summary
# ============================================================================
print()
print("=" * 70)
print("DESIGN SUMMARY")
print("=" * 70)
print(f"  Antenna:       Printed Dipole + Microstrip Balun")
print(f"  Target freq:   {f0/1e9:.3f} GHz")
print(f"  Substrate:     FR4 (er={eps_r}), {sub_L}x{sub_W}x{sub_H} mm")
print(f"  Dipole arm:    {dipole_arm:.1f} mm (initial)")
if need_sweep:
    print(f"  Optimized arm: {best_arm:.1f} mm")
print(f"  Dipole width:  {dipole_w:.1f} mm")
print(f"  Balun length:  {balun_L:.2f} mm")
print(f"  Resonance:     {f_res:.4f} GHz")
print(f"  S11 min:       {s11_min:.2f} dB")
if bw_rel > 0:
    print(f"  Rel bandwidth: {bw_rel:.2f}%")
print(f"  Max gain:      {max_gain:.2f} dBi")
print()
print(f"  Output files:")
print(f"    S11 CSV:   {os.path.join(PROJECT_DIR, 's11_data.csv')}")
print(f"    S11 Plot:  {os.path.join(PROJECT_DIR, 's11_plot.png')}")
if e_csv:
    print(f"    E-plane:   {e_csv}")
if h_csv:
    print(f"    H-plane:   {h_csv}")
print(f"    Pattern:   {os.path.join(PROJECT_DIR, 'radiation_pattern.png')}")
print(f"    Log:       {LOG_FILE}")
print()

# ============================================================================
# 13. Release AEDT
# ============================================================================
print("Releasing AEDT ...")
try:
    hfss.release_desktop()
except Exception:
    pass
print("DONE!")
