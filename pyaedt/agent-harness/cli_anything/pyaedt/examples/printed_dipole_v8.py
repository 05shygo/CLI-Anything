#!/usr/bin/env python3
"""
printed_dipole_v8.py - 印刷偶极子天线 v8
  关键修复: 让饲线、基板和地板延伸到 AirBox y_min 边界
  这样 AutoIdentifyPorts 能在该面上识别微带截面并创建端口
"""

import numpy as np
import os, sys, time, csv, traceback

os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
os.makedirs(PROJECT_DIR, exist_ok=True)
LOG_FILE = os.path.join(PROJECT_DIR, "dipole_v8.log")

class DualLogger:
    def __init__(self, logpath):
        self.terminal = sys.stdout
        self.log = open(logpath, "w", encoding="utf-8")
    def write(self, msg):
        try: self.terminal.write(msg)
        except UnicodeEncodeError: self.terminal.write(msg.encode('ascii','replace').decode())
        self.log.write(msg); self.log.flush()
    def flush(self):
        self.terminal.flush(); self.log.flush()
sys.stdout = DualLogger(LOG_FILE)
sys.stderr = sys.stdout

try:
    from pyaedt import desktop as _dm
    _o1 = _dm.Desktop.__init__
    def _p1(self, *a, **kw):
        _o1(self, *a, **kw)
        if not hasattr(self, 'student_version'): self.student_version = False
    _dm.Desktop.__init__ = _p1
except: pass
try:
    import pyaedt.application.Design as _dd
    _o2 = _dd.DesignSettings.__init__
    def _p2(self, app):
        try: _o2(self, app)
        except AttributeError:
            self._app = app; self.design_settings = None; self.manipulate_inputs = None
    _dd.DesignSettings.__init__ = _p2
except: pass

from pyaedt import Hfss

# ============================================================================
# 1. Parameters
# ============================================================================
print("="*70)
print("STEP 1: Parameters")
print("="*70)

c = 3e8; f0 = 2.217e9; eps_r = 4.4
lambda_0 = c / f0                       # ~135.3mm
lambda_g = lambda_0 / np.sqrt(eps_r)    # ~64.5mm

sub_L  = 120.0   # x 方向基板长度
sub_H  = 1.6     # 基板厚度
cu_t   = 0.035   # 铜厚
dipole_arm = 34.0 # 初始臂长
dipole_w   = 3.0  # 臂宽
gap    = 1.0      # 两臂间隙
balun_L = lambda_g * 1e3 / 4.0  # ~16.13mm
feed_w  = 3.0    # 微带线宽
pad = lambda_0 * 1e3 / 4.0      # ~33.83mm

# 坐标系
z_top = sub_H
balun_y_top = -dipole_w / 2        # 巴伦顶端 y
balun_y_bot = balun_y_top - balun_L # 巴伦底端 y

# 关键: 基板、地板和馈线必须延伸到 AirBox y_min 面
# AirBox y_min = sub_y_bot - pad, 所以 sub_y_bot 就是基板底部
# 为确保导体穿过 AirBox y_min 面, 让基板延伸到 AirBox y_min
# 即 sub_y_bot = airbox_ymin
# 这意味着: no padding on y_min side, substrate IS the boundary

sub_y_top = 40.0                        # 偶极子上方空间
# 像 v6 一样，让基板底部远离巴伦，馈线延伸下去
# v6 用了 sub_y_bot = balun_y_bot - feed_ext where feed_ext was large
# 关键: 馈线/地板/基板都要延伸到 airbox y_min 面
# 方案: sub_y_bot = -(balun_L + dipole_w/2 + 56mm) ≈ -73.6mm (like v6)
feed_ext_total = 56.0  # 从巴伦底端继续延伸
sub_y_bot = balun_y_bot - feed_ext_total  # ≈ -73.76mm

sub_total_y = sub_y_top - sub_y_bot

# 地板 (截断: 只在巴伦和馈线下方)
gnd_y_top = balun_y_top
gnd_y_size = gnd_y_top - sub_y_bot

# AirBox
airbox_ymin = sub_y_bot  # 基板底部 = AirBox 底部! 导体直接在面上
airbox_ymax = sub_y_top + pad
airbox_xmin = -(sub_L/2 + pad)
air_x = sub_L + 2*pad
air_y = airbox_ymax - airbox_ymin
air_z = sub_H + 2*pad

f_min = 1.5e9; f_max = 3.0e9; f_step = 0.01e9
n_pts = int((f_max - f_min) / f_step) + 1

print(f"  f0={f0/1e9:.3f}GHz, lambda_0={lambda_0*1e3:.1f}mm")
print(f"  dipole_arm={dipole_arm}mm, total={2*dipole_arm+gap}mm")
print(f"  balun_L={balun_L:.2f}mm")
print(f"  sub Y=[{sub_y_bot:.2f}, {sub_y_top:.1f}] ({sub_total_y:.1f}mm)")
print(f"  gnd Y=[{sub_y_bot:.2f}, {gnd_y_top:.1f}] ({gnd_y_size:.1f}mm)")
print(f"  feed: y=[{sub_y_bot:.2f}, {balun_y_bot:.2f}] ({feed_ext_total:.1f}mm)")
print(f"  AirBox ymin={airbox_ymin:.2f} (== sub_y_bot)")
print(f"  pad={pad:.1f}mm")
print()

# ============================================================================
# 2. Launch AEDT
# ============================================================================
print("="*70)
print("STEP 2: Launch AEDT")
print("="*70)

hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "Printed_Dipole_v8"),
    designname="Dipole_v8",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=True,
    specified_version="2019.1",
)
print(f"  Project: {hfss.project_name}")

hfss.modeler.model_units = "mm"
oDesign = hfss.odesign
oEditor = oDesign.SetActiveEditor("3D Modeler")
oBnd = oDesign.GetModule("BoundarySetup")

def create_box(name, x, y, z, dx, dy, dz, mat, solve_inside=None):
    if solve_inside is None:
        solve_inside = mat.lower() not in ("copper", "pec", "aluminum")
    oEditor.CreateBox(
        ["NAME:BoxParameters",
         "XPosition:=", f"{x}mm", "YPosition:=", f"{y}mm", "ZPosition:=", f"{z}mm",
         "XSize:=", f"{dx}mm", "YSize:=", f"{dy}mm", "ZSize:=", f"{dz}mm"],
        ["NAME:Attributes",
         "Name:=", name, "Flags:=", "", "Color:=", "(143 175 131)",
         "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
         "MaterialValue:=", f'"{mat}"', "SurfaceMaterialValue:=", '""',
         "SolveInside:=", solve_inside, "IsMaterialEditable:=", True,
         "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# ============================================================================
# 3. Geometry
# ============================================================================
print()
print("="*70)
print("STEP 3: Geometry")
print("="*70)

# 3.1 基板
create_box("Substrate", -sub_L/2, sub_y_bot, 0, sub_L, sub_total_y, sub_H, "FR4_epoxy")
print(f"  Substrate: {sub_L}x{sub_total_y:.1f}x{sub_H}")

# 3.2 截断地板
create_box("Ground", -sub_L/2, sub_y_bot, -cu_t, sub_L, gnd_y_size, cu_t, "copper")
print(f"  Ground (truncated): y=[{sub_y_bot:.2f}, {gnd_y_top}]")

# 3.3 偶极子臂
left_x = -(gap/2 + dipole_arm)
create_box("Dipole_Left", left_x, -dipole_w/2, z_top, dipole_arm, dipole_w, cu_t, "copper")
create_box("Dipole_Right", gap/2, -dipole_w/2, z_top, dipole_arm, dipole_w, cu_t, "copper")
print(f"  Dipole arms: {dipole_arm}mm each")

# 3.4 巴伦
balun_segs = 8
seg_len = balun_L / balun_segs
for i in range(balun_segs):
    seg_y = balun_y_top - seg_len * (i + 1)
    create_box(f"Balun_{i}", -feed_w/2, seg_y, z_top, feed_w, seg_len, cu_t, "copper")
print(f"  Balun: {balun_segs} segs, L={balun_L:.2f}mm")

# 3.5 馈线 (延伸到基板底部 = AirBox 底部)
feed_y_len = balun_y_bot - sub_y_bot
create_box("Feed_Line", -feed_w/2, sub_y_bot, z_top, feed_w, feed_y_len, cu_t, "copper")
print(f"  Feed Line: y=[{sub_y_bot:.2f}, {balun_y_bot:.2f}] ({feed_y_len:.1f}mm)")

# 3.6 AirBox
create_box("Air", airbox_xmin, airbox_ymin, -pad, air_x, air_y, air_z, "vacuum", True)
print(f"  AirBox: {air_x:.1f}x{air_y:.1f}x{air_z:.1f}, ymin={airbox_ymin:.2f}")

oBnd.AssignRadiation(
    ["NAME:Rad1", "Objects:=", ["Air"], "IsFssReference:=", False, "IsForPML:=", False])
print("  Radiation BC OK")

# ============================================================================
# 4. Port
# ============================================================================
print()
print("="*70)
print("STEP 4: Port (AutoIdentifyPorts on y_min face)")
print("="*70)

air_faces = oEditor.GetFaceIDs("Air")
print(f"  AirBox faces: {air_faces}")

# Find y_min face
face_info = []
for fid in air_faces:
    try:
        c = oEditor.GetFaceCenter(int(fid))
        y_val = float(c[1])
        face_info.append((int(fid), y_val))
        print(f"    Face {fid}: y={y_val:.2f}")
    except: pass

ymin_face = min(face_info, key=lambda x: x[1])
print(f"  y_min face: ID={ymin_face[0]} (y={ymin_face[1]:.2f})")

oBnd.AutoIdentifyPorts(
    ["NAME:Faces", ymin_face[0]], True,
    ["NAME:ReferenceConductors", "Ground"],
    "Port1", True)
print("  AutoIdentifyPorts OK")

terms = oBnd.GetExcitationsOfType("Terminal")
print(f"  Terminals: {terms}")
terminal_name = str(terms[0]) if terms else "Feed_Line_T1"
print(f"  Using terminal: {terminal_name}")

# ============================================================================
# 5. Analysis Setup (no far field)
# ============================================================================
print()
print("="*70)
print("STEP 5: Analysis Setup")
print("="*70)

oAnalysis = oDesign.GetModule("AnalysisSetup")
oAnalysis.InsertSetup("HfssDriven",
    ["NAME:Setup1",
     "Frequency:=", f"{f0/1e9}GHz",
     "MaxDeltaS:=", 0.02, "MaximumPasses:=", 15,
     "MinimumPasses:=", 2, "MinimumConvergedPasses:=", 2,
     "PercentRefinement:=", 30, "IsEnabled:=", True,
     "BasisOrder:=", 1, "UseIterativeSolver:=", False,
     "DoLambdaRefine:=", True, "DoMaterialLambdaRefine:=", True,
     "SetLambdaTarget:=", False, "Target:=", 0.3333])
print(f"  Setup1 OK")

oAnalysis.InsertFrequencySweep("Setup1",
    ["NAME:Sweep1",
     "IsEnabled:=", True, "SetupType:=", "LinearCount",
     "StartValue:=", f"{f_min/1e9}GHz", "StopValue:=", f"{f_max/1e9}GHz",
     "Count:=", n_pts,
     "Type:=", "Discrete", "SaveFields:=", True, "SaveRadFields:=", True])
print(f"  Sweep1 OK (SaveRadFields=True)")

# ============================================================================
# 6. Validate & Simulate
# ============================================================================
print()
print("="*70)
print("STEP 6: Validate & Simulate")
print("="*70)

hfss.save_project()
v = oDesign.ValidateDesign()
print(f"  Validation: {v}")

if not v:
    print("  WARNING: Validation failed!")
    exc = oBnd.GetExcitations()
    print(f"    Excitations: {exc}")

t0 = time.time()
print("  Running Setup1 ...")
try:
    oDesign.Analyze("Setup1")
    t_sim = time.time() - t0
    print(f"  OK! {t_sim:.0f}s ({t_sim/60:.1f}min)")
    sim_ok = True
except Exception as e:
    print(f"  FAILED: {e}")
    sim_ok = False

if not sim_ok:
    hfss.save_project()
    print("  FATAL: Simulation failed")
    sys.exit(1)

# ============================================================================
# 7. S11
# ============================================================================
print()
print("="*70)
print("STEP 7: S11 Extraction")
print("="*70)

oSolutions = oDesign.GetModule("Solutions")
oReport = oDesign.GetModule("ReportSetup")

s1p_path = os.path.join(PROJECT_DIR, "dipole_v8.s1p")
s11_csv = os.path.join(PROJECT_DIR, "s11_v8.csv")

try:
    oSolutions.ExportNetworkData("", ["Setup1:Sweep1"], 3, s1p_path, ["All"], True, 50)
    print(f"  Touchstone: {s1p_path}")
except Exception as e:
    print(f"  Touchstone error: {e}")

try:
    oReport.CreateReport(
        "S11_v8", "Terminal Solution Data", "Rectangular Plot",
        "Setup1 : Sweep1", [],
        ["Freq:=", ["All"]],
        ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal_name},{terminal_name}))"]],
        [])
    oReport.ExportToFile("S11_v8", s11_csv)
    print(f"  S11 CSV: {s11_csv}")
except Exception as e:
    print(f"  S11 error: {e}")

# Parse & analyze
def parse_s11_csv(path):
    f_list, s_list = [], []
    with open(path, 'r') as fh:
        rd = csv.reader(fh); next(rd)
        for row in rd:
            if len(row) >= 2:
                try: f_list.append(float(row[0])); s_list.append(float(row[1]))
                except: continue
    return np.array(f_list), np.array(s_list)

def analyze_s11(freqs, s11, target_f_ghz):
    min_idx = np.argmin(s11)
    f_res = freqs[min_idx]; s11_min = s11[min_idx]
    below_10 = np.where(s11 < -10)[0]
    if len(below_10) >= 2:
        bw_lo = freqs[below_10[0]]; bw_hi = freqs[below_10[-1]]
        bw_abs = bw_hi - bw_lo; bw_rel = bw_abs / f_res * 100
    else:
        bw_lo = bw_hi = f_res; bw_abs = bw_rel = 0
    return dict(f_res=f_res, s11_min=s11_min, bw_lo=bw_lo, bw_hi=bw_hi,
                bw_abs=bw_abs, bw_rel=bw_rel,
                freq_err=abs(f_res-target_f_ghz)/target_f_ghz*100)

print()
print("="*70)
print("STEP 8: S11 Analysis")
print("="*70)

try:
    freqs_ghz, s11_db = parse_s11_csv(s11_csv)
    r = analyze_s11(freqs_ghz, s11_db, f0/1e9)
    f_res = r['f_res']; s11_min = r['s11_min']
    bw_lo = r['bw_lo']; bw_hi = r['bw_hi']
    bw_abs = r['bw_abs']; bw_rel = r['bw_rel']
    freq_err = r['freq_err']
    print(f"  Resonance:   {f_res:.4f} GHz")
    print(f"  S11 min:     {s11_min:.2f} dB")
    if bw_abs > 0:
        print(f"  -10dB BW:    {bw_lo:.3f} ~ {bw_hi:.3f} GHz ({bw_rel:.2f}%)")
    else:
        print("  WARNING: S11 > -10 dB")
    print(f"  Freq error:  {freq_err:.2f}%")
    need_opt = freq_err > 2.0 or s11_min > -10
except Exception as e:
    print(f"  Error: {e}"); traceback.print_exc()
    need_opt = True; f_res = f0/1e9; s11_min = 0; bw_abs = bw_rel = 0
    bw_lo = bw_hi = f_res; freq_err = 999
    freqs_ghz = np.array([]); s11_db = np.array([])

# S11 plot
try:
    import matplotlib; matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(10,6))
    ax.plot(freqs_ghz, s11_db, 'b-', lw=2, label='S11')
    ax.axhline(-10, color='r', ls='--', lw=1, label='-10dB')
    ax.axvline(f0/1e9, color='g', ls=':', lw=1, label=f'Target {f0/1e9:.3f}')
    ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'Res {f_res:.3f}')
    if bw_abs > 0:
        ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green', label=f'BW {bw_abs*1e3:.0f}MHz')
    ax.set_xlabel('Frequency (GHz)'); ax.set_ylabel('S11 (dB)')
    ax.set_title(f'Printed Dipole v8 (arm={dipole_arm}mm)')
    ax.legend(); ax.grid(True, alpha=0.3); ax.set_xlim(f_min/1e9, f_max/1e9)
    fig.savefig(os.path.join(PROJECT_DIR, "s11_v8.png"), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  Plot saved")
except Exception as e:
    print(f"  Plot error: {e}")

# ============================================================================
# 9. Optimization
# ============================================================================
if need_opt:
    print()
    print("="*70)
    print("STEP 9: Optimization")
    print("="*70)

    current_arm = dipole_arm
    best_arm = current_arm; best_s11 = s11_min

    for op in range(10):
        if f_res > 0:
            new_arm = round(current_arm * (f_res / (f0/1e9)), 1)
        else:
            new_arm = current_arm + 2.0

        if abs(new_arm - current_arm) < 0.3:
            print(f"\n  Iter {op+1}: converged (delta<0.3mm)"); break
        if new_arm > 80 or new_arm < 15:
            print(f"\n  Iter {op+1}: arm out of range [{new_arm}mm]"); break

        print(f"\n  Iter {op+1}: arm {current_arm:.1f} -> {new_arm:.1f} mm")
        try:
            oEditor.Delete(["NAME:Selections", "Selections:=", "Dipole_Left,Dipole_Right"])
        except: break

        create_box("Dipole_Left", -(gap/2+new_arm), -dipole_w/2, z_top, new_arm, dipole_w, cu_t, "copper")
        create_box("Dipole_Right", gap/2, -dipole_w/2, z_top, new_arm, dipole_w, cu_t, "copper")
        hfss.save_project()

        t_opt = time.time()
        oDesign.Analyze("Setup1")
        print(f"    Sim: {time.time()-t_opt:.0f}s")

        csv_p = os.path.join(PROJECT_DIR, f"s11_v8_opt{op+1}.csv")
        rn = f"S11_Opt{op+1}"
        oReport.CreateReport(rn, "Terminal Solution Data", "Rectangular Plot",
            "Setup1 : Sweep1", [], ["Freq:=", ["All"]],
            ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal_name},{terminal_name}))"]],
            [])
        oReport.ExportToFile(rn, csv_p)
        ff, ss = parse_s11_csv(csv_p)
        r2 = analyze_s11(ff, ss, f0/1e9)

        print(f"    f={r2['f_res']:.4f}GHz, S11={r2['s11_min']:.2f}dB, err={r2['freq_err']:.2f}%")
        if r2['bw_abs'] > 0:
            print(f"    BW: {r2['bw_lo']:.3f}~{r2['bw_hi']:.3f} ({r2['bw_rel']:.1f}%)")

        current_arm = new_arm; f_res = r2['f_res']; s11_min = r2['s11_min']
        bw_rel = r2['bw_rel']; bw_abs = r2['bw_abs']
        bw_lo = r2['bw_lo']; bw_hi = r2['bw_hi']
        freq_err = r2['freq_err']; freqs_ghz = ff; s11_db = ss
        best_arm = new_arm; best_s11 = r2['s11_min']

        if r2['freq_err'] <= 2.0 and r2['s11_min'] < -10:
            print(f"    ** CONVERGED! **"); break
        if r2['freq_err'] <= 2.0:
            print(f"    Freq OK but S11={r2['s11_min']:.2f}dB"); break

    # Final S11 plot
    if len(freqs_ghz) > 0:
        try:
            fig, ax = plt.subplots(figsize=(10,6))
            ax.plot(freqs_ghz, s11_db, 'b-', lw=2, label='S11 (opt)')
            ax.axhline(-10, color='r', ls='--', lw=1); ax.axvline(f0/1e9, color='g', ls=':', lw=1)
            ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'Res {f_res:.3f}')
            if bw_abs > 0:
                ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green', label=f'BW {bw_abs*1e3:.0f}MHz ({bw_rel:.1f}%)')
            ax.set_xlabel('Freq (GHz)'); ax.set_ylabel('S11 (dB)')
            ax.set_title(f'Printed Dipole v8 (arm={best_arm:.1f}mm, opt)')
            ax.legend(); ax.grid(True, alpha=0.3)
            fig.savefig(os.path.join(PROJECT_DIR, "s11_v8_opt.png"), dpi=150, bbox_inches='tight')
            plt.close(fig)
        except: pass
else:
    best_arm = dipole_arm

# ============================================================================
# 10. Far Field
# ============================================================================
print()
print("="*70)
print("STEP 10: Radiation Patterns")
print("="*70)

max_gain = 0
e_csv = os.path.join(PROJECT_DIR, "e_plane_v8.csv")
h_csv = os.path.join(PROJECT_DIR, "h_plane_v8.csv")

oRadField = oDesign.GetModule("RadField")
ff_name = None
for mname, mfunc in [
    ("InsertFarFieldSphereSetup", lambda: oRadField.InsertFarFieldSphereSetup(
        ["NAME:FF1", "UseCustomRadiationSurface:=", False,
         "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "2deg",
         "PhiStart:=", "0deg", "PhiStop:=", "360deg", "PhiStep:=", "2deg",
         "UseLocalCS:=", False])),
    ("InsertInfiniteSphereDef", lambda: oRadField.InsertInfiniteSphereDef(
        ["NAME:FF2", "UseCustomRadiationSurface:=", False,
         "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "2deg",
         "PhiStart:=", "0deg", "PhiStop:=", "360deg", "PhiStep:=", "2deg",
         "UseLocalCS:=", False])),
]:
    try:
        mfunc()
        ff_name = "FF1" if "Sphere" in mname else "FF2"
        print(f"  {mname}: OK -> {ff_name}"); break
    except Exception as e:
        print(f"  {mname}: {e}")

try:
    existing = oRadField.GetSetupNames()
    print(f"  FF setups: {existing}")
    if existing and not ff_name:
        ff_name = str(existing[0])
except: pass

for pname, phi, csvp in [("E_Plane_v8", "0deg", e_csv), ("H_Plane_v8", "90deg", h_csv)]:
    if not ff_name:
        print(f"  {pname}: no FF setup"); continue
    for sweep in ["Setup1 : Sweep1", "Setup1 : LastAdaptive"]:
        try:
            oReport.CreateReport(
                pname, "Far Fields", "Radiation Pattern", sweep,
                ["Context:=", ff_name],
                ["Theta:=", ["All"], "Phi:=", [phi], "Freq:=", [f"{f0/1e9}GHz"]],
                ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]], [])
            oReport.ExportToFile(pname, csvp)
            print(f"  {pname}: OK ({sweep})"); break
        except Exception as e:
            print(f"  {pname} ({sweep}): {e}")

def read_pattern(path):
    th, g = [], []
    with open(path, 'r') as f:
        rd = csv.reader(f); next(rd)
        for row in rd:
            if len(row) >= 2:
                try: th.append(float(row[0])); g.append(float(row[1]))
                except: continue
    return np.array(th), np.array(g)

try:
    if os.path.exists(e_csv) and os.path.exists(h_csv):
        th_e, g_e = read_pattern(e_csv)
        th_h, g_h = read_pattern(h_csv)
        if len(g_e) > 0:
            max_e = np.max(g_e); max_h = np.max(g_h) if len(g_h) > 0 else 0
            max_gain = max(max_e, max_h)
            print(f"\n  E-plane max: {max_e:.2f} dBi")
            print(f"  H-plane max: {max_h:.2f} dBi")
            print(f"  Gain:        {max_gain:.2f} dBi")

            fig, (ax1, ax2) = plt.subplots(1, 2, subplot_kw={'projection':'polar'}, figsize=(14,6))
            ax1.plot(np.radians(th_e), g_e, 'b-', lw=2)
            ax1.set_title(f'E-Plane (phi=0)\nGain={max_e:.2f}dBi', pad=20)
            ax2.plot(np.radians(th_h), g_h, 'r-', lw=2)
            ax2.set_title(f'H-Plane (phi=90)\nGain={max_h:.2f}dBi', pad=20)
            fig.suptitle(f'Printed Dipole @ {f0/1e9:.3f}GHz', fontsize=14)
            fig.savefig(os.path.join(PROJECT_DIR, "radiation_v8.png"), dpi=150, bbox_inches='tight')
            plt.close(fig)
            print(f"  Pattern plot saved")
except Exception as e:
    print(f"  Pattern error: {e}")

# ============================================================================
# Summary
# ============================================================================
print()
print("="*70)
print("DESIGN SUMMARY")
print("="*70)
print(f"  Type:       Printed Dipole + Truncated GND + Microstrip Balun")
print(f"  Target:     {f0/1e9:.3f} GHz")
print(f"  Substrate:  FR4 (er={eps_r}), {sub_L}x{sub_total_y:.1f}x{sub_H}mm")
print(f"  Arm:        {best_arm:.1f}mm (total {2*best_arm+gap:.1f}mm)")
print(f"  Resonance:  {f_res:.4f}GHz (err={freq_err:.2f}%)")
print(f"  S11:        {s11_min:.2f}dB")
if bw_abs > 0:
    print(f"  -10dB BW:   {bw_lo:.3f}~{bw_hi:.3f}GHz ({bw_abs*1e3:.0f}MHz, {bw_rel:.2f}%)")
if max_gain > 0:
    print(f"  Gain:       {max_gain:.2f}dBi")
print()

hfss.save_project()
print("COMPLETE")
print("="*70)
