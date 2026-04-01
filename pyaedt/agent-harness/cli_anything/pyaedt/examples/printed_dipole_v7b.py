#!/usr/bin/env python3
"""
printed_dipole_v7b.py - 印刷偶极子天线 v7b
  - 手动 WavePort (基板边缘微带端口)
  - 截断地板
  - SaveRadFields=True
  - 远场设置在仿真后添加
"""

import numpy as np
import os, sys, time, csv, traceback

os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
os.makedirs(PROJECT_DIR, exist_ok=True)
LOG_FILE = os.path.join(PROJECT_DIR, "dipole_v7b.log")

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

# Compatibility patches
try:
    from pyaedt import desktop as _dm
    _o1 = _dm.Desktop.__init__
    def _p1(self, *a, **kw):
        _o1(self, *a, **kw)
        if not hasattr(self, 'student_version'): self.student_version = False
    _dm.Desktop.__init__ = _p1
except Exception: pass
try:
    import pyaedt.application.Design as _dd
    _o2 = _dd.DesignSettings.__init__
    def _p2(self, app):
        try: _o2(self, app)
        except AttributeError:
            self._app = app; self.design_settings = None; self.manipulate_inputs = None
    _dd.DesignSettings.__init__ = _p2
except Exception: pass

from pyaedt import Hfss

# ============================================================================
# 1. 设计参数
# ============================================================================
print("="*70)
print("STEP 1: Design Parameters")
print("="*70)

c = 3e8; f0 = 2.217e9; eps_r = 4.4
lambda_0 = c / f0
lambda_g = lambda_0 / np.sqrt(eps_r)
eps_eff = (eps_r + 1) / 2.0

# 基板 (mm)
sub_L = 120.0; sub_W = 80.0; sub_H = 1.6; cu_t = 0.035

# 偶极子 (mm)
dipole_arm = 34.0
dipole_w = 3.0
gap = 1.0

# 巴伦 (mm) - lambda_g/4 渐变微带
balun_L = lambda_g * 1e3 / 4.0  # ~16.1mm
feed_w = 3.0

# 馈线延长到基板底部
feed_ext = 20.0

# 频率扫描
f_min = 1.5e9; f_max = 3.0e9; f_step = 0.01e9
n_pts = int((f_max - f_min) / f_step) + 1

# AirBox padding >= lambda_0/4
pad = lambda_0 * 1e3 / 4.0

# 坐标计算
z_top = sub_H
balun_y_top = -dipole_w / 2
balun_y_bot = balun_y_top - balun_L
sub_y_top = sub_W / 2
sub_y_bot = balun_y_bot - feed_ext
sub_total_y = sub_y_top - sub_y_bot
gnd_y_top = balun_y_top  # 截断地板

# WavePort 尺寸 (在 y=sub_y_bot 处, 微带截面)
port_width  = 15 * sub_H    # x方向宽度 ≈ 24mm (足够宽)
port_height = 8 * sub_H    # z方向高度 ≈ 12.8mm (含地+基板+信号+空气)
port_z_bot  = -(port_height - sub_H) / 2.0  # 端口z起点 (居中在基板附近)

print(f"  f0              = {f0/1e9:.3f} GHz")
print(f"  lambda_0        = {lambda_0*1e3:.2f} mm")
print(f"  dipole_arm      = {dipole_arm:.1f} mm")
print(f"  total dipole    = {2*dipole_arm+gap:.1f} mm")
print(f"  balun_L         = {balun_L:.2f} mm")
print(f"  sub Y range     = [{sub_y_bot:.1f}, {sub_y_top:.1f}]")
print(f"  ground Y range  = [{sub_y_bot:.1f}, {gnd_y_top:.1f}]")
print(f"  WavePort at:    y={sub_y_bot:.1f}, w={port_width:.1f}, h={port_height:.1f}")
print(f"  freq sweep:     {f_min/1e9:.1f}~{f_max/1e9:.1f} GHz, {n_pts} pts")
print()

# ============================================================================
# 2. 启动 AEDT
# ============================================================================
print("="*70)
print("STEP 2: Launch AEDT")
print("="*70)

PROJECT_NAME = "Printed_Dipole_v7b"
hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, PROJECT_NAME),
    designname="Dipole_v7b",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=True,
    specified_version="2019.1",
)
print(f"  Project: {hfss.project_name}")
print(f"  Design:  {hfss.design_name}")

hfss.modeler.model_units = "mm"
oDesign = hfss.odesign
oEditor = oDesign.SetActiveEditor("3D Modeler")
oBnd    = oDesign.GetModule("BoundarySetup")

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

def create_rect_sheet(name, x, y, z, w, h, axis="Y"):
    """Create a 2D sheet. For axis='Y', the rectangle lies in the XZ plane at the given Y."""
    oEditor.CreateRectangle(
        ["NAME:RectangleParameters",
         "IsCovered:=", True,
         "XStart:=", f"{x}mm", "YStart:=", f"{y}mm", "ZStart:=", f"{z}mm",
         "Width:=", f"{w}mm", "Height:=", f"{h}mm",
         "WhichAxis:=", axis],
        ["NAME:Attributes",
         "Name:=", name, "Flags:=", "", "Color:=", "(0 0 255)",
         "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global",
         "UDMId:=", "", "MaterialValue:=", '"vacuum"',
         "SurfaceMaterialValue:=", '""', "SolveInside:=", True,
         "IsMaterialEditable:=", True, "UseMaterialAppearance:=", False,
         "IsLightweight:=", False])

# ============================================================================
# 3. 创建天线几何模型
# ============================================================================
print()
print("="*70)
print("STEP 3: Antenna Geometry")
print("="*70)

# 3.1 FR4 基板
print(f"  [3.1] Substrate")
create_box("Substrate", -sub_L/2, sub_y_bot, 0, sub_L, sub_total_y, sub_H, "FR4_epoxy")

# 3.2 截断地板 (底面铜, 巴伦+馈线区下方)
gnd_y_size = gnd_y_top - sub_y_bot
print(f"  [3.2] Ground (truncated)")
create_box("Ground", -sub_L/2, sub_y_bot, -cu_t, sub_L, gnd_y_size, cu_t, "copper")

# 3.3 偶极子臂
left_x = -(gap/2 + dipole_arm)
right_x = gap/2
print(f"  [3.3] Dipole arms (arm={dipole_arm}mm)")
create_box("Dipole_Left",  left_x,  -dipole_w/2, z_top, dipole_arm, dipole_w, cu_t, "copper")
create_box("Dipole_Right", right_x, -dipole_w/2, z_top, dipole_arm, dipole_w, cu_t, "copper")

# 3.4 巴伦
balun_segs = 8
seg_len = balun_L / balun_segs
print(f"  [3.4] Balun: {balun_segs} segs")
for i in range(balun_segs):
    seg_y = balun_y_top - seg_len * (i + 1)
    create_box(f"Balun_{i}", -feed_w/2, seg_y, z_top, feed_w, seg_len, cu_t, "copper")

# 3.5 馈线
feed_y_start = sub_y_bot
feed_y_end   = balun_y_bot
feed_len_y   = feed_y_end - feed_y_start
print(f"  [3.5] Feed Line (y=[{feed_y_start:.1f},{feed_y_end:.2f}])")
create_box("Feed_Line", -feed_w/2, feed_y_start, z_top, feed_w, feed_len_y, cu_t, "copper")

# 3.6 AirBox
airbox_xmin = -(sub_L/2 + pad)
airbox_ymin = sub_y_bot - pad
airbox_ymax = sub_y_top + pad
air_x_size  = sub_L + 2*pad
air_y_size  = airbox_ymax - airbox_ymin
air_z_size  = sub_H + 2*pad
print(f"  [3.6] AirBox")
create_box("Air", airbox_xmin, airbox_ymin, -pad, air_x_size, air_y_size, air_z_size, "vacuum", True)

oBnd.AssignRadiation(
    ["NAME:Rad1", "Objects:=", ["Air"], "IsFssReference:=", False, "IsForPML:=", False])
print("        Radiation boundary OK")

# ============================================================================
# 4. WavePort (手动创建微带端口)
# ============================================================================
print()
print("="*70)
print("STEP 4: WavePort (manual microstrip port)")
print("="*70)

# 在基板底边 (y=sub_y_bot) 创建一个矩形 sheet, 位于 XZ 平面
# 端口以 feed line 为中心, 宽度覆盖足够范围, 高度覆盖地+基板+信号线+空气
port_x_start = -port_width / 2
port_z_start = port_z_bot
print(f"  Port sheet at y={sub_y_bot:.1f}")
print(f"    x: [{port_x_start:.1f}, {port_x_start+port_width:.1f}]")
print(f"    z: [{port_z_start:.2f}, {port_z_start+port_height:.2f}]")

create_rect_sheet("Port_Sheet", port_x_start, sub_y_bot, port_z_start, port_width, port_height, "Y")

# 在此 sheet 上分配 WavePort
# 积分线: 从地板顶面 (z=0) 到信号线底面 (z=sub_H), x=0
try:
    oBnd.AssignWavePort(
        ["NAME:Port1",
         "Objects:=", ["Port_Sheet"],
         "NumModes:=", 1,
         "RenormalizeAllTerminals:=", True,
         "UseLineModeAlignment:=", False,
         "DoDeembed:=", False,
         ["NAME:Modes",
          ["NAME:Mode1",
           "ModeNum:=", 1,
           "UseIntLine:=", True,
           ["NAME:IntLine",
            "Start:=", [f"0mm", f"{sub_y_bot}mm", "0mm"],
            "End:=",   [f"0mm", f"{sub_y_bot}mm", f"{sub_H}mm"]],
           "CharImp:=", "Zpi"]],
         "ShowReporterFilter:=", False,
         "ReporterFilter:=", [True],
         "UseAnalyticAlignment:=", False])
    print("  WavePort assigned OK")
except Exception as e:
    print(f"  WavePort ERROR: {e}")
    # Fallback: try AutoIdentifyPorts on the port sheet face
    try:
        port_faces = oEditor.GetFaceIDs("Port_Sheet")
        print(f"  Trying AutoIdentifyPorts on Port_Sheet face {port_faces[0]}")
        oBnd.AutoIdentifyPorts(
            ["NAME:Faces", int(port_faces[0])], True,
            ["NAME:ReferenceConductors", "Ground"],
            "Port1", True)
        print("  AutoIdentifyPorts on Port_Sheet OK")
    except Exception as e2:
        print(f"  AutoIdentifyPorts fallback ERROR: {e2}")

# Check terminals
try:
    terms = oBnd.GetExcitationsOfType("Terminal")
    print(f"  Terminals: {terms}")
    terminal_name = str(terms[0]) if terms else "Port1_T1"
except Exception:
    terminal_name = "Port1_T1"
print(f"  Using terminal: {terminal_name}")

# ============================================================================
# 5. 分析设置
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
print(f"  Setup1: f0={f0/1e9:.3f}GHz")

oAnalysis.InsertFrequencySweep("Setup1",
    ["NAME:Sweep1",
     "IsEnabled:=", True, "SetupType:=", "LinearCount",
     "StartValue:=", f"{f_min/1e9}GHz", "StopValue:=", f"{f_max/1e9}GHz",
     "Count:=", n_pts,
     "Type:=", "Discrete", "SaveFields:=", True, "SaveRadFields:=", True])
print(f"  Sweep1: {f_min/1e9:.1f}~{f_max/1e9:.1f}GHz, {n_pts}pts, SaveRadFields=True")

# ============================================================================
# 6. 验证 & 仿真
# ============================================================================
print()
print("="*70)
print("STEP 6: Validate & Simulate")
print("="*70)

hfss.save_project()
print(f"  Project saved")

v = oDesign.ValidateDesign()
print(f"  Validation: {v} {'(OK)' if v else '(FAILED - attempting anyway)'}")

# Print messages
try:
    oMsg = hfss.odesktop.GetMessages("", "", 2)  # severity 2 = error
    if oMsg:
        for m in oMsg[:5]:
            print(f"    MSG: {m}")
except Exception:
    pass

t0 = time.time()
print("  Running Setup1 ...")
try:
    oDesign.Analyze("Setup1")
    t_sim = time.time() - t0
    print(f"  Simulation OK! {t_sim:.0f}s ({t_sim/60:.1f}min)")
    sim_ok = True
except Exception as e:
    print(f"  Simulation FAILED: {e}")
    sim_ok = False

    # Try DrivenModal if Terminal fails
    print("\n  Retrying with DrivenModal solution type...")
    try:
        oDesign.SetSolutionType("DrivenModal")
        oDesign.Analyze("Setup1")
        t_sim = time.time() - t0
        print(f"  DrivenModal OK! {t_sim:.0f}s")
        sim_ok = True
        terminal_name = None  # no terminal in modal
    except Exception as e2:
        print(f"  DrivenModal also FAILED: {e2}")

if not sim_ok:
    print("\n  FATAL: Cannot run simulation. Exiting.")
    hfss.save_project()
    sys.exit(1)

# ============================================================================
# 7. S11 提取
# ============================================================================
print()
print("="*70)
print("STEP 7: S11 Extraction")
print("="*70)

oSolutions = oDesign.GetModule("Solutions")
oReport    = oDesign.GetModule("ReportSetup")

s1p_path = os.path.join(PROJECT_DIR, "dipole_v7b.s1p")
s11_csv  = os.path.join(PROJECT_DIR, "s11_v7b.csv")

# Touchstone
try:
    oSolutions.ExportNetworkData("", ["Setup1:Sweep1"], 3, s1p_path, ["All"], True, 50)
    print(f"  Touchstone: {s1p_path}")
except Exception as e:
    print(f"  Touchstone error: {e}")

# S11 CSV
if terminal_name:
    s11_expr = f"dB(St({terminal_name},{terminal_name}))"
    report_type = "Terminal Solution Data"
else:
    s11_expr = "dB(S(Port1,Port1))"
    report_type = "Modal Solution Data"

try:
    oReport.CreateReport(
        "S11_Report", report_type, "Rectangular Plot",
        "Setup1 : Sweep1", [],
        ["Freq:=", ["All"]],
        ["X Component:=", "Freq", "Y Component:=", [s11_expr]],
        [])
    oReport.ExportToFile("S11_Report", s11_csv)
    print(f"  S11 CSV: {s11_csv}")
except Exception as e:
    print(f"  S11 report error: {e}")

# ============================================================================
# 8. S11 分析
# ============================================================================
def parse_s11_csv(path):
    f_list, s_list = [], []
    with open(path, 'r') as fh:
        rd = csv.reader(fh); next(rd)
        for row in rd:
            if len(row) >= 2:
                try: f_list.append(float(row[0])); s_list.append(float(row[1]))
                except ValueError: continue
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
    freq_err = abs(f_res - target_f_ghz) / target_f_ghz * 100
    return dict(f_res=f_res, s11_min=s11_min, bw_lo=bw_lo, bw_hi=bw_hi,
                bw_abs=bw_abs, bw_rel=bw_rel, freq_err=freq_err)

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
        print("  WARNING: S11 does not reach -10 dB")
    print(f"  Freq error:  {freq_err:.2f}%")

    need_sweep = freq_err > 2.0 or s11_min > -10
except Exception as e:
    print(f"  Analysis error: {e}"); traceback.print_exc()
    need_sweep = True; f_res = f0/1e9; s11_min = 0; bw_abs = bw_rel = 0
    bw_lo = bw_hi = f_res; freq_err = 999; freqs_ghz = np.array([]); s11_db = np.array([])

# S11 Plot
try:
    import matplotlib; matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(freqs_ghz, s11_db, 'b-', lw=2, label='S11')
    ax.axhline(-10, color='r', ls='--', lw=1, label='-10 dB')
    ax.axvline(f0/1e9, color='g', ls=':', lw=1, label=f'Target {f0/1e9:.3f} GHz')
    ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'Res {f_res:.3f} GHz')
    if bw_abs > 0:
        ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green',
                   label=f'BW {bw_abs*1e3:.0f}MHz ({bw_rel:.1f}%)')
    ax.set_xlabel('Frequency (GHz)'); ax.set_ylabel('S11 (dB)')
    ax.set_title(f'Printed Dipole v7b (arm={dipole_arm}mm) S11')
    ax.legend(); ax.grid(True, alpha=0.3); ax.set_xlim(f_min/1e9, f_max/1e9)
    fig.savefig(os.path.join(PROJECT_DIR, "s11_v7b.png"), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  S11 plot saved")
except Exception as e:
    print(f"  Plot error: {e}")

# ============================================================================
# 9. 参数扫描优化
# ============================================================================
if need_sweep:
    print()
    print("="*70)
    print("STEP 9: Optimization")
    print("="*70)

    current_arm = dipole_arm
    best_arm = current_arm; best_f_res = f_res; best_s11 = s11_min
    best_bw_rel = bw_rel; best_bw_abs = bw_abs

    for op in range(10):
        if best_f_res > 0:
            new_arm = round(current_arm * (best_f_res / (f0/1e9)), 1)
        else:
            new_arm = current_arm + 2.0

        if abs(new_arm - current_arm) < 0.3:
            print(f"\n  Iter {op+1}: converged (delta < 0.3mm)")
            break
        if new_arm > 80 or new_arm < 15:
            print(f"\n  Iter {op+1}: arm {new_arm:.1f}mm out of range, stopping")
            break

        print(f"\n  Iter {op+1}: arm {current_arm:.1f} -> {new_arm:.1f} mm")

        try:
            oEditor.Delete(["NAME:Selections", "Selections:=", "Dipole_Left,Dipole_Right"])
        except Exception as e:
            print(f"    Delete error: {e}"); break

        new_left_x = -(gap/2 + new_arm)
        create_box("Dipole_Left",  new_left_x,  -dipole_w/2, z_top, new_arm, dipole_w, cu_t, "copper")
        create_box("Dipole_Right", gap/2, -dipole_w/2, z_top, new_arm, dipole_w, cu_t, "copper")

        hfss.save_project()
        t_opt = time.time()
        print("    Analyzing...")
        oDesign.Analyze("Setup1")
        print(f"    Done in {time.time()-t_opt:.0f}s")

        opt_csv = os.path.join(PROJECT_DIR, f"s11_v7b_opt{op+1}.csv")
        try:
            rname = f"S11_Opt{op+1}"
            oReport.CreateReport(
                rname, report_type, "Rectangular Plot",
                "Setup1 : Sweep1", [],
                ["Freq:=", ["All"]],
                ["X Component:=", "Freq", "Y Component:=", [s11_expr]],
                [])
            oReport.ExportToFile(rname, opt_csv)

            opt_f, opt_s = parse_s11_csv(opt_csv)
            r2 = analyze_s11(opt_f, opt_s, f0/1e9)

            print(f"    f={r2['f_res']:.4f}GHz, S11={r2['s11_min']:.2f}dB, err={r2['freq_err']:.2f}%")
            if r2['bw_abs'] > 0:
                print(f"    BW: {r2['bw_lo']:.3f}~{r2['bw_hi']:.3f}GHz ({r2['bw_rel']:.1f}%)")

            current_arm = new_arm
            best_f_res = r2['f_res']; best_s11 = r2['s11_min']; best_arm = new_arm
            best_bw_rel = r2['bw_rel']; best_bw_abs = r2['bw_abs']
            f_res = r2['f_res']; s11_min = r2['s11_min']
            bw_rel = r2['bw_rel']; bw_abs = r2['bw_abs']
            bw_lo = r2['bw_lo']; bw_hi = r2['bw_hi']
            freq_err = r2['freq_err']; freqs_ghz = opt_f; s11_db = opt_s

            if r2['freq_err'] <= 2.0 and r2['s11_min'] < -10:
                print(f"    ** CONVERGED! **")
                break
            if r2['freq_err'] <= 2.0 and r2['s11_min'] > -10:
                print(f"    Freq OK but impedance mismatch. Fine-tuning...")
                # 微调: 尝试 ±1mm
                for delta in [-1, +1, -2, +2]:
                    trial_arm = round(new_arm + delta, 1)
                    print(f"    Fine: arm={trial_arm:.1f}")
                    try:
                        oEditor.Delete(["NAME:Selections", "Selections:=", "Dipole_Left,Dipole_Right"])
                        create_box("Dipole_Left",  -(gap/2+trial_arm), -dipole_w/2, z_top, trial_arm, dipole_w, cu_t, "copper")
                        create_box("Dipole_Right", gap/2, -dipole_w/2, z_top, trial_arm, dipole_w, cu_t, "copper")
                        hfss.save_project()
                        oDesign.Analyze("Setup1")
                        fc = os.path.join(PROJECT_DIR, f"s11_v7b_fine_{trial_arm}.csv")
                        fn = f"S11_Fine_{trial_arm}"
                        oReport.CreateReport(fn, report_type, "Rectangular Plot",
                            "Setup1 : Sweep1", [], ["Freq:=", ["All"]],
                            ["X Component:=", "Freq", "Y Component:=", [s11_expr]], [])
                        oReport.ExportToFile(fn, fc)
                        ff, ss = parse_s11_csv(fc)
                        rf = analyze_s11(ff, ss, f0/1e9)
                        print(f"      f={rf['f_res']:.4f}, S11={rf['s11_min']:.2f}")
                        if rf['s11_min'] < best_s11:
                            best_s11 = rf['s11_min']; best_arm = trial_arm
                            f_res = rf['f_res']; s11_min = rf['s11_min']
                            bw_rel = rf['bw_rel']; bw_abs = rf['bw_abs']
                            bw_lo = rf['bw_lo']; bw_hi = rf['bw_hi']
                            freq_err = rf['freq_err']; freqs_ghz = ff; s11_db = ss
                        if rf['s11_min'] < -10 and rf['freq_err'] <= 3:
                            print(f"      ** Found good match **")
                            break
                    except Exception as ef:
                        print(f"      Fine error: {ef}"); continue
                break

        except Exception as e:
            print(f"    Opt error: {e}"); traceback.print_exc(); break
    else:
        print(f"\n  Max iter. Best: arm={best_arm}, f={best_f_res:.4f}, S11={best_s11:.2f}")

    # 最终 S11 图
    if len(freqs_ghz) > 0:
        try:
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.plot(freqs_ghz, s11_db, 'b-', lw=2, label='S11 (optimized)')
            ax.axhline(-10, color='r', ls='--', lw=1, label='-10 dB')
            ax.axvline(f0/1e9, color='g', ls=':', lw=1, label=f'Target {f0/1e9:.3f}')
            ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'Res {f_res:.3f}')
            if bw_abs > 0:
                ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green',
                           label=f'BW {bw_abs*1e3:.0f}MHz ({bw_rel:.1f}%)')
            ax.set_xlabel('Frequency (GHz)'); ax.set_ylabel('S11 (dB)')
            ax.set_title(f'Printed Dipole v7b (arm={best_arm:.1f}mm, opt)')
            ax.legend(); ax.grid(True, alpha=0.3); ax.set_xlim(f_min/1e9, f_max/1e9)
            fig.savefig(os.path.join(PROJECT_DIR, "s11_v7b_opt.png"), dpi=150, bbox_inches='tight')
            plt.close(fig)
            print(f"\n  Opt S11 plot saved")
        except Exception as ee:
            print(f"  Plot error: {ee}")
else:
    best_arm = dipole_arm
    print("  No optimization needed")

# ============================================================================
# 10. 远场辐射方向图
# ============================================================================
print()
print("="*70)
print("STEP 10: Radiation Patterns")
print("="*70)

max_gain = 0
e_csv = os.path.join(PROJECT_DIR, "e_plane_v7b.csv")
h_csv = os.path.join(PROJECT_DIR, "h_plane_v7b.csv")

# 尝试添加远场设置 (仿真后)
oRadField = oDesign.GetModule("RadField")
ff_setup_name = None

for method_name, method_func in [
    ("InsertFarFieldSphereSetup", lambda: oRadField.InsertFarFieldSphereSetup(
        ["NAME:FF_Setup", "UseCustomRadiationSurface:=", False,
         "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "2deg",
         "PhiStart:=", "0deg", "PhiStop:=", "360deg", "PhiStep:=", "2deg",
         "UseLocalCS:=", False])),
    ("InsertInfiniteSphereDef", lambda: oRadField.InsertInfiniteSphereDef(
        ["NAME:FF_Setup2", "UseCustomRadiationSurface:=", False,
         "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "2deg",
         "PhiStart:=", "0deg", "PhiStop:=", "360deg", "PhiStep:=", "2deg",
         "UseLocalCS:=", False])),
]:
    try:
        method_func()
        ff_setup_name = "FF_Setup" if "Sphere" in method_name else "FF_Setup2"
        print(f"  {method_name}: OK -> {ff_setup_name}")
        break
    except Exception as e:
        print(f"  {method_name}: failed ({e})")

# 列出存在的远场设置
try:
    existing_ff = oRadField.GetSetupNames()
    print(f"  Existing FF setups: {existing_ff}")
    if existing_ff and not ff_setup_name:
        ff_setup_name = str(existing_ff[0])
except Exception:
    pass

if not ff_setup_name:
    # 尝试创建一个简单的 infinite sphere "3D"
    try:
        oRadField.InsertFarFieldSphereSetup(
            ["NAME:3D", "UseCustomRadiationSurface:=", False,
             "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "5deg",
             "PhiStart:=", "0deg", "PhiStop:=", "360deg", "PhiStep:=", "5deg",
             "UseLocalCS:=", False])
        ff_setup_name = "3D"
        print(f"  Created sphere '3D'")
    except Exception:
        print("  No far field setup available")

# 提取方向图
for plane_name, phi_val, csv_path in [("E_Plane_v7b", "0deg", e_csv), ("H_Plane_v7b", "90deg", h_csv)]:
    if not ff_setup_name:
        print(f"  {plane_name}: skipped (no FF setup)")
        continue
    try:
        oReport.CreateReport(
            plane_name, "Far Fields", "Radiation Pattern",
            "Setup1 : Sweep1",
            ["Context:=", ff_setup_name],
            ["Theta:=", ["All"], "Phi:=", [phi_val], "Freq:=", [f"{f0/1e9}GHz"]],
            ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]],
            [])
        oReport.ExportToFile(plane_name, csv_path)
        print(f"  {plane_name}: OK -> {csv_path}")
    except Exception as e:
        print(f"  {plane_name}: FAILED ({e})")
        # 尝试 LastAdaptive
        try:
            rn2 = f"{plane_name}_LA"
            oReport.CreateReport(
                rn2, "Far Fields", "Radiation Pattern",
                "Setup1 : LastAdaptive",
                ["Context:=", ff_setup_name],
                ["Theta:=", ["All"], "Phi:=", [phi_val], "Freq:=", [f"{f0/1e9}GHz"]],
                ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]],
                [])
            oReport.ExportToFile(rn2, csv_path)
            print(f"  {rn2}: OK -> {csv_path}")
        except Exception as e2:
            print(f"  {plane_name}_LA also FAILED ({e2})")

# 绘图
def read_csv_pattern(path):
    th, g = [], []
    with open(path, 'r') as f:
        rd = csv.reader(f); next(rd)
        for row in rd:
            if len(row) >= 2:
                try: th.append(float(row[0])); g.append(float(row[1]))
                except ValueError: continue
    return np.array(th), np.array(g)

try:
    if os.path.exists(e_csv) and os.path.exists(h_csv):
        th_e, g_e = read_csv_pattern(e_csv)
        th_h, g_h = read_csv_pattern(h_csv)
        if len(g_e) > 0 and len(g_h) > 0:
            max_gain_e = np.max(g_e); max_gain_h = np.max(g_h)
            max_gain = max(max_gain_e, max_gain_h)
            print(f"\n  E-plane max gain: {max_gain_e:.2f} dBi")
            print(f"  H-plane max gain: {max_gain_h:.2f} dBi")
            print(f"  Antenna gain:     {max_gain:.2f} dBi")

            fig, (ax1, ax2) = plt.subplots(1, 2, subplot_kw={'projection':'polar'}, figsize=(14,6))
            ax1.plot(np.radians(th_e), g_e, 'b-', lw=2)
            ax1.set_title(f'E-Plane (phi=0)\nGain={max_gain_e:.2f}dBi', pad=20)
            ax1.set_theta_zero_location('N'); ax1.set_theta_direction(-1)
            ax2.plot(np.radians(th_h), g_h, 'r-', lw=2)
            ax2.set_title(f'H-Plane (phi=90)\nGain={max_gain_h:.2f}dBi', pad=20)
            ax2.set_theta_zero_location('N'); ax2.set_theta_direction(-1)
            fig.suptitle(f'Printed Dipole @ {f0/1e9:.3f} GHz', fontsize=14)
            fig.savefig(os.path.join(PROJECT_DIR, "radiation_v7b.png"), dpi=150, bbox_inches='tight')
            plt.close(fig)
            print(f"  Pattern plot saved")
    elif os.path.exists(e_csv) or os.path.exists(h_csv):
        which = e_csv if os.path.exists(e_csv) else h_csv
        th, g = read_csv_pattern(which)
        if len(g) > 0:
            max_gain = float(np.max(g))
            print(f"  Max gain from available pattern: {max_gain:.2f} dBi")
except Exception as e:
    print(f"  Pattern analysis error: {e}")

# ============================================================================
# 11. 设计总结
# ============================================================================
print()
print("="*70)
print("="*70)
print("DESIGN SUMMARY")
print("="*70)
print()
print(f"  Antenna:          Printed Dipole + Truncated GND + Microstrip Balun")
print(f"  Target:           {f0/1e9:.3f} GHz")
print(f"  Substrate:        FR4 (er={eps_r}), {sub_L}x{sub_total_y:.1f}x{sub_H} mm")
print(f"  Optimized Arm:    {best_arm:.1f} mm  (total {2*best_arm+gap:.1f} mm)")
print(f"  Dipole Width:     {dipole_w:.1f} mm")
print(f"  Balun Length:     {balun_L:.2f} mm (lambda_g/4)")
print(f"  Ground Truncated: y < {gnd_y_top:.1f} mm")
print()
print(f"  Resonance Freq:   {f_res:.4f} GHz (error {freq_err:.2f}%)")
print(f"  Min S11:          {s11_min:.2f} dB")
if bw_abs > 0:
    print(f"  -10dB BW:         {bw_lo:.3f} ~ {bw_hi:.3f} GHz")
    print(f"  Abs BW:           {bw_abs*1e3:.1f} MHz")
    print(f"  Rel BW:           {bw_rel:.2f}%")
if max_gain > 0:
    print(f"  Max Gain:         {max_gain:.2f} dBi")
print()

hfss.save_project()
print("SIMULATION COMPLETE")
print("="*70)
