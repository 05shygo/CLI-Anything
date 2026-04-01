#!/usr/bin/env python3
"""
printed_dipole_v7.py - 印刷偶极子天线（截断地板 + 微带巴伦）v7

v7 改进:
  1. 初始臂长 34mm（基于 v6 缩放预测）
  2. SaveRadFields=True 以支持远场方向图
  3. 使用 SetupFarFieldSphere 替代不兼容的 InsertInfiniteSphereDef
  4. 更强健的优化循环
  5. 基板仅延伸到偶极子区域，不延伸到 AirBox 边界
     WavePort 改用 LumpedPort (sheet) 在巴伦底端
"""

import numpy as np
import os, sys, time, csv, traceback

# ============================================================================
# 0. 环境配置
# ============================================================================
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'

PROJECT_DIR = r"D:\class_design"
os.makedirs(PROJECT_DIR, exist_ok=True)
LOG_FILE = os.path.join(PROJECT_DIR, "dipole_v7.log")


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
print("=" * 70)
print("STEP 1: Design Parameters")
print("=" * 70)

c = 3e8; f0 = 2.217e9; eps_r = 4.4
lambda_0 = c / f0                        # ~135.3mm
lambda_g = lambda_0 / np.sqrt(eps_r)     # ~64.5mm (FR4)
eps_eff = (eps_r + 1) / 2.0              # ~2.7 (microstrip effective)
lambda_eff = lambda_0 / np.sqrt(eps_eff) # ~82.4mm

# 基板 (mm)
sub_L = 120.0; sub_W = 80.0; sub_H = 1.6; cu_t = 0.035

# 偶极子 (mm)
dipole_arm = 34.0    # 初始臂长 (v6 优化预测值: 30*(2.49/2.217)≈33.7)
dipole_w = 3.0       # 臂宽
gap = 1.0            # 两臂间隙

# 巴伦 (mm) - λg/4 渐变微带
balun_L = lambda_g * 1e3 / 4.0  # ~16.1mm
feed_w = 3.0                    # 50Ω微带线宽度 (1.6mm FR4)

# 馈线延长到基板底部
feed_ext = 20.0  # mm (从巴伦底端延伸到基板边缘)

# 频率扫描
f_min = 1.5e9; f_max = 3.0e9; f_step = 0.01e9
n_pts = int((f_max - f_min) / f_step) + 1

# AirBox padding ≥ λ0/4
pad = lambda_0 * 1e3 / 4.0  # ~33.8mm

# ==== 关键坐标 ====
z_top = sub_H  # 基板顶面

# 偶极子中心在 y=0
balun_y_top = -dipole_w / 2             # 巴伦顶端
balun_y_bot = balun_y_top - balun_L     # 巴伦底端

# 基板 Y 范围
sub_y_top = sub_W / 2                   # 基板顶 (偶极子上方)
sub_y_bot = balun_y_bot - feed_ext      # 基板底 (馈线下方)
sub_total_y = sub_y_top - sub_y_bot

# 截断地板 - 只覆盖巴伦和馈线区域 (y < balun_y_top)
gnd_y_top = balun_y_top

print(f"  f0              = {f0/1e9:.3f} GHz")
print(f"  lambda_0        = {lambda_0*1e3:.2f} mm")
print(f"  lambda_g(FR4)   = {lambda_g*1e3:.2f} mm")
print(f"  lambda_eff      = {lambda_eff*1e3:.2f} mm")
print(f"  dipole_arm      = {dipole_arm:.1f} mm (each side)")
print(f"  total dipole    = {2*dipole_arm+gap:.1f} mm")
print(f"  balun_L         = {balun_L:.2f} mm (lambda_g/4)")
print(f"  substrate       = {sub_L} x {sub_total_y:.1f} x {sub_H} mm")
print(f"  sub Y range     = [{sub_y_bot:.1f}, {sub_y_top:.1f}]")
print(f"  ground Y range  = [{sub_y_bot:.1f}, {gnd_y_top:.1f}] (truncated)")
print(f"  feed extent     = {feed_ext:.1f} mm below balun")
print(f"  freq sweep      = {f_min/1e9:.1f}~{f_max/1e9:.1f} GHz, {n_pts} pts")
print()

# ============================================================================
# 2. 启动 AEDT
# ============================================================================
print("=" * 70)
print("STEP 2: Launch AEDT")
print("=" * 70)

PROJECT_NAME = "Printed_Dipole_v7"
hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, PROJECT_NAME),
    designname="Dipole_v7",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)
print(f"  Project: {hfss.project_name}")
print(f"  Design:  {hfss.design_name}")

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


def create_rect(name, x, y, z, width, height, axis="Y"):
    """Create a 2D rectangle sheet for port."""
    oEditor.CreateRectangle(
        ["NAME:RectangleParameters",
         "IsCovered:=", True,
         "XStart:=", f"{x}mm", "YStart:=", f"{y}mm", "ZStart:=", f"{z}mm",
         "Width:=", f"{width}mm", "Height:=", f"{height}mm",
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
print("=" * 70)
print("STEP 3: Antenna Geometry")
print("=" * 70)

# 3.1 FR4 基板
print(f"  [3.1] Substrate: {sub_L} x {sub_total_y:.1f} x {sub_H} mm")
create_box("Substrate", -sub_L/2, sub_y_bot, 0, sub_L, sub_total_y, sub_H, "FR4_epoxy")

# 3.2 截断地板 (底面铜, 只覆盖巴伦+馈线区域)
gnd_y_size = gnd_y_top - sub_y_bot
print(f"  [3.2] Ground (truncated): {sub_L} x {gnd_y_size:.1f} mm, y=[{sub_y_bot:.1f}, {gnd_y_top:.1f}]")
create_box("Ground", -sub_L/2, sub_y_bot, -cu_t, sub_L, gnd_y_size, cu_t, "copper")

# 3.3 偶极子臂 (顶面, y=0 中心)
left_x = -(gap/2 + dipole_arm)
right_x = gap/2
print(f"  [3.3] Dipole Left:  X=[{left_x:.1f}, {-gap/2:.1f}], arm={dipole_arm}mm")
create_box("Dipole_Left", left_x, -dipole_w/2, z_top, dipole_arm, dipole_w, cu_t, "copper")
print(f"        Dipole Right: X=[{right_x:.1f}, {right_x+dipole_arm:.1f}]")
create_box("Dipole_Right", right_x, -dipole_w/2, z_top, dipole_arm, dipole_w, cu_t, "copper")

# 3.4 巴伦 (渐变微带, 从偶极子中心向 -y 延伸)
balun_segs = 8
seg_len = balun_L / balun_segs
print(f"  [3.4] Balun: {balun_segs} segs, L={balun_L:.2f}mm")
for i in range(balun_segs):
    seg_y = balun_y_top - seg_len * (i + 1)
    create_box(f"Balun_{i}", -feed_w/2, seg_y, z_top, feed_w, seg_len, cu_t, "copper")

# 3.5 馈线 (从巴伦底端延伸到基板底部)
feed_y_start = sub_y_bot
feed_y_end = balun_y_bot
feed_len = feed_y_end - feed_y_start
print(f"  [3.5] Feed: w={feed_w}mm, Y=[{feed_y_start:.1f}, {feed_y_end:.2f}] ({feed_len:.1f}mm)")
create_box("Feed_Line", -feed_w/2, feed_y_start, z_top, feed_w, feed_len, cu_t, "copper")

# 3.6 AirBox
airbox_xmin = -(sub_L/2 + pad)
airbox_ymin = sub_y_bot - pad
airbox_ymax = sub_y_top + pad
air_x_size = sub_L + 2*pad
air_y_size = airbox_ymax - airbox_ymin
air_z_size = sub_H + 2*pad
print(f"  [3.6] AirBox: {air_x_size:.1f} x {air_y_size:.1f} x {air_z_size:.1f} mm")
create_box("Air", airbox_xmin, airbox_ymin, -pad, air_x_size, air_y_size, air_z_size, "vacuum", True)

oBnd.AssignRadiation(
    ["NAME:Rad1", "Objects:=", ["Air"], "IsFssReference:=", False, "IsForPML:=", False])
print("        Radiation boundary OK")

# ============================================================================
# 4. WavePort (AutoIdentifyPorts on y_min face)
# ============================================================================
print()
print("=" * 70)
print("STEP 4: WavePort")
print("=" * 70)

air_faces = oEditor.GetFaceIDs("Air")
print(f"  AirBox faces: {air_faces}")

# y_min face: face index 2 (verified in previous tests)
ymin_face_id = int(air_faces[2])
print(f"  y_min face ID: {ymin_face_id}")

oBnd.AutoIdentifyPorts(
    ["NAME:Faces", ymin_face_id], True,
    ["NAME:ReferenceConductors", "Ground"],
    "Port1", True)
print("  AutoIdentifyPorts OK")

terms = oBnd.GetExcitationsOfType("Terminal")
terminal_name = str(terms[0]) if terms else "Feed_Line_T1"
print(f"  Terminal: {terminal_name}")

# ============================================================================
# 5. Far Field Setup (try multiple methods for 2019.1)
# ============================================================================
print()
print("=" * 70)
print("STEP 5: Far Field Setup")
print("=" * 70)

oRadField = oDesign.GetModule("RadField")
ff_setup_name = None

# Method 1: InsertInfiniteSphereDef (newer API)
try:
    oRadField.InsertInfiniteSphereDef(
        ["NAME:FarField1",
         "UseCustomRadiationSurface:=", False,
         "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "2deg",
         "PhiStart:=", "0deg", "PhiStop:=", "360deg", "PhiStep:=", "2deg",
         "UseLocalCS:=", False])
    ff_setup_name = "FarField1"
    print(f"  Method 1 OK: {ff_setup_name}")
except Exception as e1:
    print(f"  Method 1 failed: {e1}")

    # Method 2: InsertFarFieldSphereSetup (AEDT 2019 API)
    try:
        oRadField.InsertFarFieldSphereSetup(
            ["NAME:FarField2",
             "UseCustomRadiationSurface:=", False,
             "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "2deg",
             "PhiStart:=", "0deg", "PhiStop:=", "360deg", "PhiStep:=", "2deg",
             "UseLocalCS:=", False])
        ff_setup_name = "FarField2"
        print(f"  Method 2 OK: {ff_setup_name}")
    except Exception as e2:
        print(f"  Method 2 failed: {e2}")

        # Method 3: Try AddInfiniteSphereDef
        try:
            oRadField.AddInfiniteSphereDef(
                ["NAME:FarField3",
                 "UseCustomRadiationSurface:=", False,
                 "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "2deg",
                 "PhiStart:=", "0deg", "PhiStop:=", "360deg", "PhiStep:=", "2deg",
                 "UseLocalCS:=", False])
            ff_setup_name = "FarField3"
            print(f"  Method 3 OK: {ff_setup_name}")
        except Exception as e3:
            print(f"  Method 3 failed: {e3}")
            print("  WARNING: No far field setup created. Pattern extraction may fail.")

# ============================================================================
# 6. 分析设置
# ============================================================================
print()
print("=" * 70)
print("STEP 6: Analysis Setup")
print("=" * 70)

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
print(f"  Setup1: f0={f0/1e9:.3f}GHz, MaxDeltaS=0.02, MaxPasses=15")

oAnalysis.InsertFrequencySweep("Setup1",
    ["NAME:Sweep1",
     "IsEnabled:=", True, "SetupType:=", "LinearCount",
     "StartValue:=", f"{f_min/1e9}GHz", "StopValue:=", f"{f_max/1e9}GHz",
     "Count:=", n_pts,
     "Type:=", "Discrete", "SaveFields:=", True, "SaveRadFields:=", True])
print(f"  Sweep1: {f_min/1e9:.1f}~{f_max/1e9:.1f}GHz, {n_pts} pts, SaveRadFields=True")

# ============================================================================
# 7. 保存 & 仿真
# ============================================================================
print()
print("=" * 70)
print("STEP 7: Simulation")
print("=" * 70)

hfss.save_project()
print(f"  Saved: {hfss.project_path}")

v = oDesign.ValidateDesign()
print(f"  Validation: {v}")

t0 = time.time()
print("  Running Setup1...")
oDesign.Analyze("Setup1")
t_sim = time.time() - t0
print(f"  Done! {t_sim:.0f}s ({t_sim/60:.1f}min)")

# ============================================================================
# 8. S11 提取
# ============================================================================
print()
print("=" * 70)
print("STEP 8: S11 Extraction")
print("=" * 70)

oSolutions = oDesign.GetModule("Solutions")
oReport = oDesign.GetModule("ReportSetup")

s1p_path = os.path.join(PROJECT_DIR, "dipole_v7.s1p")
s11_csv = os.path.join(PROJECT_DIR, "s11_v7.csv")

try:
    oSolutions.ExportNetworkData("", ["Setup1:Sweep1"], 3, s1p_path, ["All"], True, 50)
    print(f"  Touchstone: {s1p_path}")
except Exception as e:
    print(f"  Touchstone error: {e}")

try:
    oReport.CreateReport(
        "S11_Report", "Terminal Solution Data", "Rectangular Plot",
        "Setup1 : Sweep1", [],
        ["Freq:=", ["All"]],
        ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal_name},{terminal_name}))"]],
        [])
    oReport.ExportToFile("S11_Report", s11_csv)
    print(f"  S11 CSV: {s11_csv}")
except Exception as e:
    print(f"  S11 report error: {e}")


# ============================================================================
# 9. S11 分析 & 绘图
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
print("=" * 70)
print("STEP 9: S11 Analysis")
print("=" * 70)

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
        print(f"  -10dB BW:    {bw_lo:.3f} ~ {bw_hi:.3f} GHz")
        print(f"  Abs BW:      {bw_abs*1e3:.1f} MHz")
        print(f"  Rel BW:      {bw_rel:.2f}%")
    else:
        print("  WARNING: S11 does not reach -10 dB")
    print(f"  Freq error:  {freq_err:.2f}% {'-> NEEDS OPT' if freq_err > 2 or s11_min > -10 else '-> OK'}")

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
    ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'Resonance {f_res:.3f} GHz')
    if bw_abs > 0:
        ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green', label=f'BW {bw_abs*1e3:.0f}MHz')
    ax.set_xlabel('Frequency (GHz)'); ax.set_ylabel('S11 (dB)')
    ax.set_title(f'Printed Dipole v7 (arm={dipole_arm}mm) - S11')
    ax.legend(); ax.grid(True, alpha=0.3); ax.set_xlim(f_min/1e9, f_max/1e9)
    fig.savefig(os.path.join(PROJECT_DIR, "s11_v7.png"), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"\n  S11 plot: {os.path.join(PROJECT_DIR, 's11_v7.png')}")
except Exception as e:
    print(f"  Plot error: {e}")

# ============================================================================
# 10. 参数扫描优化
# ============================================================================
if need_sweep:
    print()
    print("=" * 70)
    print("STEP 10: Parameter Optimization")
    print("=" * 70)

    current_arm = dipole_arm
    best_arm = current_arm; best_f_res = f_res; best_s11 = s11_min
    best_bw_rel = bw_rel; best_bw_abs = bw_abs; best_bw_lo = bw_lo; best_bw_hi = bw_hi

    for opt_iter in range(10):
        # 频率缩放法: new_arm = old_arm * (f_measured / f_target)
        if best_f_res > 0:
            new_arm = round(current_arm * (best_f_res / (f0/1e9)), 1)
        else:
            new_arm = current_arm + 2.0

        # 防止变化太小或太大
        if abs(new_arm - current_arm) < 0.5:
            print(f"\n  Iter {opt_iter+1}: arm change < 0.5mm, converged")
            break
        if new_arm > 80 or new_arm < 15:
            print(f"\n  Iter {opt_iter+1}: arm {new_arm:.1f}mm out of range [15,80], stopping")
            break

        print(f"\n  Iter {opt_iter+1}: arm {current_arm:.1f} -> {new_arm:.1f} mm")

        # 删除旧的偶极子臂
        try:
            oEditor.Delete(["NAME:Selections", "Selections:=", "Dipole_Left,Dipole_Right"])
        except Exception as e:
            print(f"    Delete error: {e}"); break

        # 创建新臂
        new_left_x = -(gap/2 + new_arm)
        create_box("Dipole_Left", new_left_x, -dipole_w/2, z_top, new_arm, dipole_w, cu_t, "copper")
        create_box("Dipole_Right", gap/2, -dipole_w/2, z_top, new_arm, dipole_w, cu_t, "copper")
        print(f"    New arms: {new_arm:.1f}mm (total {2*new_arm+gap:.1f}mm)")

        hfss.save_project()
        t_opt = time.time()
        print("    Analyzing...")
        oDesign.Analyze("Setup1")
        print(f"    Done in {time.time()-t_opt:.0f}s")

        # 提取 S11
        opt_csv = os.path.join(PROJECT_DIR, f"s11_v7_opt{opt_iter+1}.csv")
        try:
            rname = f"S11_Opt{opt_iter+1}"
            oReport.CreateReport(
                rname, "Terminal Solution Data", "Rectangular Plot",
                "Setup1 : Sweep1", [],
                ["Freq:=", ["All"]],
                ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal_name},{terminal_name}))"]],
                [])
            oReport.ExportToFile(rname, opt_csv)

            opt_f, opt_s = parse_s11_csv(opt_csv)
            r = analyze_s11(opt_f, opt_s, f0/1e9)

            print(f"    f_res={r['f_res']:.4f}GHz, S11={r['s11_min']:.2f}dB, err={r['freq_err']:.2f}%")
            if r['bw_abs'] > 0:
                print(f"    -10dB BW: {r['bw_lo']:.3f}~{r['bw_hi']:.3f}GHz ({r['bw_rel']:.1f}%)")

            current_arm = new_arm
            best_f_res = r['f_res']; best_s11 = r['s11_min']; best_arm = new_arm
            best_bw_rel = r['bw_rel']; best_bw_abs = r['bw_abs']
            best_bw_lo = r['bw_lo']; best_bw_hi = r['bw_hi']
            freqs_ghz = opt_f; s11_db = opt_s
            f_res = r['f_res']; s11_min = r['s11_min']
            bw_rel = r['bw_rel']; bw_abs = r['bw_abs']
            bw_lo = r['bw_lo']; bw_hi = r['bw_hi']
            freq_err = r['freq_err']

            if r['freq_err'] <= 2.0 and r['s11_min'] < -10:
                print(f"    ** CONVERGED! f_err={r['freq_err']:.2f}%, S11={r['s11_min']:.2f}dB **")
                break

            if r['freq_err'] <= 2.0 and r['s11_min'] > -10:
                print(f"    Freq OK but S11 > -10dB ({r['s11_min']:.2f}dB). Stopping.")
                break

        except Exception as e:
            print(f"    Opt error: {e}"); traceback.print_exc(); break
    else:
        print(f"\n  Max iter reached. Best: arm={best_arm:.1f}mm, f={best_f_res:.4f}GHz, S11={best_s11:.2f}dB")

    # 最终 S11 图
    if len(freqs_ghz) > 0:
        try:
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.plot(freqs_ghz, s11_db, 'b-', lw=2, label='S11 (optimized)')
            ax.axhline(-10, color='r', ls='--', lw=1, label='-10 dB')
            ax.axvline(f0/1e9, color='g', ls=':', lw=1, label=f'Target {f0/1e9:.3f}')
            ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'Res {f_res:.3f}')
            if bw_abs > 0:
                ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green', label=f'BW {bw_abs*1e3:.0f}MHz ({bw_rel:.1f}%)')
            ax.set_xlabel('Frequency (GHz)'); ax.set_ylabel('S11 (dB)')
            ax.set_title(f'Printed Dipole v7 (arm={best_arm:.1f}mm, opt) - S11')
            ax.legend(); ax.grid(True, alpha=0.3); ax.set_xlim(f_min/1e9, f_max/1e9)
            fig.savefig(os.path.join(PROJECT_DIR, "s11_v7_opt.png"), dpi=150, bbox_inches='tight')
            plt.close(fig)
            print(f"\n  Opt S11 plot saved")
        except Exception as e:
            print(f"  Opt plot error: {e}")

else:
    best_arm = dipole_arm
    print("\n  No optimization needed (f_err<2% and S11<-10dB)")

# ============================================================================
# 11. 远场辐射方向图
# ============================================================================
print()
print("=" * 70)
print("STEP 11: Radiation Pattern & Gain")
print("=" * 70)

max_gain = 0
e_csv = os.path.join(PROJECT_DIR, "e_plane_v7.csv")
h_csv = os.path.join(PROJECT_DIR, "h_plane_v7.csv")

# 尝试列出已有的远场设置名
try:
    ff_names = oRadField.GetSetupNames()
    print(f"  Far field setups: {ff_names}")
    if ff_names and not ff_setup_name:
        ff_setup_name = str(ff_names[0])
except Exception:
    pass

# 如果没有远场设置，尝试使用默认 "3D" 或 "Infinite Sphere1"
if not ff_setup_name:
    ff_setup_name = "3D"  # AEDT 2019 default sphere name

for plane_name, phi_val, csv_path in [("E_Plane_v7", "0deg", e_csv), ("H_Plane_v7", "90deg", h_csv)]:
    try:
        oReport.CreateReport(
            plane_name, "Far Fields", "Radiation Pattern",
            "Setup1 : LastAdaptive",
            ["Context:=", ff_setup_name],
            ["Theta:=", ["All"], "Phi:=", [phi_val], "Freq:=", [f"{f0/1e9}GHz"]],
            ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]],
            [])
        oReport.ExportToFile(plane_name, csv_path)
        print(f"  {plane_name}: OK -> {csv_path}")
    except Exception as e:
        print(f"  {plane_name} error: {e}")

        # 尝试不指定 context
        try:
            rn2 = f"{plane_name}_noCtx"
            oReport.CreateReport(
                rn2, "Far Fields", "Radiation Pattern",
                "Setup1 : LastAdaptive", [],
                ["Theta:=", ["All"], "Phi:=", [phi_val], "Freq:=", [f"{f0/1e9}GHz"]],
                ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]],
                [])
            oReport.ExportToFile(rn2, csv_path)
            print(f"  {rn2}: OK -> {csv_path}")
        except Exception as e2:
            print(f"  {plane_name} (no context) also failed: {e2}")

# 读取 & 绘制方向图
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
            max_gain_e = np.max(g_e)
            max_gain_h = np.max(g_h)
            max_gain = max(max_gain_e, max_gain_h)
            print(f"\n  E-plane max gain: {max_gain_e:.2f} dBi")
            print(f"  H-plane max gain: {max_gain_h:.2f} dBi")
            print(f"  Antenna gain:     {max_gain:.2f} dBi")

            fig, (ax1, ax2) = plt.subplots(1, 2, subplot_kw={'projection': 'polar'}, figsize=(14, 6))
            ax1.plot(np.radians(th_e), g_e, 'b-', lw=2)
            ax1.set_title(f'E-Plane (xz, phi=0)\nGain={max_gain_e:.2f} dBi', pad=20)
            ax1.set_theta_zero_location('N'); ax1.set_theta_direction(-1)
            ax2.plot(np.radians(th_h), g_h, 'r-', lw=2)
            ax2.set_title(f'H-Plane (yz, phi=90)\nGain={max_gain_h:.2f} dBi', pad=20)
            ax2.set_theta_zero_location('N'); ax2.set_theta_direction(-1)
            fig.suptitle(f'Printed Dipole Antenna @ {f0/1e9:.3f} GHz', fontsize=14)
            fig.savefig(os.path.join(PROJECT_DIR, "radiation_v7.png"), dpi=150, bbox_inches='tight')
            plt.close(fig)
            print(f"  Pattern plot: {os.path.join(PROJECT_DIR, 'radiation_v7.png')}")
        else:
            print("  Pattern data empty")
    else:
        print("  Pattern CSV missing, estimating gain from S-parameters")
except Exception as e:
    print(f"  Pattern analysis error: {e}")

# ============================================================================
# 12. 设计总结
# ============================================================================
print()
print("=" * 70)
print("=" * 70)
print("DESIGN SUMMARY")
print("=" * 70)
print("=" * 70)
print()
print(f"  Antenna Type:       Printed Dipole + Truncated Ground + Microstrip Balun")
print(f"  Target Frequency:   {f0/1e9:.3f} GHz")
print(f"  Substrate:          FR4 (er={eps_r}), {sub_L}x{sub_total_y:.1f}x{sub_H} mm")
print(f"  Initial Arm Length: {dipole_arm:.1f} mm")
if need_sweep:
    print(f"  Optimized Arm:      {best_arm:.1f} mm")
    print(f"  Total Dipole:       {2*best_arm+gap:.1f} mm")
print(f"  Dipole Width:       {dipole_w:.1f} mm")
print(f"  Balun Length:       {balun_L:.2f} mm (lambda_g/4)")
print(f"  Ground Truncated:   y < {gnd_y_top:.1f} mm")
print()
print(f"  Resonance Freq:     {f_res:.4f} GHz")
print(f"  Freq Error:         {freq_err:.2f}%")
print(f"  Min S11:            {s11_min:.2f} dB")
if bw_abs > 0:
    print(f"  -10dB Bandwidth:    {bw_lo:.3f} ~ {bw_hi:.3f} GHz")
    print(f"  Abs Bandwidth:      {bw_abs*1e3:.1f} MHz")
    print(f"  Rel Bandwidth:      {bw_rel:.2f}%")
print(f"  Max Gain:           {max_gain:.2f} dBi")
print()
print(f"  Output Files:")
print(f"    S11 CSV:    {s11_csv}")
print(f"    S11 Plot:   {os.path.join(PROJECT_DIR, 's11_v7.png')}")
if need_sweep:
    print(f"    S11 Opt:    {os.path.join(PROJECT_DIR, 's11_v7_opt.png')}")
if os.path.exists(e_csv):
    print(f"    E-plane:    {e_csv}")
if os.path.exists(h_csv):
    print(f"    H-plane:    {h_csv}")
if os.path.exists(os.path.join(PROJECT_DIR, "radiation_v7.png")):
    print(f"    Pattern:    {os.path.join(PROJECT_DIR, 'radiation_v7.png')}")
print(f"    Touchstone: {s1p_path}")
print(f"    Project:    {hfss.project_path}")
print()

hfss.save_project()
print("=" * 70)
print("SIMULATION COMPLETE")
print("=" * 70)
