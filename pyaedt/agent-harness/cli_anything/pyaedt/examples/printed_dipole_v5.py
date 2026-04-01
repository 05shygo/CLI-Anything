#!/usr/bin/env python3
"""
printed_dipole_v5.py - 印刷偶极子天线（带微带巴伦馈线）HFSS 仿真
Printed Dipole Antenna with Microstrip Balun - HFSS Simulation

AEDT 2019.1 + PyAEDT 0.8.11 - DrivenTerminal + WavePort (AutoIdentifyPorts)

设计参数:
  - 天线类型: 印刷偶极子 + 微带巴伦
  - 中心频率: 2.217 GHz
  - 介质板: FR4 (εr=4.4, 厚度1.6mm)
  - 偶极子臂长初始值: 50 mm
  - 频率扫描: 1.5 ~ 3.0 GHz, 步进 0.01 GHz

端口方法: AutoIdentifyPorts 在 AirBox y_min 面自动识别 WavePort
"""

import numpy as np
import os
import sys
import time
import csv
import math

# ============================================================================
# 0. 环境配置
# ============================================================================
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'

PROJECT_DIR = r"D:\class_design"
os.makedirs(PROJECT_DIR, exist_ok=True)
LOG_FILE = os.path.join(PROJECT_DIR, "dipole_v5.log")


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

# PyAEDT 兼容性补丁 (AEDT 2019.1)
try:
    from pyaedt import desktop as _desktop_mod
    _orig_init = _desktop_mod.Desktop.__init__
    def _patched_init(self, *a, **kw):
        _orig_init(self, *a, **kw)
        if not hasattr(self, 'student_version'):
            self.student_version = False
    _desktop_mod.Desktop.__init__ = _patched_init
except Exception:
    pass

try:
    import pyaedt.application.Design as _design_mod
    _orig_ds = _design_mod.DesignSettings.__init__
    def _patched_ds(self, app):
        try:
            _orig_ds(self, app)
        except AttributeError:
            self._app = app
            self.design_settings = None
            self.manipulate_inputs = None
    _design_mod.DesignSettings.__init__ = _patched_ds
except Exception:
    pass

from pyaedt import Hfss

# ============================================================================
# 1. 设计参数
# ============================================================================
print("=" * 70)
print("STEP 1: Design Parameters")
print("=" * 70)

c_light = 3e8          # m/s
f0 = 2.217e9           # Hz - 中心频率
eps_r = 4.4            # FR4 介电常数

lambda_0 = c_light / f0                        # 自由空间波长 (m)
lambda_g = lambda_0 / np.sqrt(eps_r)           # 介质中波长 (m)

# 基板参数 (mm)
sub_L = 120.0      # X 方向
sub_W = 60.0       # Y 方向 (原始设计宽度)
sub_H = 1.6        # Z 方向 (厚度)
cu_t = 0.035       # 铜层厚度

# 偶极子参数 (mm)
dipole_arm = 50.0   # 臂长 (单侧)
dipole_w = 3.0      # 臂宽
gap = 1.0           # 两臂间隙

# 巴伦参数 (mm)
balun_L = lambda_g * 1e3 / 4.0    # λg/4 ≈ 16.1mm
balun_w_start = 1.0               # 窄端 (馈线侧)
balun_w_end = 3.0                 # 宽端 (偶极子侧)
balun_segs = 10                   # 渐变段数

# 频率扫描
f_min = 1.5e9
f_max = 3.0e9
f_step = 0.01e9

# AirBox 填充 (λ/4)
pad = lambda_0 * 1e3 / 4.0   # mm

# ---- 关键坐标 ----
z_top = sub_H                                # 基板顶面 Z=1.6
balun_y_start = -dipole_w / 2                 # 巴伦起始 Y (与偶极子连接处)
balun_y_end = balun_y_start - balun_L         # 巴伦结束 Y
airbox_ymin = -(sub_W / 2 + pad)              # AirBox Y下界

# 基板/地板延伸到 AirBox 下界 (WavePort 需要)
sub_y_start = airbox_ymin                     # 基板 Y 起点
sub_y_end = sub_W / 2                         # 基板 Y 终点 (+30mm)
sub_total_y = sub_y_end - sub_y_start         # 基板总 Y 尺寸

# 馈线从巴伦末端延伸到 AirBox 下界
feed_y_start = airbox_ymin
feed_y_end = balun_y_end
feed_len = feed_y_end - feed_y_start          # 馈线长度

# AirBox 尺寸
air_x_half = sub_L / 2 + pad
air_y_half = sub_W / 2 + pad
air_z_top = pad
air_z_bot = pad

n_pts = int((f_max - f_min) / f_step) + 1    # 频率点数 (151)

print(f"  f0           = {f0/1e9:.3f} GHz")
print(f"  lambda_0     = {lambda_0*1e3:.2f} mm")
print(f"  lambda_g     = {lambda_g*1e3:.2f} mm")
print(f"  balun_L      = {balun_L:.2f} mm (lambda_g/4)")
print(f"  dipole_arm   = {dipole_arm:.1f} mm")
print(f"  pad          = {pad:.1f} mm (lambda_0/4)")
print(f"  airbox_ymin  = {airbox_ymin:.1f} mm")
print(f"  sub extent Y = {sub_y_start:.1f} to {sub_y_end:.1f} mm ({sub_total_y:.1f} mm)")
print(f"  feed extent  = {feed_y_start:.1f} to {feed_y_end:.1f} mm ({feed_len:.1f} mm)")
print(f"  freq sweep   = {f_min/1e9:.1f} ~ {f_max/1e9:.1f} GHz, {n_pts} pts")
print()

# ============================================================================
# 2. 启动 AEDT
# ============================================================================
print("=" * 70)
print("STEP 2: Launch AEDT")
print("=" * 70)

PROJECT_NAME = "Printed_Dipole_v5"

hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, PROJECT_NAME),
    designname="Dipole_Balun",
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


# ============================================================================
# COM 辅助函数 (AEDT 2019.1 格式)
# ============================================================================
def create_box(name, x, y, z, dx, dy, dz, mat, solve_inside=None):
    """创建 3D 实体 box"""
    if solve_inside is None:
        solve_inside = mat.lower() not in ("copper", "pec", "aluminum")
    oEditor.CreateBox(
        ["NAME:BoxParameters",
         "XPosition:=", f"{x}mm",
         "YPosition:=", f"{y}mm",
         "ZPosition:=", f"{z}mm",
         "XSize:=", f"{dx}mm",
         "YSize:=", f"{dy}mm",
         "ZSize:=", f"{dz}mm"],
        ["NAME:Attributes",
         "Name:=", name,
         "Flags:=", "",
         "Color:=", "(143 175 131)",
         "Transparency:=", 0,
         "PartCoordinateSystem:=", "Global",
         "UDMId:=", "",
         "MaterialValue:=", f'"{mat}"',
         "SurfaceMaterialValue:=", '""',
         "SolveInside:=", solve_inside,
         "IsMaterialEditable:=", True,
         "UseMaterialAppearance:=", False,
         "IsLightweight:=", False])


# ============================================================================
# 3. 创建天线几何模型
# ============================================================================
print()
print("=" * 70)
print("STEP 3: Antenna Geometry")
print("=" * 70)

# 3.1 FR4 基板 (延伸到 AirBox 边界)
print(f"  [3.1] Substrate: {sub_L} x {sub_total_y:.1f} x {sub_H} mm")
create_box("Substrate",
           -sub_L / 2, sub_y_start, 0,
           sub_L, sub_total_y, sub_H,
           "FR4_epoxy")

# 3.2 地板 (延伸到 AirBox 边界)
print(f"  [3.2] Ground:    {sub_L} x {sub_total_y:.1f} x {cu_t} mm")
create_box("Ground",
           -sub_L / 2, sub_y_start, -cu_t,
           sub_L, sub_total_y, cu_t,
           "copper")

# 3.3 偶极子臂 (基板顶面)
left_x_start = -(gap / 2 + dipole_arm)
right_x_start = gap / 2
print(f"  [3.3] Dipole Left:  X=[{left_x_start:.1f}, {-gap/2:.1f}]")
create_box("Dipole_Left",
           left_x_start, -dipole_w / 2, z_top,
           dipole_arm, dipole_w, cu_t,
           "copper")
print(f"        Dipole Right: X=[{right_x_start:.1f}, {right_x_start + dipole_arm:.1f}]")
create_box("Dipole_Right",
           right_x_start, -dipole_w / 2, z_top,
           dipole_arm, dipole_w, cu_t,
           "copper")

# 3.4 微带巴伦 (渐变段)
print(f"  [3.4] Balun: {balun_segs} segs, L={balun_L:.2f}mm, taper {balun_w_end}->{balun_w_start}mm")
for i in range(balun_segs):
    frac_s = i / balun_segs
    frac_e = (i + 1) / balun_segs
    w_s = balun_w_end - (balun_w_end - balun_w_start) * frac_s
    w_e = balun_w_end - (balun_w_end - balun_w_start) * frac_e
    avg_w = (w_s + w_e) / 2.0
    seg_len = balun_L / balun_segs
    seg_y = balun_y_start - seg_len * (i + 1)
    create_box(f"Balun_{i}",
               -avg_w / 2, seg_y, z_top,
               avg_w, seg_len, cu_t,
               "copper")

# 3.5 馈线 (从巴伦末端延伸到 AirBox 下界)
print(f"  [3.5] Feed Line: w={balun_w_start}mm, Y=[{feed_y_start:.1f}, {feed_y_end:.2f}]")
create_box("Feed_Line",
           -balun_w_start / 2, feed_y_start, z_top,
           balun_w_start, feed_len, cu_t,
           "copper")

# 3.6 AirBox + 辐射边界
air_x_size = 2 * air_x_half
air_y_size = 2 * air_y_half
air_z_size = sub_H + air_z_top + air_z_bot

print(f"  [3.6] AirBox: {air_x_size:.1f} x {air_y_size:.1f} x {air_z_size:.1f} mm")
create_box("Air",
           -air_x_half, airbox_ymin, -air_z_bot,
           air_x_size, air_y_size, air_z_size,
           "vacuum", solve_inside=True)

oBnd.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["Air"],
     "IsFssReference:=", False,
     "IsForPML:=", False])
print("        Radiation boundary assigned")

# ============================================================================
# 4. WavePort (AutoIdentifyPorts)
# ============================================================================
print()
print("=" * 70)
print("STEP 4: WavePort Setup")
print("=" * 70)

air_faces = oEditor.GetFaceIDs("Air")
print(f"  AirBox faces: {air_faces}")

# Face index 2 = y_min face (verified in AEDT 2019.1 box face ordering)
ymin_face_id = int(air_faces[2])
print(f"  Using y_min face ID: {ymin_face_id}")

oBnd.AutoIdentifyPorts(
    ["NAME:Faces", ymin_face_id],
    True,
    ["NAME:ReferenceConductors", "Ground"],
    "Port1",
    True)
print("  AutoIdentifyPorts OK")

# 获取终端名称
terms = oBnd.GetExcitationsOfType("Terminal")
if terms:
    terminal = str(terms[0])
else:
    terminal = "Feed_Line_T1"
print(f"  Terminal: {terminal}")

# ============================================================================
# 5. 无限球面定义 (远场计算)
# ============================================================================
print()
print("STEP 5: Infinite Sphere Definition")
oRadField = oDesign.GetModule("RadField")
try:
    oRadField.InsertInfiniteSphereDef(
        ["NAME:InfSphere1",
         "UseCustomRadiationSurface:=", False,
         "ThetaStart:=", "-180deg",
         "ThetaStop:=", "180deg",
         "ThetaStep:=", "2deg",
         "PhiStart:=", "0deg",
         "PhiStop:=", "360deg",
         "PhiStep:=", "2deg",
         "UseLocalCS:=", False])
    print("  InfSphere1 defined (theta -180~180, phi 0~360, step 2 deg)")
except Exception as e:
    print(f"  InfSphere note: {e}")

# ============================================================================
# 6. 分析设置 + 频率扫描
# ============================================================================
print()
print("=" * 70)
print("STEP 6: Analysis Setup")
print("=" * 70)

oAnalysis = oDesign.GetModule("AnalysisSetup")

oAnalysis.InsertSetup("HfssDriven",
    ["NAME:Setup1",
     "Frequency:=", f"{f0/1e9}GHz",
     "MaxDeltaS:=", 0.02,
     "MaximumPasses:=", 12,
     "MinimumPasses:=", 2,
     "MinimumConvergedPasses:=", 2,
     "PercentRefinement:=", 30,
     "IsEnabled:=", True,
     "BasisOrder:=", 1,
     "UseIterativeSolver:=", False,
     "DoLambdaRefine:=", True,
     "DoMaterialLambdaRefine:=", True,
     "SetLambdaTarget:=", False,
     "Target:=", 0.3333])
print(f"  Setup1: f0={f0/1e9:.3f}GHz, MaxDeltaS=0.02, MaxPasses=12")

oAnalysis.InsertFrequencySweep("Setup1",
    ["NAME:Sweep1",
     "IsEnabled:=", True,
     "SetupType:=", "LinearCount",
     "StartValue:=", f"{f_min/1e9}GHz",
     "StopValue:=", f"{f_max/1e9}GHz",
     "Count:=", n_pts,
     "Type:=", "Discrete",
     "SaveFields:=", True,
     "SaveRadFields:=", False])
print(f"  Sweep1: {f_min/1e9:.1f}~{f_max/1e9:.1f}GHz, {n_pts} pts (Discrete)")

# ============================================================================
# 7. 保存 & 仿真
# ============================================================================
print()
print("=" * 70)
print("STEP 7: Simulation")
print("=" * 70)

hfss.save_project()
print(f"  Project saved: {hfss.project_path}")

v = oDesign.ValidateDesign()
print(f"  Validation: {v}")

t_start = time.time()
print("  Running Setup1... (please wait)")
oDesign.Analyze("Setup1")
t_sim = time.time() - t_start
print(f"  Simulation complete! Time: {t_sim:.1f}s ({t_sim/60:.1f}min)")

# ============================================================================
# 8. 提取 S11
# ============================================================================
print()
print("=" * 70)
print("STEP 8: S11 Extraction")
print("=" * 70)

oSolutions = oDesign.GetModule("Solutions")
oReport = oDesign.GetModule("ReportSetup")

# 8.1 Touchstone 导出
s1p_path = os.path.join(PROJECT_DIR, "dipole_v5.s1p")
try:
    oSolutions.ExportNetworkData(
        "", ["Setup1:Sweep1"], 3, s1p_path, ["All"], True, 50)
    print(f"  Touchstone: {s1p_path}")
except Exception as e:
    print(f"  Touchstone export error: {e}")

# 8.2 S11 CSV 导出
s11_csv = os.path.join(PROJECT_DIR, "s11_v5.csv")
try:
    oReport.CreateReport(
        "S11_Report", "Terminal Solution Data", "Rectangular Plot",
        "Setup1 : Sweep1",
        [],
        ["Freq:=", ["All"]],
        ["X Component:=", "Freq",
         "Y Component:=", [f"dB(St({terminal},{terminal}))"]],
        [])
    oReport.ExportToFile("S11_Report", s11_csv)
    print(f"  S11 CSV: {s11_csv}")
except Exception as e:
    print(f"  S11 report error: {e}")

# ============================================================================
# 9. S11 分析
# ============================================================================
print()
print("=" * 70)
print("STEP 9: S11 Analysis")
print("=" * 70)

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

    print(f"  Resonance:   {f_res:.4f} GHz")
    print(f"  S11 min:     {s11_min:.2f} dB")

    below_10 = np.where(s11_db < -10)[0]
    if len(below_10) >= 2:
        bw_lo = freqs_ghz[below_10[0]]
        bw_hi = freqs_ghz[below_10[-1]]
        bw_abs = bw_hi - bw_lo
        bw_rel = bw_abs / f_res * 100
        print(f"  -10dB BW:    {bw_lo:.3f} ~ {bw_hi:.3f} GHz")
        print(f"  Abs BW:      {bw_abs*1e3:.1f} MHz")
        print(f"  Rel BW:      {bw_rel:.2f}%")
    else:
        bw_rel = 0
        bw_lo = bw_hi = f_res
        bw_abs = 0
        print("  WARNING: S11 does not reach -10 dB")

    freq_err = abs(f_res - f0 / 1e9) / (f0 / 1e9) * 100
    need_sweep = freq_err > 2.0
    print(f"  Freq error:  {freq_err:.2f}% {'-> NEEDS OPTIMIZATION' if need_sweep else '-> OK'}")

except Exception as e:
    print(f"  S11 parse error: {e}")
    import traceback
    traceback.print_exc()
    need_sweep = True
    f_res = f0 / 1e9
    s11_min = 0
    bw_rel = 0
    bw_abs = 0
    bw_lo = bw_hi = f_res
    freq_err = 999

# ============================================================================
# 10. S11 绘图
# ============================================================================
print()
print("STEP 10: S11 Plot")
try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(freqs_ghz, s11_db, 'b-', lw=2, label='S11')
    ax.axhline(-10, color='r', ls='--', lw=1, label='-10 dB')
    ax.axvline(f0 / 1e9, color='g', ls=':', lw=1, label=f'Target {f0/1e9:.3f} GHz')
    ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'Resonance {f_res:.3f} GHz')
    if bw_abs > 0:
        ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green',
                   label=f'BW {bw_abs*1e3:.0f} MHz ({bw_rel:.1f}%)')
    ax.set_xlabel('Frequency (GHz)', fontsize=13)
    ax.set_ylabel('S11 (dB)', fontsize=13)
    ax.set_title('Printed Dipole Antenna - S11', fontsize=14)
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3)
    ax.set_xlim(f_min / 1e9, f_max / 1e9)
    s11_img = os.path.join(PROJECT_DIR, "s11_v5.png")
    fig.savefig(s11_img, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  S11 plot: {s11_img}")
except Exception as e:
    print(f"  S11 plot error: {e}")

# ============================================================================
# 11. 远场辐射方向图
# ============================================================================
print()
print("=" * 70)
print("STEP 11: Radiation Pattern & Gain")
print("=" * 70)

e_csv = os.path.join(PROJECT_DIR, "e_plane_v5.csv")
h_csv = os.path.join(PROJECT_DIR, "h_plane_v5.csv")
max_gain = 0

# E-面 (phi=0, xz 面)
try:
    oReport.CreateReport(
        "E_Plane", "Far Fields", "Radiation Pattern",
        "Setup1 : LastAdaptive",
        ["Context:=", "InfSphere1"],
        ["Theta:=", ["All"], "Phi:=", ["0deg"], "Freq:=", [f"{f0/1e9}GHz"]],
        ["X Component:=", "Theta",
         "Y Component:=", ["GainTotal"]],
        [])
    oReport.ExportToFile("E_Plane", e_csv)
    print(f"  E-plane CSV: {e_csv}")
except Exception as e:
    print(f"  E-plane error: {e}")
    e_csv = None

# H-面 (phi=90, yz 面)
try:
    oReport.CreateReport(
        "H_Plane", "Far Fields", "Radiation Pattern",
        "Setup1 : LastAdaptive",
        ["Context:=", "InfSphere1"],
        ["Theta:=", ["All"], "Phi:=", ["90deg"], "Freq:=", [f"{f0/1e9}GHz"]],
        ["X Component:=", "Theta",
         "Y Component:=", ["GainTotal"]],
        [])
    oReport.ExportToFile("H_Plane", h_csv)
    print(f"  H-plane CSV: {h_csv}")
except Exception as e:
    print(f"  H-plane error: {e}")
    h_csv = None


def read_csv_pattern(path):
    """解析辐射方向图 CSV"""
    th, g = [], []
    with open(path, 'r') as f:
        reader = csv.reader(f)
        next(reader)  # skip header
        for row in reader:
            if len(row) >= 2:
                try:
                    th.append(float(row[0]))
                    g.append(float(row[1]))
                except ValueError:
                    continue
    return np.array(th), np.array(g)


try:
    if e_csv and os.path.exists(e_csv) and h_csv and os.path.exists(h_csv):
        th_e, g_e = read_csv_pattern(e_csv)
        th_h, g_h = read_csv_pattern(h_csv)

        max_gain_e = np.max(g_e) if len(g_e) > 0 else -999
        max_gain_h = np.max(g_h) if len(g_h) > 0 else -999
        max_gain = max(max_gain_e, max_gain_h)

        print(f"\n  -------- Radiation Results --------")
        print(f"  E-plane max gain: {max_gain_e:.2f} dBi")
        print(f"  H-plane max gain: {max_gain_h:.2f} dBi")
        print(f"  Antenna gain:     {max_gain:.2f} dBi")

        # 绘制辐射方向图
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

        pat_img = os.path.join(PROJECT_DIR, "radiation_pattern_v5.png")
        fig.savefig(pat_img, dpi=150, bbox_inches='tight')
        plt.close(fig)
        print(f"  Pattern plot: {pat_img}")
    else:
        print("  Skipped pattern analysis (CSV missing)")
except Exception as e:
    print(f"  Pattern analysis error: {e}")

# ============================================================================
# 12. 参数优化 (谐振频率偏差 > 2%)
# ============================================================================
if need_sweep:
    print()
    print("=" * 70)
    print("STEP 12: Parameter Optimization")
    print("=" * 70)

    current_arm = dipole_arm
    best_arm = current_arm
    best_f_res = f_res
    best_s11 = s11_min

    for opt_iter in range(5):
        # 根据频率比调整臂长: arm_new = arm_old * (f_res / f_target)
        new_arm = current_arm * (best_f_res / (f0 / 1e9))
        new_arm = round(new_arm, 1)

        if abs(new_arm - current_arm) < 0.5:
            print(f"  Iteration {opt_iter+1}: arm change < 0.5mm, stopping")
            break

        print(f"\n  Iteration {opt_iter+1}: arm {current_arm:.1f} -> {new_arm:.1f} mm")

        # 删除旧偶极子臂
        try:
            oEditor.Delete(
                ["NAME:Selections",
                 "Selections:=", "Dipole_Left,Dipole_Right"])
            print("    Deleted old dipole arms")
        except Exception as e:
            print(f"    Delete error: {e}")
            break

        # 创建新偶极子臂
        new_left_x = -(gap / 2 + new_arm)
        create_box("Dipole_Left",
                   new_left_x, -dipole_w / 2, z_top,
                   new_arm, dipole_w, cu_t,
                   "copper")
        create_box("Dipole_Right",
                   gap / 2, -dipole_w / 2, z_top,
                   new_arm, dipole_w, cu_t,
                   "copper")
        print(f"    Created new arms: L={new_arm:.1f}mm")

        # 重新仿真
        hfss.save_project()
        t_opt = time.time()
        print("    Re-analyzing...")
        oDesign.Analyze("Setup1")
        print(f"    Done in {time.time()-t_opt:.1f}s")

        # 重新提取 S11
        opt_csv = os.path.join(PROJECT_DIR, f"s11_opt_{opt_iter+1}.csv")
        try:
            rname = f"S11_Opt{opt_iter+1}"
            oReport.CreateReport(
                rname, "Terminal Solution Data", "Rectangular Plot",
                "Setup1 : Sweep1",
                [],
                ["Freq:=", ["All"]],
                ["X Component:=", "Freq",
                 "Y Component:=", [f"dB(St({terminal},{terminal}))"]],
                [])
            oReport.ExportToFile(rname, opt_csv)

            opt_f, opt_s = [], []
            with open(opt_csv, 'r') as f:
                rd = csv.reader(f)
                next(rd)
                for row in rd:
                    if len(row) >= 2:
                        try:
                            opt_f.append(float(row[0]))
                            opt_s.append(float(row[1]))
                        except ValueError:
                            continue

            opt_f = np.array(opt_f)
            opt_s = np.array(opt_s)
            idx = np.argmin(opt_s)

            iter_f_res = opt_f[idx]
            iter_s11 = opt_s[idx]
            iter_err = abs(iter_f_res - f0 / 1e9) / (f0 / 1e9) * 100

            print(f"    f_res={iter_f_res:.4f}GHz, S11={iter_s11:.2f}dB, err={iter_err:.2f}%")

            current_arm = new_arm
            best_f_res = iter_f_res
            best_s11 = iter_s11
            best_arm = new_arm

            if iter_err <= 2.0:
                print(f"    Frequency error < 2%, optimization converged!")
                freq_err = iter_err
                f_res = iter_f_res
                s11_min = iter_s11

                # 更新带宽
                below_10_opt = np.where(opt_s < -10)[0]
                if len(below_10_opt) >= 2:
                    bw_lo = opt_f[below_10_opt[0]]
                    bw_hi = opt_f[below_10_opt[-1]]
                    bw_abs = bw_hi - bw_lo
                    bw_rel = bw_abs / iter_f_res * 100

                # 更新 S11 数据 (用于最终绘图)
                freqs_ghz = opt_f
                s11_db = opt_s
                break

        except Exception as e:
            print(f"    Optimization error: {e}")
            import traceback
            traceback.print_exc()
            break
    else:
        print(f"\n  Optimization did not converge in 5 iterations")
        print(f"  Best: arm={best_arm:.1f}mm, f_res={best_f_res:.4f}GHz")
        f_res = best_f_res
        s11_min = best_s11

    # 重新绘制优化后的 S11
    if len(freqs_ghz) > 0:
        try:
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.plot(freqs_ghz, s11_db, 'b-', lw=2, label='S11 (optimized)')
            ax.axhline(-10, color='r', ls='--', lw=1, label='-10 dB')
            ax.axvline(f0 / 1e9, color='g', ls=':', lw=1, label=f'Target {f0/1e9:.3f} GHz')
            ax.axvline(f_res, color='orange', ls='-.', lw=1,
                       label=f'Resonance {f_res:.3f} GHz')
            if bw_abs > 0:
                ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green',
                           label=f'BW {bw_abs*1e3:.0f} MHz ({bw_rel:.1f}%)')
            ax.set_xlabel('Frequency (GHz)', fontsize=13)
            ax.set_ylabel('S11 (dB)', fontsize=13)
            ax.set_title('Printed Dipole Antenna - S11 (Optimized)', fontsize=14)
            ax.legend(fontsize=10)
            ax.grid(True, alpha=0.3)
            ax.set_xlim(f_min / 1e9, f_max / 1e9)
            s11_opt_img = os.path.join(PROJECT_DIR, "s11_v5_optimized.png")
            fig.savefig(s11_opt_img, dpi=150, bbox_inches='tight')
            plt.close(fig)
            print(f"  Optimized S11 plot: {s11_opt_img}")
        except Exception as e:
            print(f"  Plot error: {e}")

    # 重新提取优化后的远场
    try:
        # 先删除旧报告
        try:
            oReport.DeleteReports(["E_Plane", "H_Plane"])
        except Exception:
            pass

        oReport.CreateReport(
            "E_Plane_Opt", "Far Fields", "Radiation Pattern",
            "Setup1 : LastAdaptive",
            ["Context:=", "InfSphere1"],
            ["Theta:=", ["All"], "Phi:=", ["0deg"], "Freq:=", [f"{f0/1e9}GHz"]],
            ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]],
            [])
        e_csv_opt = os.path.join(PROJECT_DIR, "e_plane_v5_opt.csv")
        oReport.ExportToFile("E_Plane_Opt", e_csv_opt)

        oReport.CreateReport(
            "H_Plane_Opt", "Far Fields", "Radiation Pattern",
            "Setup1 : LastAdaptive",
            ["Context:=", "InfSphere1"],
            ["Theta:=", ["All"], "Phi:=", ["90deg"], "Freq:=", [f"{f0/1e9}GHz"]],
            ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]],
            [])
        h_csv_opt = os.path.join(PROJECT_DIR, "h_plane_v5_opt.csv")
        oReport.ExportToFile("H_Plane_Opt", h_csv_opt)

        if os.path.exists(e_csv_opt) and os.path.exists(h_csv_opt):
            th_e, g_e = read_csv_pattern(e_csv_opt)
            th_h, g_h = read_csv_pattern(h_csv_opt)
            max_gain_e = np.max(g_e) if len(g_e) > 0 else -999
            max_gain_h = np.max(g_h) if len(g_h) > 0 else -999
            max_gain = max(max_gain_e, max_gain_h)
            print(f"  Optimized E-plane gain: {max_gain_e:.2f} dBi")
            print(f"  Optimized H-plane gain: {max_gain_h:.2f} dBi")
            print(f"  Optimized antenna gain: {max_gain:.2f} dBi")

            fig, (ax1, ax2) = plt.subplots(1, 2, subplot_kw={'projection': 'polar'},
                                            figsize=(14, 6))
            ax1.plot(np.radians(th_e), g_e, 'b-', lw=2)
            ax1.set_title(f'E-Plane (Opt)\nGain={max_gain_e:.2f} dBi', pad=20)
            ax1.set_theta_zero_location('N')
            ax1.set_theta_direction(-1)
            ax2.plot(np.radians(th_h), g_h, 'r-', lw=2)
            ax2.set_title(f'H-Plane (Opt)\nGain={max_gain_h:.2f} dBi', pad=20)
            ax2.set_theta_zero_location('N')
            ax2.set_theta_direction(-1)
            pat_opt_img = os.path.join(PROJECT_DIR, "radiation_pattern_v5_opt.png")
            fig.savefig(pat_opt_img, dpi=150, bbox_inches='tight')
            plt.close(fig)
            print(f"  Optimized pattern plot: {pat_opt_img}")
    except Exception as e:
        print(f"  Optimized pattern error: {e}")
else:
    print("\n  Frequency error < 2%, no optimization needed")
    best_arm = dipole_arm

# ============================================================================
# 13. 设计总结
# ============================================================================
print()
print("=" * 70)
print("DESIGN SUMMARY")
print("=" * 70)
print(f"  Antenna:        Printed Dipole + Microstrip Balun")
print(f"  Target freq:    {f0/1e9:.3f} GHz")
print(f"  Substrate:      FR4 (er={eps_r}), {sub_L}x{sub_W}x{sub_H} mm")
print(f"  Dipole arm:     {dipole_arm:.1f} mm (initial)")
if need_sweep:
    print(f"  Optimized arm:  {best_arm:.1f} mm")
print(f"  Dipole width:   {dipole_w:.1f} mm")
print(f"  Balun length:   {balun_L:.2f} mm (lambda_g/4)")
print(f"  Gap:            {gap:.1f} mm")
print(f"  Resonance:      {f_res:.4f} GHz")
print(f"  S11 min:        {s11_min:.2f} dB")
if bw_abs > 0:
    print(f"  -10dB BW:       {bw_lo:.3f} ~ {bw_hi:.3f} GHz")
    print(f"  Rel bandwidth:  {bw_rel:.2f}%")
print(f"  Max gain:       {max_gain:.2f} dBi")
print(f"  Freq error:     {freq_err:.2f}%")
print()
print("  Output files:")
print(f"    Log:          {LOG_FILE}")
print(f"    Touchstone:   {s1p_path}")
print(f"    S11 CSV:      {s11_csv}")
print(f"    S11 Plot:     {os.path.join(PROJECT_DIR, 's11_v5.png')}")
if e_csv:
    print(f"    E-plane CSV:  {e_csv}")
if h_csv:
    print(f"    H-plane CSV:  {h_csv}")
print(f"    Pattern Plot: {os.path.join(PROJECT_DIR, 'radiation_pattern_v5.png')}")

hfss.save_project()
print()
print("=" * 70)
print("SIMULATION COMPLETE")
print("=" * 70)
