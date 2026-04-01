#!/usr/bin/env python3
"""
printed_dipole_v6.py - 印刷偶极子天线（截断地板 + 微带巴伦）
Printed Dipole Antenna with Truncated Ground Plane & Microstrip Balun

关键设计改进 (vs v5):
  - 地板截断：地板不延伸到偶极子下方，只在巴伦/馈线区域
  - 这使天线真正工作在偶极子模式（平衡辐射），而非微带传输线模式
  - 巴伦过渡区：地板在巴伦-偶极子交界处截止，
    实现微带（不平衡）到 CPS（平衡）的自然过渡

天线参数:
  - 中心频率: 2.217 GHz
  - 介质板: FR4 (εr=4.4, h=1.6mm)
  - 偶极子臂长: 从 λeff/4 开始（~30.6mm 每臂），可优化
  - 巴伦: 渐变微带，λg/4 ≈ 16.1mm
  - WavePort: AutoIdentifyPorts at AirBox y_min 面

端口方法: DrivenTerminal + AutoIdentifyPorts WavePort
"""

import numpy as np
import os, sys, time, csv

# ============================================================================
# 0. 环境配置
# ============================================================================
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'

PROJECT_DIR = r"D:\class_design"
os.makedirs(PROJECT_DIR, exist_ok=True)
LOG_FILE = os.path.join(PROJECT_DIR, "dipole_v6.log")


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

# PyAEDT 兼容性补丁
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
lambda_0 = c / f0                    # ~135.3 mm
lambda_g = lambda_0 / np.sqrt(eps_r) # ~64.5 mm
# 印刷偶极子有效波长（考虑介质板效应, εeff ≈ (εr+1)/2 = 2.7）
eps_eff = (eps_r + 1) / 2.0  # 微带有效介电常数
lambda_eff = lambda_0 / np.sqrt(eps_eff) * 1e3  # mm

# 基板 (mm)
sub_L = 120.0; sub_W = 80.0; sub_H = 1.6; cu_t = 0.035

# 偶极子 (mm) - 每臂 ≈ λeff/4 ≈ 30.6mm (半波偶极子总长 ≈ λeff/2)
dipole_arm = 30.0   # 初始臂长（每侧），会优化
dipole_w   = 3.0    # 臂宽
gap        = 1.0    # 两臂间隙

# 巴伦 (mm)
balun_L      = lambda_g * 1e3 / 4.0  # ≈16.1mm
balun_w_start = 3.0                   # 宽端（微带线，匹配50Ω）
balun_w_end   = 3.0                   # 窄端（匹配偶极子臂宽）
balun_segs    = 8

# 馈线宽度（50Ω微带线在1.6mm FR4上 ≈ 3.06mm）
feed_w = 3.0  # mm

# 频率扫描
f_min = 1.5e9; f_max = 3.0e9; f_step = 0.01e9

# AirBox padding
pad = lambda_0 * 1e3 / 4.0  # ≈33.8mm

# ==== 关键坐标 ====
z_top = sub_H  # 基板顶面

# 偶极子位于 y=0 平面
# 巴伦从偶极子中心向 -y 延伸
balun_y_top = -dipole_w / 2            # 巴伦顶端 (与偶极子相连)
balun_y_bot = balun_y_top - balun_L    # 巴伦底端

# AirBox 范围
airbox_ymin = -(sub_W / 2 + pad)
airbox_ymax = sub_W / 2 + pad
airbox_xmin = -(sub_L / 2 + pad)
airbox_xmax = sub_L / 2 + pad
air_z_bot = pad
air_z_top = pad

# 基板 Y 范围 (延伸到 AirBox 下界以支持 WavePort)
sub_y_start = airbox_ymin  # 延伸到 AirBox 边界
sub_y_end = sub_W / 2
sub_total_y = sub_y_end - sub_y_start

# ==== 关键设计：截断地板 ====
# 地板只覆盖 y < gnd_truncate_y 的区域 (巴伦以下)
# 不覆盖偶极子区域，使天线工作在真正的偶极子模式
gnd_truncate_y = balun_y_top  # 地板上边界 = 巴伦顶端 (偶极子中心线下方)

# 馈线从 AirBox 边界延伸到巴伦底端
feed_y_start = airbox_ymin
feed_y_end = balun_y_bot
feed_len = feed_y_end - feed_y_start

n_pts = int((f_max - f_min) / f_step) + 1

print(f"  f0              = {f0/1e9:.3f} GHz")
print(f"  lambda_0        = {lambda_0*1e3:.2f} mm")
print(f"  lambda_g (FR4)  = {lambda_g*1e3:.2f} mm")
print(f"  lambda_eff      = {lambda_eff:.2f} mm")
print(f"  eps_eff         = {eps_eff:.2f}")
print(f"  dipole_arm      = {dipole_arm:.1f} mm (each side)")
print(f"  total dipole    = {2*dipole_arm + gap:.1f} mm")
print(f"  balun_L         = {balun_L:.2f} mm (lambda_g/4)")
print(f"  pad             = {pad:.1f} mm (lambda_0/4)")
print(f"  ground truncate = y < {gnd_truncate_y:.1f} mm")
print(f"  feed extent     = y: {feed_y_start:.1f} to {feed_y_end:.2f} mm")
print(f"  freq sweep      = {f_min/1e9:.1f} ~ {f_max/1e9:.1f} GHz, {n_pts} pts")
print()

# ============================================================================
# 2. 启动 AEDT
# ============================================================================
print("=" * 70)
print("STEP 2: Launch AEDT")
print("=" * 70)

PROJECT_NAME = "Printed_Dipole_v6"
hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, PROJECT_NAME),
    designname="Dipole_Balun_v6",
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


# ============================================================================
# 3. 创建天线几何模型
# ============================================================================
print()
print("=" * 70)
print("STEP 3: Antenna Geometry")
print("=" * 70)

# 3.1 FR4 基板 (全尺寸延伸到 AirBox 边界)
print(f"  [3.1] Substrate: {sub_L} x {sub_total_y:.1f} x {sub_H} mm")
create_box("Substrate", -sub_L/2, sub_y_start, 0, sub_L, sub_total_y, sub_H, "FR4_epoxy")

# 3.2 截断地板 (底面，只覆盖 y < gnd_truncate_y 的区域)
gnd_y_start = sub_y_start
gnd_y_size = gnd_truncate_y - gnd_y_start
print(f"  [3.2] Ground (TRUNCATED): {sub_L} x {gnd_y_size:.1f} mm, y=[{gnd_y_start:.1f}, {gnd_truncate_y:.1f}]")
create_box("Ground", -sub_L/2, gnd_y_start, -cu_t, sub_L, gnd_y_size, cu_t, "copper")

# 3.3 偶极子臂 (顶面, y=0 中心)
left_x = -(gap/2 + dipole_arm)
right_x = gap/2
print(f"  [3.3] Dipole Left:  X=[{left_x:.1f}, {-gap/2:.1f}], arm={dipole_arm}mm")
create_box("Dipole_Left", left_x, -dipole_w/2, z_top, dipole_arm, dipole_w, cu_t, "copper")
print(f"        Dipole Right: X=[{right_x:.1f}, {right_x+dipole_arm:.1f}]")
create_box("Dipole_Right", right_x, -dipole_w/2, z_top, dipole_arm, dipole_w, cu_t, "copper")

# 3.4 巴伦 (渐变微带，从偶极子中心向 -y 延伸)
# 宽度保持不变 (50Ω微带线 ≈ 3mm on 1.6mm FR4)
print(f"  [3.4] Balun: {balun_segs} segs, L={balun_L:.2f}mm")
seg_len = balun_L / balun_segs
for i in range(balun_segs):
    seg_y = balun_y_top - seg_len * (i + 1)
    create_box(f"Balun_{i}", -feed_w/2, seg_y, z_top, feed_w, seg_len, cu_t, "copper")

# 3.5 馈线 (从 AirBox 边界到巴伦底端)
print(f"  [3.5] Feed: w={feed_w}mm, Y=[{feed_y_start:.1f}, {feed_y_end:.2f}] ({feed_len:.1f}mm)")
create_box("Feed_Line", -feed_w/2, feed_y_start, z_top, feed_w, feed_len, cu_t, "copper")

# 3.6 AirBox + 辐射边界
air_x_size = 2 * (sub_L/2 + pad)
air_y_size = airbox_ymax - airbox_ymin
air_z_size = sub_H + air_z_top + air_z_bot
print(f"  [3.6] AirBox: {air_x_size:.1f} x {air_y_size:.1f} x {air_z_size:.1f} mm")
create_box("Air", airbox_xmin, airbox_ymin, -air_z_bot, air_x_size, air_y_size, air_z_size, "vacuum", True)

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
ymin_face_id = int(air_faces[2])  # face index 2 = y_min (verified)
print(f"  y_min face: {ymin_face_id}")

oBnd.AutoIdentifyPorts(
    ["NAME:Faces", ymin_face_id], True,
    ["NAME:ReferenceConductors", "Ground"],
    "Port1", True)
print("  AutoIdentifyPorts OK")

terms = oBnd.GetExcitationsOfType("Terminal")
terminal = str(terms[0]) if terms else "Feed_Line_T1"
print(f"  Terminal: {terminal}")

# ============================================================================
# 5. Infinite Sphere
# ============================================================================
print()
print("STEP 5: Infinite Sphere")
oRadField = oDesign.GetModule("RadField")
try:
    oRadField.InsertInfiniteSphereDef(
        ["NAME:InfSphere1",
         "UseCustomRadiationSurface:=", False,
         "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "2deg",
         "PhiStart:=", "0deg", "PhiStop:=", "360deg", "PhiStep:=", "2deg",
         "UseLocalCS:=", False])
    print("  OK")
except Exception as e:
    print(f"  Note: {e}")

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
print(f"  Setup1: f0={f0/1e9:.3f}GHz")

oAnalysis.InsertFrequencySweep("Setup1",
    ["NAME:Sweep1",
     "IsEnabled:=", True, "SetupType:=", "LinearCount",
     "StartValue:=", f"{f_min/1e9}GHz", "StopValue:=", f"{f_max/1e9}GHz",
     "Count:=", n_pts,
     "Type:=", "Discrete", "SaveFields:=", True, "SaveRadFields:=", False])
print(f"  Sweep1: {f_min/1e9:.1f}~{f_max/1e9:.1f}GHz, {n_pts} pts")

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

if v != 1:
    print("  ERROR: Design validation failed!")
    # 尝试继续仿真

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

s1p_path = os.path.join(PROJECT_DIR, "dipole_v6.s1p")
s11_csv = os.path.join(PROJECT_DIR, "s11_v6.csv")

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
        ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal},{terminal}))"]],
        [])
    oReport.ExportToFile("S11_Report", s11_csv)
    print(f"  S11 CSV: {s11_csv}")
except Exception as e:
    print(f"  S11 report error: {e}")


# ============================================================================
# 9. S11 分析
# ============================================================================
def parse_s11_csv(path):
    """解析 S11 CSV，返回 (freqs_ghz, s11_db) numpy arrays"""
    f_list, s_list = [], []
    with open(path, 'r') as fh:
        rd = csv.reader(fh)
        next(rd)
        for row in rd:
            if len(row) >= 2:
                try:
                    f_list.append(float(row[0]))
                    s_list.append(float(row[1]))
                except ValueError:
                    continue
    return np.array(f_list), np.array(s_list)


def analyze_s11(freqs, s11, target_f_ghz):
    """分析 S11 数据，返回结果字典"""
    min_idx = np.argmin(s11)
    f_res = freqs[min_idx]
    s11_min = s11[min_idx]
    below_10 = np.where(s11 < -10)[0]
    if len(below_10) >= 2:
        bw_lo = freqs[below_10[0]]
        bw_hi = freqs[below_10[-1]]
        bw_abs = bw_hi - bw_lo
        bw_rel = bw_abs / f_res * 100
    else:
        bw_lo = bw_hi = f_res
        bw_abs = bw_rel = 0
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

    print(f"  Resonance:   {r['f_res']:.4f} GHz")
    print(f"  S11 min:     {r['s11_min']:.2f} dB")
    if r['bw_abs'] > 0:
        print(f"  -10dB BW:    {r['bw_lo']:.3f} ~ {r['bw_hi']:.3f} GHz")
        print(f"  Abs BW:      {r['bw_abs']*1e3:.1f} MHz")
        print(f"  Rel BW:      {r['bw_rel']:.2f}%")
    else:
        print("  WARNING: S11 does not reach -10 dB")
    print(f"  Freq error:  {r['freq_err']:.2f}% {'-> NEEDS OPT' if r['freq_err'] > 2 else '-> OK'}")

    need_sweep = r['freq_err'] > 2.0 or r['s11_min'] > -10
    f_res = r['f_res']
    s11_min = r['s11_min']
    bw_rel = r['bw_rel']
    bw_abs = r['bw_abs']
    bw_lo = r['bw_lo']
    bw_hi = r['bw_hi']
    freq_err = r['freq_err']

except Exception as e:
    print(f"  Parse error: {e}")
    import traceback; traceback.print_exc()
    need_sweep = True
    f_res = f0/1e9; s11_min = 0; bw_rel = bw_abs = 0; bw_lo = bw_hi = f_res; freq_err = 999

# ============================================================================
# 10. S11 绘图
# ============================================================================
print()
print("STEP 10: S11 Plot")
try:
    import matplotlib; matplotlib.use('Agg')
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(freqs_ghz, s11_db, 'b-', lw=2, label='S11')
    ax.axhline(-10, color='r', ls='--', lw=1, label='-10 dB')
    ax.axvline(f0/1e9, color='g', ls=':', lw=1, label=f'Target {f0/1e9:.3f} GHz')
    ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'Resonance {f_res:.3f} GHz')
    if bw_abs > 0:
        ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green',
                   label=f'BW {bw_abs*1e3:.0f}MHz ({bw_rel:.1f}%)')
    ax.set_xlabel('Frequency (GHz)'); ax.set_ylabel('S11 (dB)')
    ax.set_title('Printed Dipole (Truncated GND) - S11')
    ax.legend(); ax.grid(True, alpha=0.3)
    ax.set_xlim(f_min/1e9, f_max/1e9)
    s11_img = os.path.join(PROJECT_DIR, "s11_v6.png")
    fig.savefig(s11_img, dpi=150, bbox_inches='tight'); plt.close(fig)
    print(f"  Saved: {s11_img}")
except Exception as e:
    print(f"  Plot error: {e}")

# ============================================================================
# 11. 远场辐射方向图
# ============================================================================
print()
print("=" * 70)
print("STEP 11: Radiation Pattern")
print("=" * 70)

e_csv = os.path.join(PROJECT_DIR, "e_plane_v6.csv")
h_csv = os.path.join(PROJECT_DIR, "h_plane_v6.csv")
max_gain = 0

for plane_name, phi_val, csv_path in [("E_Plane", "0deg", e_csv), ("H_Plane", "90deg", h_csv)]:
    try:
        oReport.CreateReport(
            plane_name, "Far Fields", "Radiation Pattern",
            "Setup1 : LastAdaptive",
            ["Context:=", "InfSphere1"],
            ["Theta:=", ["All"], "Phi:=", [phi_val], "Freq:=", [f"{f0/1e9}GHz"]],
            ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]],
            [])
        oReport.ExportToFile(plane_name, csv_path)
        print(f"  {plane_name}: {csv_path}")
    except Exception as e:
        print(f"  {plane_name} error: {e}")

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
        max_gain_e = np.max(g_e) if len(g_e) else -999
        max_gain_h = np.max(g_h) if len(g_h) else -999
        max_gain = max(max_gain_e, max_gain_h)
        print(f"\n  E-plane max gain: {max_gain_e:.2f} dBi")
        print(f"  H-plane max gain: {max_gain_h:.2f} dBi")
        print(f"  Antenna gain:     {max_gain:.2f} dBi")

        fig, (ax1, ax2) = plt.subplots(1, 2, subplot_kw={'projection': 'polar'}, figsize=(14, 6))
        ax1.plot(np.radians(th_e), g_e, 'b-', lw=2)
        ax1.set_title(f'E-Plane (xz)\nGain={max_gain_e:.2f} dBi', pad=20)
        ax1.set_theta_zero_location('N'); ax1.set_theta_direction(-1)
        ax2.plot(np.radians(th_h), g_h, 'r-', lw=2)
        ax2.set_title(f'H-Plane (yz)\nGain={max_gain_h:.2f} dBi', pad=20)
        ax2.set_theta_zero_location('N'); ax2.set_theta_direction(-1)
        pat_img = os.path.join(PROJECT_DIR, "radiation_pattern_v6.png")
        fig.savefig(pat_img, dpi=150, bbox_inches='tight'); plt.close(fig)
        print(f"  Pattern plot: {pat_img}")
    else:
        print("  Pattern CSV missing")
except Exception as e:
    print(f"  Pattern error: {e}")

# ============================================================================
# 12. 参数扫描优化
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

    for opt_iter in range(8):
        # 频率缩放：arm_new = arm_old * (f_res / f_target)
        new_arm = round(current_arm * (best_f_res / (f0/1e9)), 1)

        if abs(new_arm - current_arm) < 0.3:
            print(f"  Iter {opt_iter+1}: arm change < 0.3mm, done")
            break

        print(f"\n  Iter {opt_iter+1}: arm {current_arm:.1f} -> {new_arm:.1f} mm")

        try:
            oEditor.Delete(["NAME:Selections", "Selections:=", "Dipole_Left,Dipole_Right"])
        except Exception as e:
            print(f"    Delete error: {e}"); break

        new_left_x = -(gap/2 + new_arm)
        create_box("Dipole_Left", new_left_x, -dipole_w/2, z_top, new_arm, dipole_w, cu_t, "copper")
        create_box("Dipole_Right", gap/2, -dipole_w/2, z_top, new_arm, dipole_w, cu_t, "copper")
        print(f"    New arms: {new_arm:.1f}mm each side")

        hfss.save_project()
        t_opt = time.time()
        print("    Analyzing...")
        oDesign.Analyze("Setup1")
        print(f"    Done in {time.time()-t_opt:.0f}s")

        # 提取 S11
        opt_csv = os.path.join(PROJECT_DIR, f"s11_v6_opt{opt_iter+1}.csv")
        try:
            rname = f"S11_Opt{opt_iter+1}"
            oReport.CreateReport(
                rname, "Terminal Solution Data", "Rectangular Plot",
                "Setup1 : Sweep1", [],
                ["Freq:=", ["All"]],
                ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal},{terminal}))"]],
                [])
            oReport.ExportToFile(rname, opt_csv)

            opt_f, opt_s = parse_s11_csv(opt_csv)
            r = analyze_s11(opt_f, opt_s, f0/1e9)

            print(f"    f_res={r['f_res']:.4f}GHz, S11={r['s11_min']:.2f}dB, err={r['freq_err']:.2f}%")
            if r['bw_abs'] > 0:
                print(f"    -10dB BW: {r['bw_lo']:.3f}~{r['bw_hi']:.3f}GHz ({r['bw_rel']:.1f}%)")

            current_arm = new_arm
            best_f_res = r['f_res']
            best_s11 = r['s11_min']
            best_arm = new_arm

            if r['freq_err'] <= 2.0 and r['s11_min'] < -10:
                print(f"    CONVERGED! f_err<2%, S11<-10dB")
                f_res = r['f_res']; s11_min = r['s11_min']
                bw_rel = r['bw_rel']; bw_abs = r['bw_abs']
                bw_lo = r['bw_lo']; bw_hi = r['bw_hi']
                freq_err = r['freq_err']
                freqs_ghz = opt_f; s11_db = opt_s
                break

            if r['freq_err'] <= 2.0:
                print(f"    Frequency OK, but S11 > -10dB. Stopping.")
                f_res = r['f_res']; s11_min = r['s11_min']
                freq_err = r['freq_err']
                freqs_ghz = opt_f; s11_db = opt_s
                break

        except Exception as e:
            print(f"    Opt error: {e}")
            import traceback; traceback.print_exc()
            break
    else:
        print(f"\n  Max iterations reached. Best: arm={best_arm:.1f}mm, f_res={best_f_res:.4f}GHz")
        f_res = best_f_res; s11_min = best_s11

    # 最终 S11 绘图
    if len(freqs_ghz) > 0:
        try:
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.plot(freqs_ghz, s11_db, 'b-', lw=2, label='S11 (optimized)')
            ax.axhline(-10, color='r', ls='--', lw=1)
            ax.axvline(f0/1e9, color='g', ls=':', lw=1, label=f'Target {f0/1e9:.3f}')
            ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'Res {f_res:.3f}')
            if bw_abs > 0:
                ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green',
                           label=f'BW {bw_abs*1e3:.0f}MHz')
            ax.set_xlabel('Frequency (GHz)'); ax.set_ylabel('S11 (dB)')
            ax.set_title('Printed Dipole (Opt) - S11')
            ax.legend(); ax.grid(True, alpha=0.3); ax.set_xlim(f_min/1e9, f_max/1e9)
            fig.savefig(os.path.join(PROJECT_DIR, "s11_v6_opt.png"), dpi=150, bbox_inches='tight')
            plt.close(fig)
            print(f"  Opt S11 plot saved")
        except Exception as e:
            print(f"  Opt plot error: {e}")

    # 优化后远场
    try:
        for pn, pv, cp in [("E_Opt", "0deg", "e_plane_v6_opt.csv"),
                           ("H_Opt", "90deg", "h_plane_v6_opt.csv")]:
            oReport.CreateReport(
                pn, "Far Fields", "Radiation Pattern",
                "Setup1 : LastAdaptive", ["Context:=", "InfSphere1"],
                ["Theta:=", ["All"], "Phi:=", [pv], "Freq:=", [f"{f0/1e9}GHz"]],
                ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]], [])
            oReport.ExportToFile(pn, os.path.join(PROJECT_DIR, cp))

        e_opt_csv = os.path.join(PROJECT_DIR, "e_plane_v6_opt.csv")
        h_opt_csv = os.path.join(PROJECT_DIR, "h_plane_v6_opt.csv")
        if os.path.exists(e_opt_csv) and os.path.exists(h_opt_csv):
            th_e, g_e = read_csv_pattern(e_opt_csv)
            th_h, g_h = read_csv_pattern(h_opt_csv)
            max_gain_e = np.max(g_e) if len(g_e) else -999
            max_gain_h = np.max(g_h) if len(g_h) else -999
            max_gain = max(max_gain_e, max_gain_h)
            print(f"  Opt E-plane gain: {max_gain_e:.2f} dBi")
            print(f"  Opt H-plane gain: {max_gain_h:.2f} dBi")
            print(f"  Opt antenna gain: {max_gain:.2f} dBi")

            fig, (a1, a2) = plt.subplots(1, 2, subplot_kw={'projection': 'polar'}, figsize=(14, 6))
            a1.plot(np.radians(th_e), g_e, 'b-', lw=2)
            a1.set_title(f'E-Plane (Opt)\n{max_gain_e:.2f} dBi', pad=20)
            a1.set_theta_zero_location('N'); a1.set_theta_direction(-1)
            a2.plot(np.radians(th_h), g_h, 'r-', lw=2)
            a2.set_title(f'H-Plane (Opt)\n{max_gain_h:.2f} dBi', pad=20)
            a2.set_theta_zero_location('N'); a2.set_theta_direction(-1)
            fig.savefig(os.path.join(PROJECT_DIR, "radiation_v6_opt.png"), dpi=150, bbox_inches='tight')
            plt.close(fig)
            print(f"  Opt pattern plot saved")
    except Exception as e:
        print(f"  Opt pattern error: {e}")
else:
    print("\n  Frequency error < 2% & S11 < -10dB, no optimization needed")
    best_arm = dipole_arm

# ============================================================================
# 13. 设计总结
# ============================================================================
print()
print("=" * 70)
print("DESIGN SUMMARY")
print("=" * 70)
print(f"  Antenna:         Printed Dipole + Truncated Ground + Microstrip Balun")
print(f"  Target freq:     {f0/1e9:.3f} GHz")
print(f"  Substrate:       FR4 (er={eps_r}), {sub_L}x{sub_W}x{sub_H} mm")
print(f"  Dipole arm:      {dipole_arm:.1f} mm (initial)")
if need_sweep:
    print(f"  Optimized arm:   {best_arm:.1f} mm")
print(f"  Total dipole:    {2*best_arm+gap:.1f} mm")
print(f"  Dipole width:    {dipole_w:.1f} mm")
print(f"  Balun length:    {balun_L:.2f} mm (lambda_g/4)")
print(f"  Ground truncate: y < {gnd_truncate_y:.1f} mm")
print(f"  Resonance:       {f_res:.4f} GHz")
print(f"  S11 min:         {s11_min:.2f} dB")
if bw_abs > 0:
    print(f"  -10dB BW:        {bw_lo:.3f} ~ {bw_hi:.3f} GHz")
    print(f"  Rel bandwidth:   {bw_rel:.2f}%")
print(f"  Max gain:        {max_gain:.2f} dBi")
print(f"  Freq error:      {freq_err:.2f}%")
print()

hfss.save_project()
print("=" * 70)
print("COMPLETE")
print("=" * 70)
