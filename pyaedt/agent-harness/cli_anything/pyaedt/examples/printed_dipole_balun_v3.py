#!/usr/bin/env python3
"""
印刷偶极子天线（带微带巴伦馈线）- HFSS 仿真设计
Printed Dipole Antenna with Microstrip Balun - HFSS Simulation

设计参数:
  - 天线类型：印刷偶极子 + 三角形微带巴伦
  - 中心频率：2.217 GHz
  - 介质板：FR4 (εr=4.4, 厚度1.6mm)
  - 偶极子臂长初始值：50 mm
  - 频率扫描：1.5 ~ 3.0 GHz，步进 0.01 GHz

输出:
  1. S11 反射系数曲线 + 相对带宽
  2. E面(xz) / H面(yz) 辐射方向图
  3. 天线增益
  4. 臂长参数扫描优化（如需要）

需要环境:
  - Windows + AEDT 2019.1 (AnsysEM19.3)
  - PyAEDT 0.8.x
  - numpy, matplotlib
"""

import numpy as np
import os
import sys
import time

# ============================================================================
# 0. 环境和兼容性配置
# ============================================================================
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'

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
print("步骤 1: 设计参数计算")
print("=" * 70)

c = 3e8          # 光速 m/s
f0 = 2.217e9     # 中心频率 Hz
eps_r = 4.4      # FR4 介电常数

# 波长
lambda_0 = c / f0                      # 自由空间波长
lambda_g = lambda_0 / np.sqrt(eps_r)   # FR4中有效波长

# 介质板 (单位 mm)
sub_L = 120.0    # 长度
sub_W = 60.0     # 宽度
sub_H = 1.6      # 厚度
cu_t  = 0.035    # 铜箔厚度

# 偶极子 (mm)
dipole_arm = 50.0   # 初始臂长（单臂），后续参数扫描
dipole_w   = 3.0    # 臂宽

# 微带巴伦 (mm) — 1/4 有效波长
balun_L = lambda_g * 1e3 / 4.0         # ≈16.1 mm
balun_w_start = 1.0                    # 窄端宽度
balun_w_end   = 3.0                    # 宽端宽度（匹配偶极子臂宽）
balun_segs = 10                        # 渐变段数

# 频率扫描
f_min  = 1.5e9
f_max  = 3.0e9
f_step = 0.01e9

# 间距
gap = 1.0  # 偶极子两臂间隙 mm

print(f"  中心频率 f0        = {f0/1e9:.3f} GHz")
print(f"  自由空间波长 λ0    = {lambda_0*1e3:.2f} mm")
print(f"  FR4有效波长 λg     = {lambda_g*1e3:.2f} mm")
print(f"  偶极子臂长(初始)   = {dipole_arm:.1f} mm")
print(f"  微带巴伦长度 λg/4  = {balun_L:.2f} mm")
print(f"  介质板尺寸         = {sub_L} x {sub_W} x {sub_H} mm")
print(f"  频率扫描           = {f_min/1e9:.1f} ~ {f_max/1e9:.1f} GHz, step {f_step/1e9:.2f} GHz")
print()

# ============================================================================
# 2. 启动 AEDT, 创建 HFSS 项目
# ============================================================================
print("=" * 70)
print("步骤 2: 启动 AEDT 并创建 HFSS 项目")
print("=" * 70)

PROJECT_DIR = r"D:\class_design"
os.makedirs(PROJECT_DIR, exist_ok=True)
PROJECT_NAME = "Printed_Dipole_Balun_v3"

hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, PROJECT_NAME),
    designname="Dipole_Balun",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)

print(f"  项目: {hfss.project_name}")
print(f"  设计: {hfss.design_name}")

hfss.modeler.model_units = "mm"
oEditor = hfss.odesign.SetActiveEditor("3D Modeler")  # COM handle

# ============================================================================
# helper — 用 COM 接口快速创建 box
# ============================================================================
def make_box(name, x, y, z, dx, dy, dz, mat, solve_inside=None):
    """通过 COM 创建 box，避免 pyaedt 高版本 API 兼容问题。"""
    if solve_inside is None:
        solve_inside = mat.lower() not in ("copper", "pec", "aluminum")
    oEditor.CreateBox(
        ["NAME:BoxParameters",
         "XPosition:=", str(x),
         "YPosition:=", str(y),
         "ZPosition:=", str(z),
         "XSize:=",     str(dx),
         "YSize:=",     str(dy),
         "ZSize:=",     str(dz)],
        ["NAME:Attributes",
         "Name:=",        name,
         "Flags:=",       "",
         "Color:=",       "(143 175 131)",
         "Transparency:=", 0,
         "PartCoordinateSystem:=", "Global",
         "UDMId:=",       "",
         "MaterialValue:=", '"' + mat + '"',
         "SurfaceMaterialValue:=", '""',
         "SolveInside:=", solve_inside,
         "IsMaterialEditable:=", True,
         "UseMaterialAppearance:=", False,
         "IsLightweight:=", False])

def make_polyline_sheet(name, points, mat="copper"):
    """用折线创建薄片(Sheet)，适用于二维导体。"""
    pts_args = ["NAME:PolylinePoints"]
    for p in points:
        pts_args.append(["NAME:PLPoint",
                         "X:=", str(p[0]),
                         "Y:=", str(p[1]),
                         "Z:=", str(p[2])])
    seg_args = ["NAME:PolylineSegments"]
    for i in range(len(points) - 1):
        seg_args.append(["NAME:PLSegment",
                         "SegmentType:=", "Line",
                         "StartIndex:=", i,
                         "NoOfPoints:=", 2])
    oEditor.CreatePolyline(
        ["NAME:PolylineParameters",
         "IsPolylineCovered:=", True,
         "IsPolylineClosed:=", True,
         pts_args, seg_args,
         ["NAME:PolylineXSection",
          "XSectionType:=", "None",
          "XSectionOrient:=", "Auto",
          "XSectionWidth:=", "0mm",
          "XSectionTopWidth:=", "0mm",
          "XSectionHeight:=", "0mm",
          "XSectionNumSegments:=", "0",
          "XSectionBendType:=", "Corner"]],
        ["NAME:Attributes",
         "Name:=", name,
         "Flags:=", "",
         "Color:=", "(255 128 0)",
         "Transparency:=", 0,
         "PartCoordinateSystem:=", "Global",
         "UDMId:=", "",
         "MaterialValue:=", '"' + mat + '"',
         "SurfaceMaterialValue:=", '""',
         "SolveInside:=", False,
         "IsMaterialEditable:=", True,
         "UseMaterialAppearance:=", False,
         "IsLightweight:=", False])

# ============================================================================
# 3. 建模 — 介质板
# ============================================================================
print()
print("=" * 70)
print("步骤 3: 创建天线几何模型")
print("=" * 70)

print("  [3.1] FR4 介质板 ...")
make_box("Substrate",
         -sub_L/2, -sub_W/2, 0,
         sub_L, sub_W, sub_H,
         "FR4_epoxy")
print(f"        {sub_L} x {sub_W} x {sub_H} mm")

# ============================================================================
# 4. 建模 — 地板（在介质板底面）
# ============================================================================
print("  [3.2] 铜地板（底面）...")
make_box("Ground",
         -sub_L/2, -sub_W/2, 0,
         sub_L, sub_W, -cu_t,
         "copper")
print(f"        {sub_L} x {sub_W} x {cu_t} mm (z < 0)")

# ============================================================================
# 5. 建模 — 偶极子两臂（在介质板顶面）
# ============================================================================
print("  [3.3] 偶极子臂 ...")
z_top = sub_H  # 介质板顶面

# 左臂：从 -(gap/2 + dipole_arm) 到 -gap/2
left_x = -(gap/2 + dipole_arm)
make_box("Dipole_Left",
         left_x, -dipole_w/2, z_top,
         dipole_arm, dipole_w, cu_t,
         "copper")
print(f"        左臂: x=[{left_x:.1f}, {-gap/2:.1f}] mm, 长度 {dipole_arm} mm")

# 右臂：从 +gap/2 到 +(gap/2 + dipole_arm)
right_x = gap / 2
make_box("Dipole_Right",
         right_x, -dipole_w/2, z_top,
         dipole_arm, dipole_w, cu_t,
         "copper")
print(f"        右臂: x=[{right_x:.1f}, {right_x+dipole_arm:.1f}] mm, 长度 {dipole_arm} mm")

# ============================================================================
# 6. 建模 — 三角形微带巴伦（渐变型，在底面地板侧做开槽）
#    实际设计：微带到平衡馈线的过渡放在基板顶面
#    巴伦从偶极子中心向下延伸，宽度线性渐变
# ============================================================================
print("  [3.4] 微带巴伦（渐变馈线）...")

# 巴伦从 y=0 向 -y 方向延伸 balun_L
# 宽度从 balun_w_end（靠近偶极子）渐变到 balun_w_start（馈电端）
for i in range(balun_segs):
    frac_start = i / balun_segs
    frac_end = (i + 1) / balun_segs
    w_s = balun_w_end - (balun_w_end - balun_w_start) * frac_start
    w_e = balun_w_end - (balun_w_end - balun_w_start) * frac_end
    avg_w = (w_s + w_e) / 2.0
    seg_len = balun_L / balun_segs
    y_pos = -dipole_w / 2 - seg_len * (i + 1)
    make_box(f"Balun_{i}",
             -avg_w / 2, y_pos, z_top,
             avg_w, seg_len, cu_t,
             "copper")

# 馈电微带线延长段（连接到端口）
feed_ext_len = 5.0  # mm
feed_y = -dipole_w / 2 - balun_L - feed_ext_len
make_box("Feed_Line",
         -balun_w_start / 2, feed_y, z_top,
         balun_w_start, feed_ext_len, cu_t,
         "copper")

print(f"        巴伦段数: {balun_segs}, 长度 {balun_L:.2f} mm")
print(f"        宽度渐变: {balun_w_end} → {balun_w_start} mm")
print(f"        馈电延长线: {feed_ext_len} mm")

# ============================================================================
# 7. 激励端口 — Lumped Port
# ============================================================================
print("  [3.5] 激励端口 ...")

# 在馈电线底端创建矩形薄片作为端口面
port_y = feed_y
port_z = z_top

# 创建端口矩形 sheet（从馈线底端到地板）
port_pts = [
    (-balun_w_start/2, port_y, 0),
    ( balun_w_start/2, port_y, 0),
    ( balun_w_start/2, port_y, z_top + cu_t),
    (-balun_w_start/2, port_y, z_top + cu_t),
    (-balun_w_start/2, port_y, 0),  # 闭合
]
make_polyline_sheet("Port_Sheet", port_pts, "vacuum")

# 使用 COM 分配 lumped port
oModule_Bnd = hfss.odesign.GetModule("BoundarySetup")
oModule_Bnd.AssignLumpedPort(
    ["NAME:Port1",
     "Objects:=", ["Port_Sheet"],
     "RenormalizeAllTerminals:=", True,
     "DoDeembed:=", False,
     ["NAME:Modes",
      ["NAME:Mode1",
       "ModeNum:=", 1,
       "UseIntLine:=", True,
       ["NAME:IntLine",
        "Start:=", [str(0), str(port_y), str(0)],
        "End:=",   [str(0), str(port_y), str(z_top + cu_t)]],
       "CharImp:=", "Zpi"]]])
print("        Lumped Port 'Port1' (50Ω) 已创建")

# ============================================================================
# 8. 辐射边界 — 空气盒 + Radiation Boundary
# ============================================================================
print("  [3.6] 辐射边界 ...")

pad = lambda_0 * 1e3 / 4.0  # 距模型 ≥ λ0/4 的空气区域
air_x = sub_L / 2 + pad
air_y = sub_W / 2 + pad + balun_L + feed_ext_len
air_z_top = pad
air_z_bot = pad

make_box("AirBox",
         -air_x, -air_y, -air_z_bot,
         2 * air_x, 2 * air_y, sub_H + air_z_top + air_z_bot,
         "vacuum", solve_inside=True)

# Radiation boundary on all faces of AirBox
oModule_Bnd.AssignRadiation(
    ["NAME:Radiation1",
     "Objects:=", ["AirBox"],
     "IsFssReference:=", False,
     "IsForPML:=", False])
print(f"        AirBox padding = {pad:.1f} mm (λ0/4)")

# ============================================================================
# 9. 仿真设置 — Setup + 频率扫描
# ============================================================================
print()
print("=" * 70)
print("步骤 4: 创建仿真设置")
print("=" * 70)

oModule_Analysis = hfss.odesign.GetModule("AnalysisSetup")
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

print(f"  Setup1: 解析频率 {f0/1e9:.3f} GHz, MaxDeltaS=0.02, MaxPasses=15")

n_points = int((f_max - f_min) / f_step) + 1
oModule_Analysis.InsertFrequencySweep("Setup1",
    ["NAME:Sweep1",
     "IsEnabled:=", True,
     "SetupType:=", "LinearCount",
     "StartValue:=", f"{f_min/1e9}GHz",
     "StopValue:=",  f"{f_max/1e9}GHz",
     "Count:=", n_points,
     "Type:=", "Discrete",
     "SaveFields:=", True,
     "SaveRadFields:=", True])

print(f"  Sweep1: {f_min/1e9:.1f} ~ {f_max/1e9:.1f} GHz, {n_points} 点 (Discrete)")

# ============================================================================
# 10. 保存项目
# ============================================================================
print()
print("  保存项目 ...")
hfss.save_project()
print(f"  项目已保存至: {hfss.project_path}")

# ============================================================================
# 11. 运行仿真
# ============================================================================
print()
print("=" * 70)
print("步骤 5: 运行仿真")
print("=" * 70)

t_start = time.time()
print("  正在运行 Setup1 ... (请耐心等待)")

oModule_Analysis.Solve(["Setup1"])

t_elapsed = time.time() - t_start
print(f"  仿真完成! 用时 {t_elapsed/60:.1f} 分钟")

# ============================================================================
# 12. 提取 S11 结果
# ============================================================================
print()
print("=" * 70)
print("步骤 6: S11 反射系数分析")
print("=" * 70)

oModule_Report = hfss.odesign.GetModule("ReportSetup")

# 创建 S11 报告
oModule_Report.CreateReport(
    "S11_Plot", "Terminal Solution Data", "Rectangular Plot",
    "Setup1 : Sweep1",
    ["Domain:=", "Sweep"],
    ["Freq:=", ["All"]],
    ["X Component:=", "Freq",
     "Y Component:=", ["dB(St(Port1,Port1))"]])

# 导出 S11 数据到 CSV
s11_csv = os.path.join(PROJECT_DIR, "s11_data.csv")
oModule_Report.ExportToFile("S11_Plot", s11_csv)
print(f"  S11 数据已导出: {s11_csv}")

# 读取并分析 S11
try:
    import csv
    freqs_ghz = []
    s11_db = []
    with open(s11_csv, 'r') as f:
        reader = csv.reader(f)
        header = next(reader)  # skip header
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

    print(f"\n  -------- S11 分析结果 --------")
    print(f"  谐振频率: {f_res:.4f} GHz")
    print(f"  最小 S11:  {s11_min:.2f} dB")

    # -10 dB 带宽
    below_10 = np.where(s11_db < -10)[0]
    if len(below_10) >= 2:
        bw_lo = freqs_ghz[below_10[0]]
        bw_hi = freqs_ghz[below_10[-1]]
        bw_abs = bw_hi - bw_lo
        bw_rel = bw_abs / f_res * 100
        print(f"  -10dB 带宽: {bw_lo:.3f} ~ {bw_hi:.3f} GHz")
        print(f"  绝对带宽:   {bw_abs*1e3:.1f} MHz")
        print(f"  相对带宽:   {bw_rel:.2f}%")
    else:
        bw_rel = 0
        print("  S11 未达 -10 dB，天线匹配不佳")

    # 判断是否需要优化
    freq_err = abs(f_res - f0/1e9) / (f0/1e9) * 100
    need_sweep = freq_err > 2.0
    print(f"\n  谐振偏移: {freq_err:.2f}%  {'→ 需要参数扫描优化' if need_sweep else '→ 合格'}")

except Exception as e:
    print(f"  读取 S11 CSV 失败: {e}")
    need_sweep = True
    f_res = f0 / 1e9
    s11_min = 0

# ============================================================================
# 13. S11 曲线图
# ============================================================================
print()
print("  正在生成 S11 曲线图 ...")
try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(freqs_ghz, s11_db, 'b-', linewidth=2, label='S11')
    ax.axhline(-10, color='r', ls='--', lw=1, label='-10 dB')
    ax.axvline(f0/1e9, color='g', ls=':', lw=1, label=f'目标 {f0/1e9:.3f} GHz')
    ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'谐振 {f_res:.3f} GHz')
    if len(below_10) >= 2:
        ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green', label=f'BW {bw_abs*1e3:.0f} MHz')
    ax.set_xlabel('Frequency (GHz)', fontsize=13)
    ax.set_ylabel('S11 (dB)', fontsize=13)
    ax.set_title('Printed Dipole Antenna — S11', fontsize=14)
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3)
    ax.set_xlim(f_min/1e9, f_max/1e9)
    s11_img = os.path.join(PROJECT_DIR, "s11_plot.png")
    fig.savefig(s11_img, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  S11 图已保存: {s11_img}")
except Exception as e:
    print(f"  绘图失败: {e}")

# ============================================================================
# 14. 远场辐射方向图
# ============================================================================
print()
print("=" * 70)
print("步骤 7: 辐射方向图 & 增益")
print("=" * 70)

# 创建 Infinite Sphere
oModule_RadField = hfss.odesign.GetModule("RadField")
try:
    oModule_RadField.InsertInfiniteSphereDef(
        ["NAME:Infinite_Sphere",
         "UseCustomRadiationSurface:=", False,
         "ThetaStart:=", "-180deg",
         "ThetaStop:=", "180deg",
         "ThetaStep:=", "1deg",
         "PhiStart:=", "0deg",
         "PhiStop:=", "360deg",
         "PhiStep:=", "1deg",
         "UseLocalCS:=", False])
    print("  Infinite Sphere 已创建")
except Exception as e:
    print(f"  Infinite Sphere 创建提示: {e}")

# E-plane (phi=0, xz面) 方向图
oModule_Report.CreateReport(
    "E_Plane", "Far Fields", "Radiation Pattern",
    "Setup1 : LastAdaptive",
    ["Context:=", "Infinite_Sphere"],
    ["Theta:=", ["All"], "Phi:=", ["0deg"], "Freq:=", [f"{f0/1e9}GHz"]],
    ["X Component:=", "Theta",
     "Y Component:=", ["GainTotal"]])

e_csv = os.path.join(PROJECT_DIR, "e_plane.csv")
oModule_Report.ExportToFile("E_Plane", e_csv)
print(f"  E面方向图数据已导出: {e_csv}")

# H-plane (phi=90, yz面) 方向图
oModule_Report.CreateReport(
    "H_Plane", "Far Fields", "Radiation Pattern",
    "Setup1 : LastAdaptive",
    ["Context:=", "Infinite_Sphere"],
    ["Theta:=", ["All"], "Phi:=", ["90deg"], "Freq:=", [f"{f0/1e9}GHz"]],
    ["X Component:=", "Theta",
     "Y Component:=", ["GainTotal"]])

h_csv = os.path.join(PROJECT_DIR, "h_plane.csv")
oModule_Report.ExportToFile("H_Plane", h_csv)
print(f"  H面方向图数据已导出: {h_csv}")

# 读取 & 绘制方向图
try:
    def read_pattern_csv(path):
        theta, gain = [], []
        with open(path, 'r') as f:
            reader = csv.reader(f)
            next(reader)  # skip header
            for row in reader:
                if len(row) >= 2:
                    try:
                        theta.append(float(row[0]))
                        gain.append(float(row[1]))
                    except ValueError:
                        continue
        return np.array(theta), np.array(gain)

    th_e, g_e = read_pattern_csv(e_csv)
    th_h, g_h = read_pattern_csv(h_csv)

    max_gain_e = np.max(g_e) if len(g_e) > 0 else 0
    max_gain_h = np.max(g_h) if len(g_h) > 0 else 0
    max_gain = max(max_gain_e, max_gain_h)

    print(f"\n  -------- 辐射方向图结果 --------")
    print(f"  E面最大增益: {max_gain_e:.2f} dBi")
    print(f"  H面最大增益: {max_gain_h:.2f} dBi")
    print(f"  天线最大增益: {max_gain:.2f} dBi")

    # 极坐标绘图
    fig, (ax1, ax2) = plt.subplots(1, 2, subplot_kw={'projection': 'polar'},
                                    figsize=(14, 6))

    ax1.plot(np.radians(th_e), g_e, 'b-', lw=2)
    ax1.set_title(f'E-Plane (xz, φ=0°)\nMax Gain = {max_gain_e:.2f} dBi', pad=20)
    ax1.set_theta_zero_location('N')
    ax1.set_theta_direction(-1)

    ax2.plot(np.radians(th_h), g_h, 'r-', lw=2)
    ax2.set_title(f'H-Plane (yz, φ=90°)\nMax Gain = {max_gain_h:.2f} dBi', pad=20)
    ax2.set_theta_zero_location('N')
    ax2.set_theta_direction(-1)

    pattern_img = os.path.join(PROJECT_DIR, "radiation_pattern.png")
    fig.savefig(pattern_img, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  方向图已保存: {pattern_img}")

except Exception as e:
    print(f"  方向图处理失败: {e}")
    max_gain = 0

# ============================================================================
# 15. 参数扫描优化（如果谐振频率偏移 > 2%）
# ============================================================================
if need_sweep:
    print()
    print("=" * 70)
    print("步骤 8: 参数扫描优化臂长")
    print("=" * 70)

    # 根据谐振偏移方向确定扫描范围
    # 如果谐振频率偏高 → 臂长需要增大; 偏低 → 臂长需要减小
    if f_res > f0 / 1e9:
        arm_min = dipole_arm
        arm_max = dipole_arm * 1.3
    else:
        arm_min = dipole_arm * 0.7
        arm_max = dipole_arm
    arm_step = 1.0  # mm

    print(f"  扫描范围: {arm_min:.0f} ~ {arm_max:.0f} mm, 步进 {arm_step} mm")
    print(f"  目标谐振: {f0/1e9:.3f} GHz")
    print()

    # 创建设计变量
    hfss.odesign.ChangeProperty(
        ["NAME:AllTabs",
         ["NAME:LocalVariableTab",
          ["NAME:PropServers", "LocalVariables"],
          ["NAME:NewProps",
           ["NAME:arm_len",
            "PropType:=", "VariableProp",
            "UserDef:=", True,
            "Value:=", f"{dipole_arm}mm"]]]])

    # 创建参数扫描
    oModule_Optimetrics = hfss.odesign.GetModule("Optimetrics")
    oModule_Optimetrics.InsertSetup("OptiParametric",
        ["NAME:ParametricSetup1",
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
           "Data:=", f"LIN {arm_min}mm {arm_max}mm {arm_step}mm",
           "OffsetF1:=", False,
           "Synchronize:=", 0]],
         ["NAME:Sweep Operations"],
         ["NAME:Goals"]])

    print("  参数扫描设置完成，开始运行 ...")
    hfss.save_project()

    t2 = time.time()
    oModule_Optimetrics.SolveSetup("ParametricSetup1")
    t2_elapsed = time.time() - t2
    print(f"  参数扫描完成! 用时 {t2_elapsed/60:.1f} 分钟")

    # 导出各臂长对应的 S11
    print("  正在分析各臂长结果 ...")

    best_arm = dipole_arm
    best_err = 999
    best_s11 = 0

    arm_vals = np.arange(arm_min, arm_max + arm_step/2, arm_step)
    for arm_val in arm_vals:
        try:
            sweep_csv = os.path.join(PROJECT_DIR, f"s11_arm_{arm_val:.0f}.csv")
            # 设置变量值并导出
            oModule_Report.CreateReport(
                f"S11_arm{arm_val:.0f}", "Terminal Solution Data", "Rectangular Plot",
                "Setup1 : Sweep1",
                ["Domain:=", "Sweep", "arm_len:=", [f"{arm_val}mm"]],
                ["Freq:=", ["All"]],
                ["X Component:=", "Freq",
                 "Y Component:=", ["dB(St(Port1,Port1))"]])
            oModule_Report.ExportToFile(f"S11_arm{arm_val:.0f}", sweep_csv)

            # 分析
            f_list, s_list = [], []
            with open(sweep_csv, 'r') as f:
                reader = csv.reader(f)
                next(reader)
                for row in reader:
                    if len(row) >= 2:
                        try:
                            f_list.append(float(row[0]))
                            s_list.append(float(row[1]))
                        except ValueError:
                            continue
            if len(s_list) > 0:
                idx = np.argmin(s_list)
                fr = f_list[idx]
                err = abs(fr - f0/1e9)
                print(f"    arm={arm_val:.0f}mm → f_res={fr:.4f} GHz, S11={s_list[idx]:.2f} dB, 偏差={err*1e3:.1f} MHz")
                if err < best_err:
                    best_err = err
                    best_arm = arm_val
                    best_s11 = s_list[idx]
        except Exception as e:
            print(f"    arm={arm_val:.0f}mm → 分析失败: {e}")

    print(f"\n  ======== 优化结果 ========")
    print(f"  最优臂长: {best_arm:.1f} mm")
    print(f"  谐振偏差: {best_err*1e3:.1f} MHz")
    print(f"  S11:      {best_s11:.2f} dB")
else:
    print()
    print("  谐振频率偏差 < 2%，无需参数扫描优化")

# ============================================================================
# 16. 最终设计总结
# ============================================================================
print()
print("=" * 70)
print("设计总结")
print("=" * 70)
print(f"  天线类型:          印刷偶极子 + 微带巴伦")
print(f"  中心频率(目标):    {f0/1e9:.3f} GHz")
print(f"  介质板:            FR4 (εr={eps_r}), {sub_L}x{sub_W}x{sub_H} mm")
print(f"  偶极子臂长:        {dipole_arm:.1f} mm (初始)")
if need_sweep:
    print(f"  优化臂长:          {best_arm:.1f} mm")
print(f"  偶极子臂宽:        {dipole_w:.1f} mm")
print(f"  巴伦长度:          {balun_L:.2f} mm (λg/4)")
print(f"  谐振频率(仿真):    {f_res:.4f} GHz")
print(f"  最小 S11:          {s11_min:.2f} dB")
if 'bw_rel' in dir() and bw_rel > 0:
    print(f"  -10dB 相对带宽:    {bw_rel:.2f}%")
print(f"  最大增益:          {max_gain:.2f} dBi")
print()
print(f"  输出文件:")
print(f"    S11 数据:  {s11_csv}")
print(f"    S11 图:    {os.path.join(PROJECT_DIR, 's11_plot.png')}")
print(f"    E面方向图: {e_csv}")
print(f"    H面方向图: {h_csv}")
print(f"    方向图:    {os.path.join(PROJECT_DIR, 'radiation_pattern.png')}")
print()

# ============================================================================
# 17. 释放 AEDT
# ============================================================================
print("释放 AEDT ...")
hfss.release_desktop()
print("完成!")
