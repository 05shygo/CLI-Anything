#!/usr/bin/env python3
"""
monopole_antennas_2210_5800.py
单频与双频单极子天线设计及数据提取

使用 cli-anything-pyaedt 在 HFSS 中完成两款单极子天线的建模与仿真：
  - Design 1: Single_Band_Monopole  — 中心频率 2.21 GHz
  - Design 2: Dual_Band_Monopole    — 双频 2.21 GHz + 5.8 GHz
  - 后处理: S11、相对带宽、3D/2D 辐射方向图、峰值增益

天线结构:
  - 介质板: FR4_epoxy (εr=4.4, H_sub=1.6mm)
  - 顶层: 微带馈线 + 单极子辐射臂 (2D Sheet, Perfect E)
  - 底层: 截断地平面 DGS，覆盖下半部分 (2D Sheet, Perfect E)
  - 馈电: Lumped Port 50Ω，积分线从地平面到馈线
  - 求解: Driven Modal
"""

import os, sys, time, csv, traceback, subprocess, gc
import numpy as np

# ============================================================================
# 全局配置
# ============================================================================
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design\final"
PROJ_SUB   = os.path.join(PROJECT_DIR, "project")
DATA_SUB   = os.path.join(PROJECT_DIR, "data")
PLOTS_SUB  = os.path.join(PROJECT_DIR, "plots")
LOGS_SUB   = os.path.join(PROJECT_DIR, "logs")
for _d in (PROJECT_DIR, PROJ_SUB, DATA_SUB, PLOTS_SUB, LOGS_SUB):
    os.makedirs(_d, exist_ok=True)

PROJ_NAME = "Monopole_Antennas"
TARGET_F1 = 2.21   # GHz — 低频
TARGET_F2 = 5.8    # GHz — 高频 (双频设计)
S11_TARGET = -10.0  # dB

LOG = os.path.join(LOGS_SUB, "monopole_antennas.log")


# ============================================================================
# 日志 & 工具函数
# ============================================================================
class Logger:
    def __init__(self, p):
        self.t = sys.stdout
        self.f = open(p, "w", encoding="utf-8")

    def write(self, m):
        try:
            self.t.write(m)
        except UnicodeEncodeError:
            self.t.write(m.encode('ascii', 'replace').decode())
        self.f.write(m)
        self.f.flush()

    def flush(self):
        self.t.flush()
        self.f.flush()


sys.stdout = Logger(LOG)
sys.stderr = sys.stdout


def kill_aedt():
    """强制关闭所有 AEDT 进程"""
    try:
        subprocess.run(["taskkill", "/F", "/IM", "ansysedt.exe"],
                       capture_output=True, timeout=15)
        time.sleep(3)
    except Exception:
        pass


def find_bands(freqs, s11_vals, threshold=-10.0):
    """在 S11 曲线中查找所有低于 threshold 的连续频带"""
    below = s11_vals < threshold
    bands = []
    in_band = False
    start = 0
    for i in range(len(below)):
        if below[i] and not in_band:
            start = i
            in_band = True
        elif not below[i] and in_band:
            bands.append((freqs[start], freqs[i - 1]))
            in_band = False
    if in_band:
        bands.append((freqs[start], freqs[-1]))
    return bands


def calc_relative_bw(bands, fc):
    """计算相对带宽 BW = (fH - fL) / fC * 100%，选择包含 fc 的频带"""
    for fL, fH in bands:
        if fL <= fc <= fH:
            bw = (fH - fL) / fc * 100.0
            return fL, fH, bw
    # 若 fc 不在任何频带内，选择最近的
    if bands:
        dists = [min(abs(fc - fL), abs(fc - fH)) for fL, fH in bands]
        idx = int(np.argmin(dists))
        fL, fH = bands[idx]
        bw = (fH - fL) / fc * 100.0
        return fL, fH, bw
    return 0.0, 0.0, 0.0


def read_pattern_csv(path):
    """读取 Far Fields Rectangular Plot CSV，返回 angle(deg) 和 GainTotal(线性)"""
    angles, gains = [], []
    with open(path) as f:
        rd = csv.reader(f)
        next(rd)  # skip header
        for row in rd:
            if len(row) >= 2:
                try:
                    angles.append(float(row[0]))
                    gains.append(float(row[1]))
                except ValueError:
                    pass
    return np.array(angles), np.array(gains)


# ============================================================================
# 主函数
# ============================================================================
def main():
    print("=" * 70)
    print("单频与双频单极子天线设计及数据提取")
    print("=" * 70)

    # ------------------------------------------------------------------
    # 0. 清理旧进程 + PyAEDT 兼容性补丁 (AEDT 2019.1)
    # ------------------------------------------------------------------
    kill_aedt()
    time.sleep(2)

    try:
        from pyaedt import desktop as _dm
        _orig = _dm.Desktop.__init__

        def _patched(self, *a, **k):
            _orig(self, *a, **k)
            self.student_version = getattr(self, 'student_version', False)
        _dm.Desktop.__init__ = _patched
    except Exception:
        pass

    try:
        import pyaedt.application.Design as _dd
        _orig2 = _dd.DesignSettings.__init__

        def _patched2(s, a):
            try:
                _orig2(s, a)
            except AttributeError:
                s._app = a
                s.design_settings = None
                s.manipulate_inputs = None
        _dd.DesignSettings.__init__ = _patched2
    except Exception:
        pass

    from pyaedt import Hfss

    hfss = None
    try:
        # ==============================================================
        # 1. 项目初始化
        # ==============================================================
        print("\n[Step 1] 新建 HFSS 工程 (Driven Modal)...")
        hfss = Hfss(
            projectname=os.path.join(PROJ_SUB, PROJ_NAME),
            designname="Single_Band_Monopole",
            solution_type="DrivenModal",
            non_graphical=True,
            new_desktop_session=True,
            specified_version="2019.1",
        )
        print(f"  项目: {hfss.project_name}")
        hfss.modeler.model_units = "mm"
        print("  单位: mm")

        oProject = hfss.oproject

        # ==============================================================
        # 辅助函数（基于 oEditor 闭包）
        # ==============================================================
        def get_modules(design):
            ed = design.SetActiveEditor("3D Modeler")
            bnd = design.GetModule("BoundarySetup")
            ana = design.GetModule("AnalysisSetup")
            rad = design.GetModule("RadField")
            rpt = design.GetModule("ReportSetup")
            sol = design.GetModule("Solutions")
            return ed, bnd, ana, rad, rpt, sol

        def create_box(editor, name, x, y, z, dx, dy, dz, mat,
                       solve_inside=None):
            if solve_inside is None:
                solve_inside = mat.lower() not in ("copper", "pec", "aluminum")
            editor.CreateBox(
                ["NAME:BoxParameters",
                 "XPosition:=", x, "YPosition:=", y, "ZPosition:=", z,
                 "XSize:=", dx, "YSize:=", dy, "ZSize:=", dz],
                ["NAME:Attributes",
                 "Name:=", name, "Flags:=", "", "Color:=", "(143 175 143)",
                 "Transparency:=", 0.3,
                 "PartCoordinateSystem:=", "Global",
                 "UDMId:=", "",
                 "MaterialValue:=", f'"{mat}"',
                 "SurfaceMaterialValue:=", '""',
                 "SolveInside:=", solve_inside,
                 "IsMaterialEditable:=", True,
                 "UseMaterialAppearance:=", False,
                 "IsLightweight:=", False])

        def create_rect(editor, name, x, y, z, w, h, axis="Z",
                        color="(255 128 0)"):
            editor.CreateRectangle(
                ["NAME:RectangleParameters",
                 "IsCovered:=", True,
                 "XStart:=", x, "YStart:=", y, "ZStart:=", z,
                 "Width:=", w, "Height:=", h,
                 "WhichAxis:=", axis],
                ["NAME:Attributes",
                 "Name:=", name, "Flags:=", "", "Color:=", color,
                 "Transparency:=", 0,
                 "PartCoordinateSystem:=", "Global",
                 "UDMId:=", "",
                 "MaterialValue:=", '"vacuum"',
                 "SurfaceMaterialValue:=", '""',
                 "SolveInside:=", True,
                 "IsMaterialEditable:=", True,
                 "UseMaterialAppearance:=", False,
                 "IsLightweight:=", False])

        def define_variables(design, variables):
            new_props = []
            for name, value in variables.items():
                new_props.append(
                    ["NAME:" + name,
                     "PropType:=", "VariableProp",
                     "UserDef:=", True,
                     "Value:=", value])
            design.ChangeProperty(
                ["NAME:AllTabs",
                 ["NAME:LocalVariableTab",
                  ["NAME:PropServers", "LocalVariables"],
                  ["NAME:NewProps"] + new_props]])

        def assign_lumped_port(bnd_mod, editor, port_obj_name,
                               int_start, int_end):
            """分配 Lumped Port 50Ω，带多种回退方式"""
            try:
                bnd_mod.AssignLumpedPort(
                    ["NAME:Port1",
                     "Objects:=", [port_obj_name],
                     "DoDeembed:=", False,
                     "RenormImp:=", "50ohm",
                     ["NAME:Modes",
                      ["NAME:Mode1",
                       "ModeNum:=", 1,
                       "UseIntLine:=", True,
                       ["NAME:IntLine",
                        "Start:=", int_start,
                        "End:=", int_end]]],
                     "FullResistance:=", "50ohm",
                     "FullReactance:=", "0ohm"])
                print("  Port1: Lumped Port 50ohm (方式1)")
                return
            except Exception as e1:
                print(f"  方式1 失败: {e1}")

            # 回退: 通过面 ID
            try:
                feed_faces = editor.GetFaceIDs(port_obj_name)
                if feed_faces:
                    face_id = int(feed_faces[0])
                    bnd_mod.AssignLumpedPort(
                        ["NAME:Port1",
                         "Faces:=", [face_id],
                         "DoDeembed:=", False,
                         "RenormImp:=", "50ohm",
                         ["NAME:Modes",
                          ["NAME:Mode1",
                           "ModeNum:=", 1,
                           "UseIntLine:=", True,
                           ["NAME:IntLine",
                            "Start:=", int_start,
                            "End:=", int_end]]],
                         "FullResistance:=", "50ohm",
                         "FullReactance:=", "0ohm"])
                    print(f"  Port1: Lumped Port 50ohm (方式2, face={face_id})")
                    return
            except Exception as e2:
                print(f"  方式2 失败: {e2}")

            # 回退: 无积分线
            try:
                bnd_mod.AssignLumpedPort(
                    ["NAME:Port1",
                     "Objects:=", [port_obj_name],
                     "DoDeembed:=", False,
                     "RenormImp:=", "50ohm",
                     ["NAME:Modes",
                      ["NAME:Mode1",
                       "ModeNum:=", 1,
                       "UseIntLine:=", False]],
                     "FullResistance:=", "50ohm",
                     "FullReactance:=", "0ohm"])
                print("  Port1: Lumped Port 50ohm (方式3, 无积分线)")
            except Exception as e3:
                print(f"  方式3 失败: {e3}")
                raise RuntimeError("无法分配 Lumped Port，请检查 Feed_Port 几何")

        # ==============================================================
        # ██  DESIGN 1: Single_Band_Monopole (2.21 GHz)
        # ==============================================================
        print("\n" + "=" * 70)
        print("  DESIGN 1: Single_Band_Monopole (2.21 GHz)")
        print("=" * 70)

        oDesign = oProject.SetActiveDesign("Single_Band_Monopole")
        oEditor, oBnd, oAnalysis, oRadField, oReport, oSolutions = \
            get_modules(oDesign)

        # ---- 2. 定义设计变量 ----
        print("\n[D1-Step 2] 定义设计变量...")
        vars_d1 = {
            "H_sub":  "1.6mm",
            "W_sub":  "40mm",
            "L_sub":  "50mm",
            "L_mono": "25mm",
            "W_mono": "2mm",
        }
        define_variables(oDesign, vars_d1)
        for n, v in vars_d1.items():
            print(f"  {n} = {v}")

        # ---- 3. 几何建模 ----
        print("\n[D1-Step 3] 几何建模...")

        # A. 介质板 Substrate (3D Box)
        print("  A. 创建介质板 Substrate (3D Box, FR4_epoxy)...")
        create_box(oEditor, "Substrate",
                   "0mm", "0mm", "0mm",
                   "W_sub", "L_sub", "H_sub",
                   "FR4_epoxy")
        print("     (0,0,0) -> (W_sub, L_sub, H_sub)")

        # B. 顶层: 馈线 + 单极子辐射臂 (2D Sheet)
        print("  B. 创建顶层辐射贴片 (2D Sheet)...")

        # 馈线段: 从底边 y=0 到地平面边缘 y=L_sub/2
        create_rect(oEditor, "Feed_Strip",
                    "W_sub/2-W_mono/2", "0mm", "H_sub",
                    "W_mono", "L_sub/2")
        print("     Feed_Strip: (W_sub/2-W_mono/2, 0, H_sub) W=W_mono H=L_sub/2")

        # 单极子辐射臂: 从 y=L_sub/2 延伸 L_mono
        create_rect(oEditor, "Monopole_Arm",
                    "W_sub/2-W_mono/2", "L_sub/2", "H_sub",
                    "W_mono", "L_mono")
        print("     Monopole_Arm: (W_sub/2-W_mono/2, L_sub/2, H_sub) W=W_mono H=L_mono")

        # 合并馈线与辐射臂
        oEditor.Unite(
            ["NAME:Selections",
             "Selections:=", "Feed_Strip,Monopole_Arm"],
            ["NAME:UniteParameters", "KeepOriginals:=", False])
        print("     Unite -> Feed_Strip")

        # C. 底层: 截断地平面 DGS (2D Sheet)
        print("  C. 创建底层截断地平面 (2D Sheet)...")
        create_rect(oEditor, "Ground_Plane",
                    "0mm", "0mm", "0mm",
                    "W_sub", "L_sub/2")
        print("     Ground_Plane: (0,0,0) W=W_sub H=L_sub/2 (覆盖下半部分)")

        print("  几何建模完成!")

        # ---- 4. 边界条件 ----
        print("\n[D1-Step 4] 设置边界条件...")

        # Perfect E
        oBnd.AssignPerfectE(
            ["NAME:PerfE_Top",
             "Objects:=", ["Feed_Strip"],
             "InfGroundPlane:=", False])
        print("  Feed_Strip  -> Perfect E")

        oBnd.AssignPerfectE(
            ["NAME:PerfE_Ground",
             "Objects:=", ["Ground_Plane"],
             "InfGroundPlane:=", False])
        print("  Ground_Plane -> Perfect E")

        # AirBox + Radiation Boundary
        create_box(oEditor, "AirBox",
                   "-50mm", "-50mm", "-50mm",
                   "W_sub+100mm", "L_sub+100mm", "H_sub+100mm",
                   "air", True)
        oBnd.AssignRadiation(
            ["NAME:Radiation1",
             "Objects:=", ["AirBox"],
             "IsFssReference:=", False,
             "IsForPML:=", False])
        print("  AirBox -> Radiation Boundary")
        print("    (-50,-50,-50) size(W_sub+100, L_sub+100, H_sub+100)")

        # ---- 5. 激励源 (Lumped Port) ----
        print("\n[D1-Step 5] 设置 Lumped Port...")

        # Port 矩形: 在 y=0, XZ 平面
        create_rect(oEditor, "Feed_Port",
                    "W_sub/2-W_mono/2", "0mm", "0mm",
                    "W_mono", "H_sub",
                    axis="Y", color="(255 0 0)")
        print("  Feed_Port: (W_sub/2-W_mono/2, 0, 0) W=W_mono H=H_sub on XZ plane")

        # 积分线: 从地平面 z=0 到馈线 z=H_sub (数值坐标)
        assign_lumped_port(oBnd, oEditor, "Feed_Port",
                           ["20mm", "0mm", "0mm"],
                           ["20mm", "0mm", "1.6mm"])
        print("  积分线: (20,0,0) -> (20,0,1.6)")

        # ---- 6. 求解设置 ----
        print("\n[D1-Step 6] 求解设置...")

        # Setup: 2.21 GHz, MaxPasses=20
        oAnalysis.InsertSetup("HfssDriven",
            ["NAME:Setup1",
             "Frequency:=", f"{TARGET_F1}GHz",
             "MaxDeltaS:=", 0.02,
             "MaximumPasses:=", 20,
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
        print(f"  Setup1: {TARGET_F1} GHz, MaxPasses=20, DeltaS=0.02")

        # Fast Sweep (Interpolating): 1.5 ~ 3.0 GHz
        oAnalysis.InsertFrequencySweep("Setup1",
            ["NAME:Sweep1",
             "IsEnabled:=", True,
             "SetupType:=", "LinearStep",
             "StartValue:=", "1.5GHz",
             "StopValue:=", "3.0GHz",
             "StepSize:=", "0.01GHz",
             "Type:=", "Fast",
             "SaveFields:=", True,
             "SaveRadFields:=", True])
        print("  Sweep1: Fast 1.5 ~ 3.0 GHz, step=0.01 GHz")

        # Discrete Sweep @ 2.21 GHz (远场数据)
        oAnalysis.InsertFrequencySweep("Setup1",
            ["NAME:Sweep_FF",
             "IsEnabled:=", True,
             "SetupType:=", "LinearCount",
             "StartValue:=", f"{TARGET_F1}GHz",
             "StopValue:=", f"{TARGET_F1}GHz",
             "Count:=", 1,
             "Type:=", "Discrete",
             "SaveFields:=", True,
             "SaveRadFields:=", True])
        print(f"  Sweep_FF: Discrete @ {TARGET_F1} GHz (远场)")

        # Infinite Sphere
        oRadField.InsertFarFieldSphereSetup(
            ["NAME:InfSphere1",
             "UseCustomRadiationSurface:=", False,
             "ThetaStart:=", "0deg",
             "ThetaStop:=", "180deg",
             "ThetaStep:=", "2deg",
             "PhiStart:=", "0deg",
             "PhiStop:=", "360deg",
             "PhiStep:=", "2deg",
             "UseLocalCS:=", False])
        print("  InfSphere1: theta=0~180, phi=0~360, step=2 deg")

        # ---- 7. 验证 & 仿真 ----
        print("\n[D1-Step 7] 设计验证...")
        hfss.save_project()
        v = oDesign.ValidateDesign()
        print(f"  Validation: {v}")
        if v != 1:
            print("  WARNING: 设计验证未通过! 仍尝试仿真...")

        print("\n  运行仿真 (Analyze All)...")
        t0 = time.time()
        oDesign.Analyze("Setup1")
        ts = time.time() - t0
        print(f"  仿真完成! 用时 {ts:.0f}s ({ts / 60:.1f} min)")

        # ---- 8. 后处理 Design 1 ----
        print("\n[D1-Step 8] 后处理 & 数据提取...")

        d1_prefix = "d1_single"
        s11_csv_d1 = os.path.join(DATA_SUB, f"{d1_prefix}_s11.csv")
        s1p_d1 = os.path.join(DATA_SUB, f"{d1_prefix}.s1p")

        # Touchstone
        try:
            oSolutions.ExportNetworkData(
                "", ["Setup1:Sweep1"], 3, s1p_d1, ["All"], True, 50)
            print(f"  Touchstone: {s1p_d1}")
        except Exception as e:
            print(f"  Touchstone 导出失败: {e}")

        # S11 报告
        s_expr = "dB(S(Port1,Port1))"
        try:
            oReport.CreateReport(
                "S11_Plot", "Modal Solution Data", "Rectangular Plot",
                "Setup1 : Sweep1",
                [],
                ["Freq:=", ["All"]],
                ["X Component:=", "Freq",
                 "Y Component:=", [s_expr]],
                [])
            oReport.ExportToFile("S11_Plot", s11_csv_d1)
            print(f"  S11 CSV: {s11_csv_d1}")
            try:
                oReport.ExportImageToFile("S11_Plot",
                    os.path.join(PLOTS_SUB, f"{d1_prefix}_S11_HFSS.jpg"),
                    1920, 1080)
            except Exception:
                pass
        except Exception as e:
            print(f"  S11 报告失败: {e}")
            s_expr = "dB(S(1,1))"
            try:
                oReport.CreateReport(
                    "S11_Plot", "Modal Solution Data", "Rectangular Plot",
                    "Setup1 : Sweep1", [],
                    ["Freq:=", ["All"]],
                    ["X Component:=", "Freq",
                     "Y Component:=", [s_expr]], [])
                oReport.ExportToFile("S11_Plot", s11_csv_d1)
                print(f"  S11 CSV (备用): {s11_csv_d1}")
            except Exception as e2:
                print(f"  S11 备用也失败: {e2}")

        # 解析 S11 数据 & 带宽计算
        freqs_d1, s11_d1 = [], []
        if os.path.exists(s11_csv_d1):
            with open(s11_csv_d1) as fh:
                rd = csv.reader(fh)
                next(rd)
                for row in rd:
                    if len(row) >= 2:
                        try:
                            freqs_d1.append(float(row[0]))
                            s11_d1.append(float(row[1]))
                        except ValueError:
                            pass
            freqs_d1 = np.array(freqs_d1)
            s11_d1 = np.array(s11_d1)

        if len(freqs_d1) > 0:
            mi = np.argmin(s11_d1)
            f_res = freqs_d1[mi]
            s11_min = s11_d1[mi]

            idx_target = np.argmin(np.abs(freqs_d1 - TARGET_F1))
            s11_at_f1 = s11_d1[idx_target]

            bands_d1 = find_bands(freqs_d1, s11_d1, S11_TARGET)
            fL, fH, bw_pct = calc_relative_bw(bands_d1, TARGET_F1)

            print(f"\n  === Design 1 S11 结果 ===")
            print(f"  谐振频率:        {f_res:.4f} GHz")
            print(f"  S11 最小值:      {s11_min:.2f} dB")
            print(f"  S11 @ {TARGET_F1} GHz: {s11_at_f1:.2f} dB")
            if fL > 0 and fH > 0:
                bw_mhz = (fH - fL) * 1000
                print(f"  -10dB 带宽:      {fL:.3f} ~ {fH:.3f} GHz "
                      f"({bw_mhz:.0f} MHz)")
                print(f"  相对带宽 BW:     {bw_pct:.2f}% "
                      f"(fc={TARGET_F1} GHz)")
        else:
            print("  WARNING: 未能解析 S11 数据!")

        # 3D GainTotal (所有 theta, phi)
        print("\n  提取 3D 增益方向图...")
        gain3d_csv_d1 = os.path.join(DATA_SUB, f"{d1_prefix}_gain3d.csv")
        for sweep_ctx in ["Setup1 : Sweep_FF", "Setup1 : LastAdaptive",
                          "Setup1 : Sweep1"]:
            try:
                oReport.CreateReport(
                    "Gain_3D", "Far Fields", "Rectangular Plot",
                    sweep_ctx,
                    ["Context:=", "InfSphere1"],
                    ["Theta:=", ["All"], "Phi:=", ["All"],
                     "Freq:=", [f"{TARGET_F1}GHz"]],
                    ["X Component:=", "Theta",
                     "Y Component:=", ["GainTotal"]],
                    [])
                oReport.ExportToFile("Gain_3D", gain3d_csv_d1)
                print(f"  3D GainTotal: {gain3d_csv_d1} ({sweep_ctx})")
                break
            except Exception as e:
                print(f"  3D ({sweep_ctx}): {e}")

        # 提取峰值增益
        if os.path.exists(gain3d_csv_d1):
            _, g_lin = read_pattern_csv(gain3d_csv_d1)
            if len(g_lin) > 0:
                peak_gain_d1 = 10.0 * np.log10(np.maximum(np.max(g_lin), 1e-12))
                print(f"  ★ Design 1 峰值增益 @ {TARGET_F1} GHz: "
                      f"{peak_gain_d1:.2f} dBi")

        # E/H 面 2D 辐射方向图
        print("\n  提取 2D 辐射方向图 (E面/H面)...")
        pattern_configs = [
            # (报告名, phi/theta 约束, CSV 文件名, X轴变量)
            ("E_Plane_phi0",  {"Theta": "All", "Phi": "0deg"},
             f"{d1_prefix}_e_plane_phi0.csv", "Theta"),
            ("E_Plane_phi90", {"Theta": "All", "Phi": "90deg"},
             f"{d1_prefix}_e_plane_phi90.csv", "Theta"),
            ("H_Plane",       {"Theta": "90deg", "Phi": "All"},
             f"{d1_prefix}_h_plane.csv", "Phi"),
        ]
        for rpt_name, constraints, csv_name, x_comp in pattern_configs:
            csv_path = os.path.join(DATA_SUB, csv_name)
            theta_val = [constraints["Theta"]] if constraints["Theta"] != "All" \
                        else ["All"]
            phi_val = [constraints["Phi"]] if constraints["Phi"] != "All" \
                      else ["All"]
            for sweep_ctx in ["Setup1 : Sweep_FF", "Setup1 : LastAdaptive",
                              "Setup1 : Sweep1"]:
                try:
                    oReport.CreateReport(
                        rpt_name, "Far Fields", "Rectangular Plot",
                        sweep_ctx,
                        ["Context:=", "InfSphere1"],
                        ["Theta:=", theta_val, "Phi:=", phi_val,
                         "Freq:=", [f"{TARGET_F1}GHz"]],
                        ["X Component:=", x_comp,
                         "Y Component:=", ["GainTotal"]],
                        [])
                    oReport.ExportToFile(rpt_name, csv_path)
                    print(f"  {rpt_name}: {csv_path}")
                    try:
                        oReport.ExportImageToFile(rpt_name,
                            os.path.join(PLOTS_SUB,
                                         f"{d1_prefix}_{rpt_name}_HFSS.jpg"),
                            1920, 1080)
                    except Exception:
                        pass
                    break
                except Exception as e:
                    print(f"  {rpt_name} ({sweep_ctx}): {e}")

        # matplotlib 绘图 — Design 1
        try:
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt

            # S11 图
            if len(freqs_d1) > 0:
                fig, ax = plt.subplots(figsize=(10, 6))
                ax.plot(freqs_d1, s11_d1, 'b-', lw=2, label='S11')
                ax.axhline(-10, color='r', ls='--', lw=1, label='-10 dB')
                ax.axvline(TARGET_F1, color='g', ls=':', lw=1.5,
                           label=f'Target {TARGET_F1} GHz')
                ax.axvline(f_res, color='orange', ls='-.', lw=1.5,
                           label=f'Resonance {f_res:.3f} GHz')
                if fL > 0 and fH > 0:
                    ax.axvspan(fL, fH, alpha=0.15, color='green',
                               label=f'BW {(fH-fL)*1e3:.0f} MHz ({bw_pct:.1f}%)')
                ax.set_xlabel('Frequency (GHz)')
                ax.set_ylabel('S11 (dB)')
                ax.set_title(f'Single Band Monopole @ {TARGET_F1} GHz\n'
                             f'f_res={f_res:.4f} GHz, S11_min={s11_min:.2f} dB, '
                             f'BW={bw_pct:.1f}%')
                ax.legend()
                ax.grid(True, alpha=0.3)
                ax.set_xlim(1.5, 3.0)
                s11_png = os.path.join(PLOTS_SUB, f"{d1_prefix}_s11.png")
                fig.savefig(s11_png, dpi=150, bbox_inches='tight')
                plt.close(fig)
                print(f"  S11 图: {s11_png}")

            # 辐射方向图 (极坐标)
            e0_csv = os.path.join(DATA_SUB, f"{d1_prefix}_e_plane_phi0.csv")
            e90_csv = os.path.join(DATA_SUB, f"{d1_prefix}_e_plane_phi90.csv")
            h_csv = os.path.join(DATA_SUB, f"{d1_prefix}_h_plane.csv")

            rad_data = {}
            for label, path in [("E(phi=0)", e0_csv),
                                ("E(phi=90)", e90_csv),
                                ("H(theta=90)", h_csv)]:
                if os.path.exists(path):
                    a, g = read_pattern_csv(path)
                    if len(g) > 0:
                        g_dbi = 10.0 * np.log10(np.maximum(g, 1e-12))
                        rad_data[label] = (a, g_dbi)

            if rad_data:
                n_plots = len(rad_data)
                fig, axes = plt.subplots(1, n_plots,
                    subplot_kw={'projection': 'polar'},
                    figsize=(6 * n_plots, 6))
                if n_plots == 1:
                    axes = [axes]
                colors = ['b', 'r', 'g']
                r_floor = -30
                for idx, (label, (ang, g_dbi)) in enumerate(rad_data.items()):
                    ax = axes[idx]
                    g_plot = np.clip(g_dbi - r_floor, 0, None)
                    ax.plot(np.radians(ang), g_plot, f'{colors[idx]}-', lw=2)
                    ax.fill(np.radians(ang), g_plot, alpha=0.12,
                            color=colors[idx])
                    peak = np.max(g_dbi)
                    ax.set_title(f'{label}\nMax={peak:.2f} dBi', pad=20)
                    ax.set_theta_zero_location('N')
                    ax.set_theta_direction(-1)
                fig.suptitle(f'Single Band Monopole @ {TARGET_F1} GHz\n'
                             f'Radiation Patterns', fontsize=14)
                rad_png = os.path.join(PLOTS_SUB, f"{d1_prefix}_radiation.png")
                fig.savefig(rad_png, dpi=150, bbox_inches='tight')
                plt.close(fig)
                print(f"  辐射方向图: {rad_png}")

                # 打印各面最大增益
                for label, (_, g_dbi) in rad_data.items():
                    print(f"  {label} 最大增益: {np.max(g_dbi):.2f} dBi")

        except Exception as e:
            print(f"  matplotlib 绘图失败: {e}")
            traceback.print_exc()

        hfss.save_project()

        # ==============================================================
        # ██  DESIGN 2: Dual_Band_Monopole (2.21 + 5.8 GHz)
        # ==============================================================
        print("\n" + "=" * 70)
        print("  DESIGN 2: Dual_Band_Monopole (2.21 + 5.8 GHz)")
        print("=" * 70)

        # 插入新设计
        oProject.InsertDesign("HFSS", "Dual_Band_Monopole", "DrivenModal", "")
        oDesign = oProject.SetActiveDesign("Dual_Band_Monopole")
        oEditor, oBnd, oAnalysis, oRadField, oReport, oSolutions = \
            get_modules(oDesign)

        # 设置单位
        oEditor.SetModelUnits(
            ["NAME:Units Parameter", "Units:=", "mm", "Rescale:=", False])
        print("  单位: mm")

        # ---- 定义变量 ----
        print("\n[D2-Step 2] 定义设计变量...")
        vars_d2 = {
            "H_sub":       "1.6mm",
            "W_sub":       "40mm",
            "L_sub":       "50mm",
            "L_mono":      "28mm",
            "W_mono":      "2mm",
            "L_mono_high": "11mm",
            "Y_branch":    "0mm",
        }
        define_variables(oDesign, vars_d2)
        for n, v in vars_d2.items():
            print(f"  {n} = {v}")

        # ---- 几何建模 ----
        print("\n[D2-Step 3] 几何建模 (双分支 F 型)...")

        # A. 介质板
        print("  A. 介质板 Substrate...")
        create_box(oEditor, "Substrate",
                   "0mm", "0mm", "0mm",
                   "W_sub", "L_sub", "H_sub",
                   "FR4_epoxy")

        # B. 顶层: 馈线 + 长分支 + 短分支 (F 型)
        print("  B. 顶层双分支辐射贴片...")

        # B1. 馈线段
        create_rect(oEditor, "Feed_Strip",
                    "W_sub/2-W_mono/2", "0mm", "H_sub",
                    "W_mono", "L_sub/2")
        print("     Feed_Strip: (W_sub/2-W_mono/2, 0, H_sub)")

        # B2. 长分支 (2.21 GHz)
        create_rect(oEditor, "Long_Arm",
                    "W_sub/2-W_mono/2", "L_sub/2", "H_sub",
                    "W_mono", "L_mono")
        print("     Long_Arm:   L_mono (2.21 GHz)")

        # B3. 短分支 (5.8 GHz) — 垂直延伸形成 Y 型双分支
        create_rect(oEditor, "Short_Arm",
                    "W_sub/2+W_mono/2", "L_sub/2+Y_branch", "H_sub",
                    "W_mono", "L_mono_high")
        print("     Short_Arm:  L_mono_high 垂直, Y型双分支 (5.8 GHz)")

        # 合并所有顶层金属
        oEditor.Unite(
            ["NAME:Selections",
             "Selections:=", "Feed_Strip,Long_Arm,Short_Arm"],
            ["NAME:UniteParameters", "KeepOriginals:=", False])
        print("     Unite -> Feed_Strip (包含双分支)")

        # C. 底层: 截断地平面
        print("  C. 截断地平面...")
        create_rect(oEditor, "Ground_Plane",
                    "0mm", "0mm", "0mm",
                    "W_sub", "L_sub/2")
        print("     Ground_Plane: 覆盖下半部分")

        print("  几何建模完成!")

        # ---- 边界条件 ----
        print("\n[D2-Step 4] 设置边界条件...")
        oBnd.AssignPerfectE(
            ["NAME:PerfE_Top",
             "Objects:=", ["Feed_Strip"],
             "InfGroundPlane:=", False])
        print("  Feed_Strip  -> Perfect E")

        oBnd.AssignPerfectE(
            ["NAME:PerfE_Ground",
             "Objects:=", ["Ground_Plane"],
             "InfGroundPlane:=", False])
        print("  Ground_Plane -> Perfect E")

        create_box(oEditor, "AirBox",
                   "-50mm", "-50mm", "-50mm",
                   "W_sub+100mm", "L_sub+100mm", "H_sub+100mm",
                   "air", True)
        oBnd.AssignRadiation(
            ["NAME:Radiation1",
             "Objects:=", ["AirBox"],
             "IsFssReference:=", False,
             "IsForPML:=", False])
        print("  AirBox -> Radiation Boundary")

        # ---- Lumped Port ----
        print("\n[D2-Step 5] 设置 Lumped Port...")
        create_rect(oEditor, "Feed_Port",
                    "W_sub/2-W_mono/2", "0mm", "0mm",
                    "W_mono", "H_sub",
                    axis="Y", color="(255 0 0)")
        assign_lumped_port(oBnd, oEditor, "Feed_Port",
                           ["20mm", "0mm", "0mm"],
                           ["20mm", "0mm", "1.6mm"])

        # ---- 求解设置 ----
        print("\n[D2-Step 6] 求解设置 (双频宽带)...")

        # Setup: 使用最高频率 5.8 GHz 作为中心频率确保网格精度
        oAnalysis.InsertSetup("HfssDriven",
            ["NAME:Setup1",
             "Frequency:=", f"{TARGET_F2}GHz",
             "MaxDeltaS:=", 0.02,
             "MaximumPasses:=", 20,
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
        print(f"  Setup1: {TARGET_F2} GHz (最高频), MaxPasses=20")

        # Fast Sweep (Interpolating): 1.5 ~ 7.0 GHz
        oAnalysis.InsertFrequencySweep("Setup1",
            ["NAME:Sweep1",
             "IsEnabled:=", True,
             "SetupType:=", "LinearStep",
             "StartValue:=", "1.5GHz",
             "StopValue:=", "7.0GHz",
             "StepSize:=", "0.01GHz",
             "Type:=", "Fast",
             "SaveFields:=", True,
             "SaveRadFields:=", True])
        print("  Sweep1: Fast 1.5 ~ 7.0 GHz, step=0.01 GHz")

        # Discrete Sweep @ 2.21 GHz 和 5.8 GHz (远场)
        oAnalysis.InsertFrequencySweep("Setup1",
            ["NAME:Sweep_FF",
             "IsEnabled:=", True,
             "SetupType:=", "LinearCount",
             "StartValue:=", f"{TARGET_F1}GHz",
             "StopValue:=", f"{TARGET_F2}GHz",
             "Count:=", 2,
             "Type:=", "Discrete",
             "SaveFields:=", True,
             "SaveRadFields:=", True])
        print(f"  Sweep_FF: Discrete @ {TARGET_F1}, {TARGET_F2} GHz")

        # Infinite Sphere
        oRadField.InsertFarFieldSphereSetup(
            ["NAME:InfSphere1",
             "UseCustomRadiationSurface:=", False,
             "ThetaStart:=", "0deg",
             "ThetaStop:=", "180deg",
             "ThetaStep:=", "2deg",
             "PhiStart:=", "0deg",
             "PhiStop:=", "360deg",
             "PhiStep:=", "2deg",
             "UseLocalCS:=", False])
        print("  InfSphere1: theta=0~180, phi=0~360, step=2 deg")

        # ---- 验证 & 仿真 ----
        print("\n[D2-Step 7] 设计验证...")
        hfss.save_project()
        v = oDesign.ValidateDesign()
        print(f"  Validation: {v}")
        if v != 1:
            print("  WARNING: 设计验证未通过! 仍尝试仿真...")

        print("\n  运行仿真 (Analyze All)...")
        t0 = time.time()
        oDesign.Analyze("Setup1")
        ts = time.time() - t0
        print(f"  仿真完成! 用时 {ts:.0f}s ({ts / 60:.1f} min)")

        # ---- 后处理 Design 2 ----
        print("\n[D2-Step 8] 后处理 & 数据提取...")

        d2_prefix = "d2_dual"
        s11_csv_d2 = os.path.join(DATA_SUB, f"{d2_prefix}_s11.csv")
        s1p_d2 = os.path.join(DATA_SUB, f"{d2_prefix}.s1p")

        # Touchstone
        try:
            oSolutions.ExportNetworkData(
                "", ["Setup1:Sweep1"], 3, s1p_d2, ["All"], True, 50)
            print(f"  Touchstone: {s1p_d2}")
        except Exception as e:
            print(f"  Touchstone 导出失败: {e}")

        # S11
        s_expr = "dB(S(Port1,Port1))"
        try:
            oReport.CreateReport(
                "S11_Plot", "Modal Solution Data", "Rectangular Plot",
                "Setup1 : Sweep1", [],
                ["Freq:=", ["All"]],
                ["X Component:=", "Freq",
                 "Y Component:=", [s_expr]], [])
            oReport.ExportToFile("S11_Plot", s11_csv_d2)
            print(f"  S11 CSV: {s11_csv_d2}")
            try:
                oReport.ExportImageToFile("S11_Plot",
                    os.path.join(PLOTS_SUB, f"{d2_prefix}_S11_HFSS.jpg"),
                    1920, 1080)
            except Exception:
                pass
        except Exception as e:
            print(f"  S11 报告失败: {e}")
            s_expr = "dB(S(1,1))"
            try:
                oReport.CreateReport(
                    "S11_Plot", "Modal Solution Data", "Rectangular Plot",
                    "Setup1 : Sweep1", [],
                    ["Freq:=", ["All"]],
                    ["X Component:=", "Freq",
                     "Y Component:=", [s_expr]], [])
                oReport.ExportToFile("S11_Plot", s11_csv_d2)
                print(f"  S11 CSV (备用): {s11_csv_d2}")
            except Exception as e2:
                print(f"  S11 备用也失败: {e2}")

        # 解析 S11 & 带宽
        freqs_d2, s11_d2 = [], []
        if os.path.exists(s11_csv_d2):
            with open(s11_csv_d2) as fh:
                rd = csv.reader(fh)
                next(rd)
                for row in rd:
                    if len(row) >= 2:
                        try:
                            freqs_d2.append(float(row[0]))
                            s11_d2.append(float(row[1]))
                        except ValueError:
                            pass
            freqs_d2 = np.array(freqs_d2)
            s11_d2 = np.array(s11_d2)

        if len(freqs_d2) > 0:
            bands_d2 = find_bands(freqs_d2, s11_d2, S11_TARGET)

            print(f"\n  === Design 2 S11 结果 ===")
            mi = np.argmin(s11_d2)
            print(f"  全局 S11 最小值: {s11_d2[mi]:.2f} dB @ "
                  f"{freqs_d2[mi]:.4f} GHz")

            # 低频带 (2.21 GHz)
            idx_f1 = np.argmin(np.abs(freqs_d2 - TARGET_F1))
            s11_at_f1_d2 = s11_d2[idx_f1]
            fL1, fH1, bw1 = calc_relative_bw(bands_d2, TARGET_F1)
            print(f"\n  --- 低频带 ({TARGET_F1} GHz) ---")
            print(f"  S11 @ {TARGET_F1} GHz: {s11_at_f1_d2:.2f} dB")
            if fL1 > 0 and fH1 > 0:
                print(f"  -10dB 带宽: {fL1:.3f} ~ {fH1:.3f} GHz "
                      f"({(fH1-fL1)*1e3:.0f} MHz)")
                print(f"  相对带宽 BW: {bw1:.2f}% (fc={TARGET_F1} GHz)")

            # 高频带 (5.8 GHz)
            idx_f2 = np.argmin(np.abs(freqs_d2 - TARGET_F2))
            s11_at_f2_d2 = s11_d2[idx_f2]
            fL2, fH2, bw2 = calc_relative_bw(bands_d2, TARGET_F2)
            print(f"\n  --- 高频带 ({TARGET_F2} GHz) ---")
            print(f"  S11 @ {TARGET_F2} GHz: {s11_at_f2_d2:.2f} dB")
            if fL2 > 0 and fH2 > 0:
                print(f"  -10dB 带宽: {fL2:.3f} ~ {fH2:.3f} GHz "
                      f"({(fH2-fL2)*1e3:.0f} MHz)")
                print(f"  相对带宽 BW: {bw2:.2f}% (fc={TARGET_F2} GHz)")
        else:
            print("  WARNING: 未能解析 S11 数据!")

        # 远场 — 双频均需提取
        for freq_ghz in [TARGET_F1, TARGET_F2]:
            freq_tag = f"{freq_ghz:.2f}".replace(".", "p")
            print(f"\n  提取远场数据 @ {freq_ghz} GHz...")

            # 3D GainTotal
            gain3d_csv = os.path.join(DATA_SUB,
                                      f"{d2_prefix}_gain3d_{freq_tag}.csv")
            rpt_3d = f"Gain_3D_{freq_tag}"
            for sweep_ctx in ["Setup1 : Sweep_FF", "Setup1 : LastAdaptive",
                              "Setup1 : Sweep1"]:
                try:
                    oReport.CreateReport(
                        rpt_3d, "Far Fields", "Rectangular Plot",
                        sweep_ctx,
                        ["Context:=", "InfSphere1"],
                        ["Theta:=", ["All"], "Phi:=", ["All"],
                         "Freq:=", [f"{freq_ghz}GHz"]],
                        ["X Component:=", "Theta",
                         "Y Component:=", ["GainTotal"]],
                        [])
                    oReport.ExportToFile(rpt_3d, gain3d_csv)
                    print(f"  3D GainTotal: {gain3d_csv}")
                    break
                except Exception as e:
                    print(f"  3D ({sweep_ctx}): {e}")

            # 峰值增益
            if os.path.exists(gain3d_csv):
                _, g_lin = read_pattern_csv(gain3d_csv)
                if len(g_lin) > 0:
                    pk = 10.0 * np.log10(np.maximum(np.max(g_lin), 1e-12))
                    print(f"  ★ 峰值增益 @ {freq_ghz} GHz: {pk:.2f} dBi")

            # E/H 面
            d2_patterns = [
                (f"E_phi0_{freq_tag}",
                 {"Theta": "All", "Phi": "0deg"},
                 f"{d2_prefix}_e_phi0_{freq_tag}.csv", "Theta"),
                (f"E_phi90_{freq_tag}",
                 {"Theta": "All", "Phi": "90deg"},
                 f"{d2_prefix}_e_phi90_{freq_tag}.csv", "Theta"),
                (f"H_Plane_{freq_tag}",
                 {"Theta": "90deg", "Phi": "All"},
                 f"{d2_prefix}_h_plane_{freq_tag}.csv", "Phi"),
            ]
            for rpt_name, constraints, csv_name, x_comp in d2_patterns:
                csv_path = os.path.join(DATA_SUB, csv_name)
                theta_val = [constraints["Theta"]] \
                    if constraints["Theta"] != "All" else ["All"]
                phi_val = [constraints["Phi"]] \
                    if constraints["Phi"] != "All" else ["All"]
                for sweep_ctx in ["Setup1 : Sweep_FF",
                                  "Setup1 : LastAdaptive",
                                  "Setup1 : Sweep1"]:
                    try:
                        oReport.CreateReport(
                            rpt_name, "Far Fields", "Rectangular Plot",
                            sweep_ctx,
                            ["Context:=", "InfSphere1"],
                            ["Theta:=", theta_val, "Phi:=", phi_val,
                             "Freq:=", [f"{freq_ghz}GHz"]],
                            ["X Component:=", x_comp,
                             "Y Component:=", ["GainTotal"]],
                            [])
                        oReport.ExportToFile(rpt_name, csv_path)
                        print(f"  {rpt_name}: {csv_path}")
                        try:
                            oReport.ExportImageToFile(rpt_name,
                                os.path.join(PLOTS_SUB,
                                    f"{d2_prefix}_{rpt_name}_HFSS.jpg"),
                                1920, 1080)
                        except Exception:
                            pass
                        break
                    except Exception as e:
                        print(f"  {rpt_name} ({sweep_ctx}): {e}")

        # matplotlib 绘图 — Design 2
        try:
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt

            # S11 双频
            if len(freqs_d2) > 0:
                fig, ax = plt.subplots(figsize=(12, 6))
                ax.plot(freqs_d2, s11_d2, 'b-', lw=2, label='S11')
                ax.axhline(-10, color='r', ls='--', lw=1, label='-10 dB')
                ax.axvline(TARGET_F1, color='g', ls=':', lw=1.5,
                           label=f'f1={TARGET_F1} GHz')
                ax.axvline(TARGET_F2, color='m', ls=':', lw=1.5,
                           label=f'f2={TARGET_F2} GHz')
                # 着色各频带
                for fL, fH in bands_d2:
                    ax.axvspan(fL, fH, alpha=0.12, color='green')
                title_parts = [f'Dual Band Monopole: f1={TARGET_F1}, '
                               f'f2={TARGET_F2} GHz']
                if fL1 > 0:
                    title_parts.append(f'BW1={bw1:.1f}%')
                if fL2 > 0:
                    title_parts.append(f'BW2={bw2:.1f}%')
                ax.set_xlabel('Frequency (GHz)')
                ax.set_ylabel('S11 (dB)')
                ax.set_title(' | '.join(title_parts))
                ax.legend()
                ax.grid(True, alpha=0.3)
                ax.set_xlim(1.5, 7.0)
                s11_png_d2 = os.path.join(PLOTS_SUB, f"{d2_prefix}_s11.png")
                fig.savefig(s11_png_d2, dpi=150, bbox_inches='tight')
                plt.close(fig)
                print(f"  S11 图: {s11_png_d2}")

            # 辐射方向图 — 双频各频点
            for freq_ghz in [TARGET_F1, TARGET_F2]:
                freq_tag = f"{freq_ghz:.2f}".replace(".", "p")
                e0_p = os.path.join(DATA_SUB,
                                    f"{d2_prefix}_e_phi0_{freq_tag}.csv")
                e90_p = os.path.join(DATA_SUB,
                                     f"{d2_prefix}_e_phi90_{freq_tag}.csv")
                h_p = os.path.join(DATA_SUB,
                                   f"{d2_prefix}_h_plane_{freq_tag}.csv")

                rad_data = {}
                for label, path in [("E(phi=0)", e0_p),
                                    ("E(phi=90)", e90_p),
                                    ("H(theta=90)", h_p)]:
                    if os.path.exists(path):
                        a, g = read_pattern_csv(path)
                        if len(g) > 0:
                            g_dbi = 10.0 * np.log10(np.maximum(g, 1e-12))
                            rad_data[label] = (a, g_dbi)

                if rad_data:
                    n_plots = len(rad_data)
                    fig, axes = plt.subplots(1, n_plots,
                        subplot_kw={'projection': 'polar'},
                        figsize=(6 * n_plots, 6))
                    if n_plots == 1:
                        axes = [axes]
                    colors = ['b', 'r', 'g']
                    r_floor = -30
                    for idx, (label, (ang, g_dbi)) in enumerate(
                            rad_data.items()):
                        ax = axes[idx]
                        g_plot = np.clip(g_dbi - r_floor, 0, None)
                        ax.plot(np.radians(ang), g_plot,
                                f'{colors[idx]}-', lw=2)
                        ax.fill(np.radians(ang), g_plot, alpha=0.12,
                                color=colors[idx])
                        peak = np.max(g_dbi)
                        ax.set_title(f'{label}\nMax={peak:.2f} dBi', pad=20)
                        ax.set_theta_zero_location('N')
                        ax.set_theta_direction(-1)
                    fig.suptitle(
                        f'Dual Band Monopole @ {freq_ghz} GHz\n'
                        f'Radiation Patterns', fontsize=14)
                    rad_png = os.path.join(PLOTS_SUB,
                        f"{d2_prefix}_radiation_{freq_tag}.png")
                    fig.savefig(rad_png, dpi=150, bbox_inches='tight')
                    plt.close(fig)
                    print(f"  辐射方向图 @ {freq_ghz} GHz: {rad_png}")

                    for label, (_, g_dbi) in rad_data.items():
                        print(f"    {label} 最大增益: "
                              f"{np.max(g_dbi):.2f} dBi")

        except Exception as e:
            print(f"  matplotlib 绘图失败: {e}")
            traceback.print_exc()

        # ==============================================================
        # 导出 3D 模型截图
        # ==============================================================
        for dname in ["Single_Band_Monopole", "Dual_Band_Monopole"]:
            try:
                oD = oProject.SetActiveDesign(dname)
                oE = oD.SetActiveEditor("3D Modeler")
                oE.FitAll()
                img = os.path.join(PLOTS_SUB, f"{dname}_3D.jpg")
                oE.ExportModelImageToFile(img, 1920, 1080,
                    ["NAME:SaveImageParams",
                     "ShowAxis:=", "True",
                     "ShowGrid:=", "True",
                     "ShowRuler:=", "True"])
                print(f"  3D 截图: {img}")
            except Exception:
                try:
                    oD.ExportImage(img, 1920, 1080)
                    print(f"  3D 截图(备用): {img}")
                except Exception as e2:
                    print(f"  3D 截图失败 ({dname}): {e2}")

        # ==============================================================
        # 最终保存 & 汇总
        # ==============================================================
        hfss.save_project()
        proj_path = os.path.join(PROJ_SUB, PROJ_NAME + ".aedt")

        print("\n" + "=" * 70)
        print("全部设计完成! 结果汇总")
        print("=" * 70)

        print(f"\n  项目文件:     {proj_path}")
        print(f"  输出目录:     {PROJECT_DIR}")
        print(f"  日志文件:     {LOG}")

        print(f"\n  --- Design 1: Single_Band_Monopole ---")
        print(f"  中心频率:     {TARGET_F1} GHz")
        print(f"  介质板:       FR4_epoxy (H=1.6mm)")
        print(f"  求解类型:     Driven Modal")
        if len(freqs_d1) > 0:
            print(f"  谐振频率:     {f_res:.4f} GHz")
            print(f"  S11 最小值:   {s11_min:.2f} dB")
            if fL > 0:
                print(f"  -10dB 带宽:   {fL:.3f}~{fH:.3f} GHz ({bw_pct:.1f}%)")

        print(f"\n  --- Design 2: Dual_Band_Monopole ---")
        print(f"  目标频率:     {TARGET_F1} + {TARGET_F2} GHz")
        if len(freqs_d2) > 0:
            if fL1 > 0:
                print(f"  低频带 BW:    {fL1:.3f}~{fH1:.3f} GHz ({bw1:.1f}%)")
            if fL2 > 0:
                print(f"  高频带 BW:    {fL2:.3f}~{fH2:.3f} GHz ({bw2:.1f}%)")

        # 写入设计摘要文件
        summary_path = os.path.join(PROJECT_DIR, "monopole_design_summary.txt")
        with open(summary_path, "w", encoding="utf-8") as sf:
            sf.write("单频与双频单极子天线设计摘要\n")
            sf.write("=" * 50 + "\n\n")

            sf.write("Design 1: Single_Band_Monopole\n")
            sf.write(f"  中心频率: {TARGET_F1} GHz\n")
            sf.write("  设计变量:\n")
            for n, v in vars_d1.items():
                sf.write(f"    {n} = {v}\n")
            if len(freqs_d1) > 0:
                sf.write(f"  谐振频率: {f_res:.4f} GHz\n")
                sf.write(f"  S11 最小值: {s11_min:.2f} dB\n")
                if fL > 0:
                    sf.write(f"  -10dB BW: {fL:.3f}~{fH:.3f} GHz "
                             f"({bw_pct:.1f}%)\n")

            sf.write(f"\nDesign 2: Dual_Band_Monopole\n")
            sf.write(f"  目标频率: {TARGET_F1} + {TARGET_F2} GHz\n")
            sf.write("  设计变量:\n")
            for n, v in vars_d2.items():
                sf.write(f"    {n} = {v}\n")
            if len(freqs_d2) > 0:
                if fL1 > 0:
                    sf.write(f"  低频带 BW: {fL1:.3f}~{fH1:.3f} GHz "
                             f"({bw1:.1f}%)\n")
                if fL2 > 0:
                    sf.write(f"  高频带 BW: {fL2:.3f}~{fH2:.3f} GHz "
                             f"({bw2:.1f}%)\n")

            sf.write(f"\n项目文件: {proj_path}\n")
            sf.write(f"日志文件: {LOG}\n")
        print(f"  设计摘要:     {summary_path}")

        hfss.release_desktop(close_projects=True)

    except Exception as e:
        print(f"\n  致命错误: {e}")
        traceback.print_exc()
        try:
            if hfss is not None:
                hfss.release_desktop(close_projects=True)
        except Exception:
            pass

    kill_aedt()
    gc.collect()
    print("\nDONE")


if __name__ == "__main__":
    main()
