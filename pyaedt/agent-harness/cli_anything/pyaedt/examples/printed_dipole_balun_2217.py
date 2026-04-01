#!/usr/bin/env python3
"""
printed_dipole_balun_2217.py
印刷偶极子天线（带微带巴伦馈线）自动建模 - 2.217 GHz

使用 cli-anything-pyaedt 在 HFSS 中自动完成天线设计：
  - 中心频率: 2.217 GHz
  - 介质板: FR4 (εr = 4.4, H = 1.6mm)
  - 馈电方式: 微带巴伦 + 集总端口 (Lumped Port, 50Ω)
  - 求解类型: Driven Modal
  - 含 Optimetrics 参数扫描优化 (L2, L3)
"""

import os, sys, time, csv, traceback, subprocess, gc
import numpy as np

# ============================================================================
# 全局配置
# ============================================================================
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
os.makedirs(PROJECT_DIR, exist_ok=True)

PROJ_NAME = "Printed_Dipole_Balun_2217"
DESIGN_NAME = "Dipole_Balun"
TARGET_F = 2.217   # GHz
S11_TARGET = -10.0  # dB

LOG = os.path.join(PROJECT_DIR, "dipole_balun_2217.log")


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


def main():
    print("=" * 70)
    print("印刷偶极子天线（带微带巴伦馈线）- 2.217 GHz 自动建模")
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
        # 1. 项目初始化与基本设置
        # ==============================================================
        print("\n[Step 1] 新建 HFSS 工程 (Driven Modal)...")
        hfss = Hfss(
            projectname=os.path.join(PROJECT_DIR, PROJ_NAME),
            designname=DESIGN_NAME,
            solution_type="DrivenModal",
            non_graphical=True,
            new_desktop_session=True,
            specified_version="2019.1",
        )
        print(f"  项目: {hfss.project_name}")
        print(f"  设计: {hfss.design_name}")

        # 设置单位为 mm
        hfss.modeler.model_units = "mm"
        print("  单位: mm")

        # 获取 COM 模块句柄
        oDesign = hfss.odesign
        oEditor = oDesign.SetActiveEditor("3D Modeler")
        oBnd = oDesign.GetModule("BoundarySetup")
        oAnalysis = oDesign.GetModule("AnalysisSetup")
        oRadField = oDesign.GetModule("RadField")
        oReport = oDesign.GetModule("ReportSetup")
        oSolutions = oDesign.GetModule("Solutions")

        # ==============================================================
        # 2. 定义设计变量
        # ==============================================================
        print("\n[Step 2] 定义设计变量...")
        variables = {
            "H":  "1.6mm",   # 介质层厚度
            "W1": "3mm",     # 微带传输线宽度
            "L1": "22mm",    # 微带传输线长度
            "W2": "3mm",     # 偶极子金属片宽度
            "L2": "24mm",    # 偶极子单臂长度 (2.217GHz 初始值, 原 2.45GHz 为 21mm)
            "L3": "10mm",    # 巴伦三角形侧边直角边长
            "L4": "12mm",    # 巴伦三角形底边直角边长
            "W3": "3mm",     # 微波巴伦长方形部分宽度
        }

        new_props = []
        for name, value in variables.items():
            new_props.append(
                ["NAME:" + name,
                 "PropType:=", "VariableProp",
                 "UserDef:=", True,
                 "Value:=", value])
        oDesign.ChangeProperty(
            ["NAME:AllTabs",
             ["NAME:LocalVariableTab",
              ["NAME:PropServers", "LocalVariables"],
              ["NAME:NewProps"] + new_props]])

        for name, value in variables.items():
            print(f"  {name} = {value}")

        # ==============================================================
        # 3. 几何建模
        # ==============================================================
        print("\n[Step 3] 几何建模...")

        # ---------- 辅助函数 ----------
        def create_box(name, x, y, z, dx, dy, dz, mat, solve_inside=None):
            if solve_inside is None:
                solve_inside = mat.lower() not in ("copper", "pec", "aluminum")
            oEditor.CreateBox(
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

        def create_rect(name, x, y, z, w, h, axis="Z", color="(255 128 0)"):
            oEditor.CreateRectangle(
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

        def create_triangle(name, pts, color="(255 128 0)"):
            """创建三角形面片。pts = [(x1,y1,z1), (x2,y2,z2), (x3,y3,z3)]，各坐标为 AEDT 表达式字符串"""
            oEditor.CreatePolyline(
                ["NAME:PolylineParameters",
                 "IsPolylineCovered:=", True,
                 "IsPolylineClosed:=", True,
                 ["NAME:PolylinePoints",
                  ["NAME:PLPoint", "X:=", pts[0][0], "Y:=", pts[0][1], "Z:=", pts[0][2]],
                  ["NAME:PLPoint", "X:=", pts[1][0], "Y:=", pts[1][1], "Z:=", pts[1][2]],
                  ["NAME:PLPoint", "X:=", pts[2][0], "Y:=", pts[2][1], "Z:=", pts[2][2]]],
                 ["NAME:PolylineSegments",
                  ["NAME:PLSegment", "SegmentType:=", "Line",
                   "StartIndex:=", 0, "NoOfPoints:=", 2],
                  ["NAME:PLSegment", "SegmentType:=", "Line",
                   "StartIndex:=", 1, "NoOfPoints:=", 2],
                  ["NAME:PLSegment", "SegmentType:=", "Line",
                   "StartIndex:=", 2, "NoOfPoints:=", 2]]],
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

        # ---- A. 介质层 (Substrate) ----
        print("  A. 创建介质层 Substrate...")
        create_box("Substrate",
                    "-40mm", "-30mm", "0mm",
                    "40mm", "60mm", "H",
                    "FR4_epoxy")
        print("     Substrate: (-40,-30,0) -> size(40, 60, H), FR4_epoxy")

        # ---- B. 顶层辐射臂与传输线 (Top Layer) ----
        print("  B. 创建顶层 Top Layer...")

        # 主传输线 Top_Patch: 起点 (0, -W1/2, H), 尺寸 -L1(X), W1(Y)
        create_rect("Top_Patch",
                     "0mm", "-W1/2", "H",
                     "-L1", "W1")
        print("     Top_Patch:  (0, -W1/2, H) -> (-L1, W1)")

        # 右侧天线臂 Dip_Patch: 起点 (-L1, -W1/2, H), 尺寸 -W2(X), -L2(Y)
        create_rect("Dip_Patch",
                     "-L1", "-W1/2", "H",
                     "-W2", "-L2")
        print("     Dip_Patch:  (-L1, -W1/2, H) -> (-W2, -L2)")

        # 连接斜切角 Polyline1 (三角形面)
        create_triangle("Polyline1", [
            ("-L1",     "-W1/2", "H"),
            ("-L1",     "W1/2",  "H"),
            ("-L1-W2",  "-W1/2", "H"),
        ])
        print("     Polyline1:  三角形连接面")

        # 布尔合并: Top_Patch + Dip_Patch + Polyline1 → Top_Patch
        oEditor.Unite(
            ["NAME:Selections",
             "Selections:=", "Top_Patch,Dip_Patch,Polyline1"],
            ["NAME:UniteParameters", "KeepOriginals:=", False])
        print("     Unite -> Top_Patch")

        # ---- C. 底层辐射臂与微带巴伦 (Bottom Layer) ----
        print("  C. 创建底层 Bottom Layer...")

        # 沿 xoz 面 (y=0 平面) 镜像复制 Top_Patch → Top_Patch_1
        oEditor.DuplicateMirror(
            ["NAME:Selections",
             "Selections:=", "Top_Patch",
             "NewPartsModelFlag:=", "Model"],
            ["NAME:DuplicateToMirrorParameters",
             "DuplicateMirrorBaseX:=", "0mm",
             "DuplicateMirrorBaseY:=", "0mm",
             "DuplicateMirrorBaseZ:=", "0mm",
             "DuplicateMirrorNormalX:=", "0",
             "DuplicateMirrorNormalY:=", "1",
             "DuplicateMirrorNormalZ:=", "0"],
            ["NAME:Options", "DuplicateAssignments:=", False])
        print("     DuplicateMirror Top_Patch -> Top_Patch_1")

        # 沿 Z 轴向下平移 -H，使底层位于 z=0
        oEditor.Move(
            ["NAME:Selections",
             "Selections:=", "Top_Patch_1",
             "NewPartsModelFlag:=", "Model"],
            ["NAME:TranslateParameters",
             "TranslateVectorX:=", "0mm",
             "TranslateVectorY:=", "0mm",
             "TranslateVectorZ:=", "-H"])
        print("     Move Top_Patch_1: Z -= H")

        # 巴伦结构右侧 Rectangle1: 起点 (0, W1/2, 0), 尺寸 -W3(X), L3(Y)
        create_rect("Rectangle1",
                     "0mm", "W1/2", "0mm",
                     "-W3", "L3")
        print("     Rectangle1: (0, W1/2, 0) -> (-W3, L3)")

        # 巴伦结构右侧三角形 Polyline2
        create_triangle("Polyline2", [
            ("-W3",      "W1/2",      "0mm"),
            ("-W3",      "W1/2+L3",   "0mm"),
            ("-W3-L4",   "W1/2",      "0mm"),
        ])
        print("     Polyline2: 巴伦三角形右侧")

        # 镜像复制 Rectangle1 + Polyline2 到左侧 (关于 xoz 面)
        oEditor.DuplicateMirror(
            ["NAME:Selections",
             "Selections:=", "Rectangle1,Polyline2",
             "NewPartsModelFlag:=", "Model"],
            ["NAME:DuplicateToMirrorParameters",
             "DuplicateMirrorBaseX:=", "0mm",
             "DuplicateMirrorBaseY:=", "0mm",
             "DuplicateMirrorBaseZ:=", "0mm",
             "DuplicateMirrorNormalX:=", "0",
             "DuplicateMirrorNormalY:=", "1",
             "DuplicateMirrorNormalZ:=", "0"],
            ["NAME:Options", "DuplicateAssignments:=", False])
        print("     DuplicateMirror Rectangle1,Polyline2 -> Rectangle1_1,Polyline2_1")

        # 布尔合并底层所有 5 个金属面 → Top_Patch_1
        oEditor.Unite(
            ["NAME:Selections",
             "Selections:=",
             "Top_Patch_1,Rectangle1,Polyline2,Rectangle1_1,Polyline2_1"],
            ["NAME:UniteParameters", "KeepOriginals:=", False])
        print("     Unite 5 faces -> Top_Patch_1")

        print("  几何建模完成!")

        # ==============================================================
        # 4. 边界条件设置
        # ==============================================================
        print("\n[Step 4] 设置边界条件...")

        # 理想导体面 (Perfect E)
        oBnd.AssignPerfectE(
            ["NAME:PerfE_Top",
             "Objects:=", ["Top_Patch"],
             "InfGroundPlane:=", False])
        print("  Top_Patch   -> Perfect E")

        oBnd.AssignPerfectE(
            ["NAME:PerfE_Bottom",
             "Objects:=", ["Top_Patch_1"],
             "InfGroundPlane:=", False])
        print("  Top_Patch_1 -> Perfect E")

        # 辐射边界 AirBox
        create_box("AirBox",
                    "-100mm", "-90mm", "-60mm",
                    "160mm", "180mm", "120mm",
                    "air", True)
        oBnd.AssignRadiation(
            ["NAME:Radiation1",
             "Objects:=", ["AirBox"],
             "IsFssReference:=", False,
             "IsForPML:=", False])
        print("  AirBox      -> Radiation Boundary")
        print("               (-100,-90,-60) size(160,180,120)")

        # ==============================================================
        # 5. 激励源设置 (Lumped Port)
        # ==============================================================
        print("\n[Step 5] 设置集总端口激励 (Lumped Port)...")

        # 在 YZ 面创建 Feed_Port 矩形
        # 起点 (0, -W1/2, H), Y 向尺寸 W1, Z 向尺寸 -H
        create_rect("Feed_Port",
                     "0mm", "-W1/2", "H",
                     "W1", "-H",
                     axis="X", color="(255 0 0)")
        print("  Feed_Port: (0, -W1/2, H), size(W1, -H) on YZ plane")

        # 分配 Lumped Port，积分线从下边缘中点 → 上边缘中点
        # AEDT 2019.1 API: 使用数值坐标，精简参数格式
        sub_h_val = 1.6  # mm — H 的数值，用于积分线坐标
        try:
            oBnd.AssignLumpedPort(
                ["NAME:Port1",
                 "Objects:=", ["Feed_Port"],
                 "DoDeembed:=", False,
                 "RenormImp:=", "50ohm",
                 ["NAME:Modes",
                  ["NAME:Mode1",
                   "ModeNum:=", 1,
                   "UseIntLine:=", True,
                   ["NAME:IntLine",
                    "Start:=", ["0mm", "0mm", "0mm"],
                    "End:=", ["0mm", "0mm", f"{sub_h_val}mm"]]]],
                 "FullResistance:=", "50ohm",
                 "FullReactance:=", "0ohm"])
            print("  Port1: Lumped Port 50Ω (方式1)")
        except Exception as e1:
            print(f"  Lumped Port 方式1 失败: {e1}")
            print("  尝试 Lumped Port 方式2 (面选取)...")
            # 备选方案: 通过面 ID 分配
            feed_faces = oEditor.GetFaceIDs("Feed_Port")
            if feed_faces:
                face_id = int(feed_faces[0])
                try:
                    oBnd.AssignLumpedPort(
                        ["NAME:Port1",
                         "Faces:=", [face_id],
                         "DoDeembed:=", False,
                         "RenormImp:=", "50ohm",
                         ["NAME:Modes",
                          ["NAME:Mode1",
                           "ModeNum:=", 1,
                           "UseIntLine:=", True,
                           ["NAME:IntLine",
                            "Start:=", ["0mm", "0mm", "0mm"],
                            "End:=", ["0mm", "0mm", f"{sub_h_val}mm"]]]],
                         "FullResistance:=", "50ohm",
                         "FullReactance:=", "0ohm"])
                    print(f"  Port1: Lumped Port 50Ω (方式2, face={face_id})")
                except Exception as e2:
                    print(f"  Lumped Port 方式2 也失败: {e2}")
                    print("  尝试 Lumped Port 方式3 (无积分线)...")
                    try:
                        oBnd.AssignLumpedPort(
                            ["NAME:Port1",
                             "Objects:=", ["Feed_Port"],
                             "DoDeembed:=", False,
                             "RenormImp:=", "50ohm",
                             ["NAME:Modes",
                              ["NAME:Mode1",
                               "ModeNum:=", 1,
                               "UseIntLine:=", False]],
                             "FullResistance:=", "50ohm",
                             "FullReactance:=", "0ohm"])
                        print("  Port1: Lumped Port 50Ω (方式3, 无积分线)")
                    except Exception as e3:
                        print(f"  Lumped Port 方式3 失败: {e3}")
                        raise RuntimeError("无法分配 Lumped Port，请检查 Feed_Port 几何")
        print(f"  积分线: (0,0,0) -> (0,0,{sub_h_val}mm)")

        # ==============================================================
        # 6. 求解设置与后处理
        # ==============================================================
        print("\n[Step 6] 配置求解器...")

        # 6a. Solution Setup: 2.217 GHz, MaxPasses=20, DeltaS=0.02
        oAnalysis.InsertSetup("HfssDriven",
            ["NAME:Setup1",
             "Frequency:=", f"{TARGET_F}GHz",
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
        print(f"  Setup1: {TARGET_F} GHz, MaxPasses=20, DeltaS=0.02")

        # 6b. Fast Sweep: 1.5 ~ 3.0 GHz, step 0.01 GHz
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
        print("  Sweep1: Fast Sweep 1.5 ~ 3.0 GHz, step=0.01 GHz")

        # 6b2. Discrete Sweep (单点，专用于远场辐射数据)
        oAnalysis.InsertFrequencySweep("Setup1",
            ["NAME:Sweep2",
             "IsEnabled:=", True,
             "SetupType:=", "LinearCount",
             "StartValue:=", f"{TARGET_F}GHz",
             "StopValue:=", f"{TARGET_F}GHz",
             "Count:=", 1,
             "Type:=", "Discrete",
             "SaveFields:=", True,
             "SaveRadFields:=", True])
        print(f"  Sweep2: Discrete @ {TARGET_F} GHz (远场数据)")

        # 6c. Infinite Sphere 远场设置
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

        # ==============================================================
        # 6d. 设计验证与仿真
        # ==============================================================
        print("\n[Step 6d] 设计验证...")
        hfss.save_project()
        v = oDesign.ValidateDesign()
        print(f"  Validation: {v}")

        if v != 1:
            print("  WARNING: 设计验证未通过! 仍尝试仿真...")

        print("\n[Step 6e] 运行仿真 (Analyze All)...")
        t0 = time.time()
        oDesign.Analyze("Setup1")
        ts = time.time() - t0
        print(f"  仿真完成! 用时 {ts:.0f}s ({ts / 60:.1f} min)")

        # ==============================================================
        # 6f. 提取结果
        # ==============================================================
        print("\n[Step 6f] 提取仿真结果...")

        s11_csv = os.path.join(PROJECT_DIR, "s11_balun_2217.csv")
        s1p_file = os.path.join(PROJECT_DIR, "dipole_balun_2217.s1p")

        # Touchstone 导出
        try:
            oSolutions.ExportNetworkData(
                "", ["Setup1:Sweep1"], 3, s1p_file, ["All"], True, 50)
            print(f"  Touchstone: {s1p_file}")
        except Exception as e:
            print(f"  Touchstone 导出失败: {e}")

        # S11 回波损耗报告
        # Driven Modal 模式下 S 参数格式: S(Port1,Port1)
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
            oReport.ExportToFile("S11_Plot", s11_csv)
            print(f"  S11 CSV: {s11_csv}")
            # 导出 HFSS S11 报告图片
            s11_img = os.path.join(PROJECT_DIR, "S11_Plot_HFSS.jpg")
            try:
                oReport.ExportImageToFile("S11_Plot", s11_img, 1920, 1080)
                print(f"  S11 HFSS图片: {s11_img}")
            except Exception as ei:
                print(f"  S11 图片导出失败: {ei}")
        except Exception as e:
            print(f"  S11 报告创建失败 (尝试备用格式): {e}")
            # 备用: S(1,1)
            s_expr = "dB(S(1,1))"
            try:
                oReport.CreateReport(
                    "S11_Plot", "Modal Solution Data", "Rectangular Plot",
                    "Setup1 : Sweep1",
                    [],
                    ["Freq:=", ["All"]],
                    ["X Component:=", "Freq",
                     "Y Component:=", [s_expr]],
                    [])
                oReport.ExportToFile("S11_Plot", s11_csv)
                print(f"  S11 CSV (备用格式): {s11_csv}")
                # 导出 HFSS S11 报告图片 (备用)
                s11_img = os.path.join(PROJECT_DIR, "S11_Plot_HFSS.jpg")
                try:
                    oReport.ExportImageToFile("S11_Plot", s11_img, 1920, 1080)
                    print(f"  S11 HFSS图片: {s11_img}")
                except Exception as ei:
                    print(f"  S11 图片导出失败: {ei}")
            except Exception as e2:
                print(f"  S11 报告仍然失败: {e2}")

        # 解析 S11 数据
        freqs, s11_vals = [], []
        if os.path.exists(s11_csv):
            with open(s11_csv) as fh:
                rd = csv.reader(fh)
                next(rd)  # 跳过标题行
                for row in rd:
                    if len(row) >= 2:
                        try:
                            freqs.append(float(row[0]))
                            s11_vals.append(float(row[1]))
                        except ValueError:
                            pass
            freqs = np.array(freqs)
            s11_vals = np.array(s11_vals)

        f_res = s11_min = s11_at_target = 0.0
        bw_lo = bw_hi = bw = bw_pct = 0.0
        freq_err = 999.0

        if len(freqs) > 0:
            mi = np.argmin(s11_vals)
            f_res = freqs[mi]
            s11_min = s11_vals[mi]

            # -10 dB 带宽
            b10 = np.where(s11_vals < -10)[0]
            if len(b10) >= 2:
                bw_lo = freqs[b10[0]]
                bw_hi = freqs[b10[-1]]
                bw = bw_hi - bw_lo
                bw_pct = bw / f_res * 100
            else:
                bw_lo = bw_hi = f_res

            freq_err = abs(f_res - TARGET_F) / TARGET_F * 100

            # 查找 2.217 GHz 处 S11
            idx_target = np.argmin(np.abs(freqs - TARGET_F))
            s11_at_target = s11_vals[idx_target]

            print(f"\n  === S11 结果 ===")
            print(f"  谐振频率:       {f_res:.4f} GHz")
            print(f"  S11 最小值:     {s11_min:.2f} dB")
            print(f"  S11 @ {TARGET_F} GHz: {s11_at_target:.2f} dB")
            if bw > 0:
                print(f"  -10dB 带宽:     {bw_lo:.3f} ~ {bw_hi:.3f} GHz"
                      f" ({bw * 1e3:.0f} MHz, {bw_pct:.2f}%)")
            print(f"  频率误差:       {freq_err:.2f}%")
        else:
            print("  WARNING: 未能解析 S11 数据!")

        # 3D 增益方向图 (GainTotal) — 使用 Rectangular Plot (AEDT 2019.1 兼容)
        print("\n  提取 3D 增益方向图...")
        for sweep_ctx in ["Setup1 : Sweep2", "Setup1 : LastAdaptive", "Setup1 : Sweep1"]:
            try:
                oReport.CreateReport(
                    "Gain_3D", "Far Fields", "Rectangular Plot",
                    sweep_ctx,
                    ["Context:=", "InfSphere1"],
                    ["Theta:=", ["All"], "Phi:=", ["All"],
                     "Freq:=", [f"{TARGET_F}GHz"]],
                    ["X Component:=", "Theta",
                     "Y Component:=", ["GainTotal"]],
                    [])
                gain3d_csv = os.path.join(PROJECT_DIR, "gain_3d_2217.csv")
                oReport.ExportToFile("Gain_3D", gain3d_csv)
                print(f"  3D GainTotal 导出: {gain3d_csv} ({sweep_ctx})")
                # 导出 HFSS 3D 增益方向图图片
                gain3d_img = os.path.join(PROJECT_DIR, "Gain_3D_HFSS.jpg")
                try:
                    oReport.ExportImageToFile("Gain_3D", gain3d_img, 1920, 1080)
                    print(f"  3D增益 HFSS图片: {gain3d_img}")
                except Exception as ei:
                    print(f"  3D增益图片导出失败: {ei}")
                break
            except Exception as e:
                print(f"  3D 方向图 ({sweep_ctx}): {e}")

        # E/H 面 2D 辐射方向图
        for pn, phi_val, csv_out in [
            ("E_Plane", "0deg",  os.path.join(PROJECT_DIR, "e_plane_2217.csv")),
            ("H_Plane", "90deg", os.path.join(PROJECT_DIR, "h_plane_2217.csv")),
        ]:
            for sweep_ctx in ["Setup1 : Sweep2", "Setup1 : LastAdaptive", "Setup1 : Sweep1"]:
                try:
                    oReport.CreateReport(
                        pn, "Far Fields", "Rectangular Plot",
                        sweep_ctx,
                        ["Context:=", "InfSphere1"],
                        ["Theta:=", ["All"], "Phi:=", [phi_val],
                         "Freq:=", [f"{TARGET_F}GHz"]],
                        ["X Component:=", "Theta",
                         "Y Component:=", ["GainTotal"]],
                        [])
                    oReport.ExportToFile(pn, csv_out)
                    print(f"  {pn} 导出: {csv_out}")
                    # 导出 HFSS E/H面报告图片
                    pat_img = os.path.join(PROJECT_DIR, f"{pn}_HFSS.jpg")
                    try:
                        oReport.ExportImageToFile(pn, pat_img, 1920, 1080)
                        print(f"  {pn} HFSS图片: {pat_img}")
                    except Exception as ei:
                        print(f"  {pn} 图片导出失败: {ei}")
                    break
                except Exception as e:
                    print(f"  {pn} ({sweep_ctx}): {e}")

        # S11 绘图 (matplotlib)
        if len(freqs) > 0:
            try:
                import matplotlib
                matplotlib.use('Agg')
                import matplotlib.pyplot as plt

                fig, ax = plt.subplots(figsize=(10, 6))
                ax.plot(freqs, s11_vals, 'b-', lw=2, label='S11')
                ax.axhline(-10, color='r', ls='--', lw=1, label='-10 dB')
                ax.axvline(TARGET_F, color='g', ls=':', lw=1.5,
                           label=f'Target {TARGET_F} GHz')
                ax.axvline(f_res, color='orange', ls='-.', lw=1.5,
                           label=f'Resonance {f_res:.3f} GHz')
                if bw > 0:
                    ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green',
                               label=f'BW {bw * 1e3:.0f} MHz ({bw_pct:.1f}%)')
                ax.set_xlabel('Frequency (GHz)')
                ax.set_ylabel('S11 (dB)')
                ax.set_title(f'Printed Dipole with Microstrip Balun @ {TARGET_F} GHz\n'
                             f'f_res={f_res:.4f} GHz, S11_min={s11_min:.2f} dB')
                ax.legend()
                ax.grid(True, alpha=0.3)
                ax.set_xlim(1.5, 3.0)
                png_path = os.path.join(PROJECT_DIR, "s11_balun_2217.png")
                fig.savefig(png_path, dpi=150, bbox_inches='tight')
                plt.close(fig)
                print(f"  S11 图表: {png_path}")
            except Exception as e:
                print(f"  绘图失败: {e}")

        # 辐射方向图绘制 (从 Rectangular Plot CSV 转换为极坐标图)
        # AEDT 2019.1 Far Fields CSV 的 GainTotal 为线性值，需转 dBi
        try:
            import matplotlib
            matplotlib.use('Agg')
            import matplotlib.pyplot as plt

            e_csv_path = os.path.join(PROJECT_DIR, "e_plane_2217.csv")
            h_csv_path = os.path.join(PROJECT_DIR, "h_plane_2217.csv")
            if os.path.exists(e_csv_path) and os.path.exists(h_csv_path):
                def read_pattern(p):
                    """读取 Far Fields Rectangular Plot CSV，返回 theta(deg) 和 GainTotal(线性)"""
                    t, g = [], []
                    with open(p) as f:
                        rd = csv.reader(f)
                        next(rd)  # skip header
                        for row in rd:
                            if len(row) >= 2:
                                try:
                                    t.append(float(row[0]))
                                    g.append(float(row[1]))
                                except ValueError:
                                    pass
                    return np.array(t), np.array(g)

                te, ge_lin = read_pattern(e_csv_path)
                th, gh_lin = read_pattern(h_csv_path)
                if len(ge_lin) > 0 and len(gh_lin) > 0:
                    # 线性转 dBi：10*log10(val)，避免 log(0)
                    ge_dbi = 10.0 * np.log10(np.maximum(ge_lin, 1e-12))
                    gh_dbi = 10.0 * np.log10(np.maximum(gh_lin, 1e-12))
                    me = np.max(ge_dbi)
                    mh = np.max(gh_dbi)
                    print(f"\n  E-plane 最大增益: {me:.2f} dBi")
                    print(f"  H-plane 最大增益: {mh:.2f} dBi")

                    # --- 极坐标辐射方向图 ---
                    fig, (a1, a2) = plt.subplots(
                        1, 2, subplot_kw={'projection': 'polar'}, figsize=(14, 6))

                    # 为极坐标归一化：将 dBi 映射到 [0, 1] 以便可视化
                    r_floor = -30  # dBi 下限
                    ge_plot = np.clip(ge_dbi - r_floor, 0, None)
                    gh_plot = np.clip(gh_dbi - r_floor, 0, None)

                    a1.plot(np.radians(te), ge_plot, 'b-', lw=2)
                    a1.fill(np.radians(te), ge_plot, alpha=0.15, color='b')
                    a1.set_title(f'E-Plane (phi=0°)\nMax Gain={me:.2f} dBi', pad=20)
                    a1.set_theta_zero_location('N')
                    a1.set_theta_direction(-1)

                    a2.plot(np.radians(th), gh_plot, 'r-', lw=2)
                    a2.fill(np.radians(th), gh_plot, alpha=0.15, color='r')
                    a2.set_title(f'H-Plane (phi=90°)\nMax Gain={mh:.2f} dBi', pad=20)
                    a2.set_theta_zero_location('N')
                    a2.set_theta_direction(-1)

                    fig.suptitle(
                        f'Printed Dipole with Balun @ {TARGET_F} GHz\nRadiation Pattern',
                        fontsize=14)
                    rad_png = os.path.join(PROJECT_DIR, "radiation_2217.png")
                    fig.savefig(rad_png, dpi=150, bbox_inches='tight')
                    plt.close(fig)
                    print(f"  辐射方向图 (极坐标): {rad_png}")

                    # --- 增益图 (笛卡尔坐标 dBi vs Theta) ---
                    fig2, ax2 = plt.subplots(figsize=(10, 6))
                    ax2.plot(te, ge_dbi, 'b-', lw=2, label=f'E-Plane (phi=0°)')
                    ax2.plot(th, gh_dbi, 'r--', lw=2, label=f'H-Plane (phi=90°)')
                    ax2.set_xlabel('Theta (deg)')
                    ax2.set_ylabel('Gain (dBi)')
                    ax2.set_title(f'Antenna Gain @ {TARGET_F} GHz\n'
                                  f'E-plane max={me:.2f} dBi, H-plane max={mh:.2f} dBi')
                    ax2.legend()
                    ax2.grid(True, alpha=0.3)
                    ax2.set_xlim(0, 180)
                    gain_png = os.path.join(PROJECT_DIR, "gain_2217.png")
                    fig2.savefig(gain_png, dpi=150, bbox_inches='tight')
                    plt.close(fig2)
                    print(f"  增益图 (笛卡尔): {gain_png}")
        except Exception as e:
            print(f"  辐射方向图绘制失败: {e}")

        # ==============================================================
        # 6g. 导出 3D 模型几何截图
        # ==============================================================
        print("\n[Step 6g] 导出 3D 模型截图...")
        model_img = os.path.join(PROJECT_DIR, "Model_3D_HFSS.jpg")
        try:
            oEditor.FitAll()
            oEditor.ExportModelImageToFile(model_img, 1920, 1080,
                ["NAME:SaveImageParams",
                 "ShowAxis:=", "True",
                 "ShowGrid:=", "True",
                 "ShowRuler:=", "True"])
            print(f"  3D模型截图: {model_img}")
        except Exception:
            # 备用方法: 使用 oDesign.ExportImage
            try:
                oEditor.FitAll()
                oDesign.ExportImage(model_img, 1920, 1080)
                print(f"  3D模型截图(备用): {model_img}")
            except Exception as e2:
                print(f"  3D模型截图导出失败: {e2}")

        # ==============================================================
        # 7. 参数扫描优化 (Optimetrics)
        # ==============================================================
        need_optim = (len(freqs) == 0
                      or s11_at_target > S11_TARGET
                      or freq_err > 1.0)

        if need_optim:
            print(f"\n[Step 7] 参数扫描优化 (Optimetrics)...")
            if len(freqs) > 0:
                print(f"  S11 @ {TARGET_F} GHz = {s11_at_target:.2f} dB"
                      f" (目标 < {S11_TARGET} dB)")
                print(f"  频率误差 = {freq_err:.2f}% (目标 < 1%)")
            print(f"  对 L2 和 L3 进行参数扫描优化")

            try:
                oOptimetrics = oDesign.GetModule("Optimetrics")
                oOptimetrics.InsertSetup("OptiParametric",
                    ["NAME:ParamSweep_L2_L3",
                     "IsEnabled:=", True,
                     ["NAME:ProdOptiSetupDataV2",
                      "SaveFields:=", False,
                      "CopyMesh:=", False,
                      "SolveWithCopiedMeshOnly:=", 0],
                     ["NAME:StartingPoint"],
                     "Sim. Setups:=", ["Setup1"],
                     ["NAME:Sweeps",
                      ["NAME:SweepDefinition",
                       "Variable:=", "L2",
                       "Data:=", "LIN 20mm 28mm 1mm",
                       "OffsetF1:=", False,
                       "Synchronize:=", 0],
                      ["NAME:SweepDefinition",
                       "Variable:=", "L3",
                       "Data:=", "LIN 8mm 14mm 1mm",
                       "OffsetF1:=", False,
                       "Synchronize:=", 0]],
                     ["NAME:Sweep Operations"],
                     ["NAME:Goals"]])
                print("  参数扫描已设置:")
                print("    L2: 20mm ~ 28mm, step=1mm")
                print("    L3: 8mm ~ 14mm,  step=1mm")
                print(f"    共 {9 * 7} = 63 个参数组合")

                print("  运行参数扫描 (预计耗时较长)...")
                t0 = time.time()
                oOptimetrics.SolveSetup("ParamSweep_L2_L3")
                ts = time.time() - t0
                print(f"  参数扫描完成! 用时 {ts:.0f}s ({ts / 60:.1f} min)")

            except Exception as e:
                print(f"  参数扫描失败: {e}")
                traceback.print_exc()
        else:
            print(f"\n[Step 7] 跳过参数扫描 (已达到设计目标)")
            print(f"  S11 @ {TARGET_F} GHz = {s11_at_target:.2f} dB < {S11_TARGET} dB")
            print(f"  频率误差 = {freq_err:.2f}% < 1%")

        # ==============================================================
        # 保存 & 汇总
        # ==============================================================
        hfss.save_project()
        proj_path = os.path.join(PROJECT_DIR, PROJ_NAME + ".aedt")

        print("\n" + "=" * 70)
        print("设计完成! 结果汇总")
        print("=" * 70)
        print(f"  天线类型:     印刷偶极子天线（带微带巴伦馈线）")
        print(f"  目标频率:     {TARGET_F} GHz")
        print(f"  求解类型:     Driven Modal")
        print(f"  介质板:       FR4_epoxy (εr=4.4, H=1.6mm)")
        print(f"  馈电方式:     Lumped Port 50Ω")
        if len(freqs) > 0:
            print(f"  谐振频率:     {f_res:.4f} GHz")
            print(f"  S11 最小值:   {s11_min:.2f} dB")
            print(f"  S11 @ 目标:   {s11_at_target:.2f} dB")
            if bw > 0:
                print(f"  -10dB 带宽:   {bw * 1e3:.0f} MHz ({bw_pct:.1f}%)")
            print(f"  频率误差:     {freq_err:.2f}%")
        print(f"\n  项目文件:     {proj_path}")
        print(f"  S11 数据:     {s11_csv}")
        print(f"  Touchstone:   {s1p_file}")
        print(f"  日志文件:     {LOG}")

        # 列出已导出的 HFSS 图片
        hfss_images = ["S11_Plot_HFSS.jpg", "Gain_3D_HFSS.jpg",
                       "E_Plane_HFSS.jpg", "H_Plane_HFSS.jpg",
                       "Model_3D_HFSS.jpg"]
        print(f"\n  HFSS 结果图片:")
        for img_name in hfss_images:
            img_path = os.path.join(PROJECT_DIR, img_name)
            if os.path.exists(img_path):
                print(f"    ✓ {img_path}")

        # 写入设计摘要文件
        summary_path = os.path.join(PROJECT_DIR, "design_summary_2217.txt")
        with open(summary_path, "w", encoding="utf-8") as sf:
            sf.write("印刷偶极子天线（带微带巴伦馈线）设计摘要\n")
            sf.write("=" * 50 + "\n")
            sf.write(f"目标频率: {TARGET_F} GHz\n")
            sf.write(f"求解类型: Driven Modal\n")
            sf.write(f"介质板:   FR4_epoxy (εr=4.4)\n\n")
            sf.write("设计变量:\n")
            for name, value in variables.items():
                sf.write(f"  {name} = {value}\n")
            sf.write(f"\n仿真结果:\n")
            if len(freqs) > 0:
                sf.write(f"  谐振频率:     {f_res:.4f} GHz\n")
                sf.write(f"  S11 最小值:   {s11_min:.2f} dB\n")
                sf.write(f"  S11 @ 目标:   {s11_at_target:.2f} dB\n")
                if bw > 0:
                    sf.write(f"  -10dB 带宽:   {bw_lo:.3f} ~ {bw_hi:.3f} GHz\n")
                sf.write(f"  频率误差:     {freq_err:.2f}%\n")
            sf.write(f"\n输出文件:\n")
            sf.write(f"  项目: {proj_path}\n")
            sf.write(f"  S11:  {s11_csv}\n")
            sf.write(f"  s1p:  {s1p_file}\n")
            sf.write(f"\nHFSS 结果图片:\n")
            for img_name in ["S11_Plot_HFSS.jpg", "Gain_3D_HFSS.jpg",
                             "E_Plane_HFSS.jpg", "H_Plane_HFSS.jpg",
                             "Model_3D_HFSS.jpg"]:
                img_path = os.path.join(PROJECT_DIR, img_name)
                if os.path.exists(img_path):
                    sf.write(f"  {img_path}\n")
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
