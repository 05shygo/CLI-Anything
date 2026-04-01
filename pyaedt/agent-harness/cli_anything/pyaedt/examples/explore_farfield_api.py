#!/usr/bin/env python3
"""
explore_farfield_api.py
探索 AEDT 2019.1 中 Far Fields 报告的正确 API 格式
"""
import os, sys, time, subprocess, traceback

os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
PROJ_FILE = os.path.join(PROJECT_DIR, "Printed_Dipole_Balun_2217.aedt")
DESIGN_NAME = "Dipole_Balun"
TARGET_F = 2.217

def main():
    # 补丁
    try:
        from pyaedt import desktop as _dm
        _orig = _dm.Desktop.__init__
        def _p(self, *a, **k):
            _orig(self, *a, **k)
            if not hasattr(self, 'student_version'): self.student_version = False
        _dm.Desktop.__init__ = _p
    except: pass
    try:
        import pyaedt.application.Design as _dd
        _o2 = _dd.DesignSettings.__init__
        def _p2(self, *a, **k):
            try: _o2(self, *a, **k)
            except AttributeError: pass
        _dd.DesignSettings.__init__ = _p2
    except: pass

    from pyaedt import Hfss
    print("Opening project...")
    hfss = Hfss(projectname=PROJ_FILE, designname=DESIGN_NAME,
                solution_type="DrivenModal", non_graphical=False,
                new_desktop_session=True, specified_version="2019.1")

    oDesign = hfss.odesign
    oReport = oDesign.GetModule("ReportSetup")
    oRadField = oDesign.GetModule("RadField")
    oSolutions = oDesign.GetModule("Solutions")

    time.sleep(2)
    print("Project opened.\n")

    # 1. 列出已有报告
    try:
        reports = oReport.GetAllReportNames()
        print(f"Existing reports: {list(reports)}")
    except Exception as e:
        print(f"GetAllReportNames failed: {e}")

    # 2. 获取可用的报告类别
    print("\n=== GetAvailableReportTypes ===")
    try:
        types = oReport.GetAvailableReportTypes()
        print(f"Report types: {list(types)}")
    except Exception as e:
        print(f"GetAvailableReportTypes failed: {e}")

    # 3. 获取各类别的显示类型
    print("\n=== GetAvailableDisplayTypes ===")
    for rt in ["Modal Solution Data", "Far Fields", "Near Fields",
               "Terminal Solution Data", "Fields"]:
        try:
            dt = oReport.GetAvailableDisplayTypes(rt)
            print(f"  {rt}: {list(dt)}")
        except Exception as e:
            print(f"  {rt}: FAILED - {e}")

    # 4. 获取可用的解决方案上下文
    print("\n=== GetAvailableSolutions ===")
    for rt in ["Modal Solution Data", "Far Fields"]:
        try:
            sol = oReport.GetAvailableSolutions(rt)
            print(f"  {rt}: {list(sol)}")
        except Exception as e:
            print(f"  {rt}: FAILED - {e}")

    # 5. 获取远场可用的量
    print("\n=== GetAllCategories / GetAllQuantities ===")
    for rt in ["Far Fields", "Modal Solution Data"]:
        try:
            cats = oReport.GetAllCategories(rt)
            print(f"  {rt} categories: {list(cats)}")
        except Exception as e:
            print(f"  {rt} categories: FAILED - {e}")
        try:
            quantities = oReport.GetAllQuantities(rt)
            print(f"  {rt} quantities: {list(quantities)[:20]}")
        except Exception as e:
            print(f"  {rt} quantities: FAILED - {e}")

    # 6. 尝试 Far Fields 报告的不同格式
    print("\n=== Trying CreateReport variations ===")
    ff_name = "InfSphere1"

    attempts = [
        # (desc, report_type, display_type, context, sweep_ctx, families, traces)
        ("A: Far Fields + Radiation Pattern + Sweep2",
         "Far Fields", "Radiation Pattern", "Setup1 : Sweep2",
         ["Context:=", ff_name],
         ["Theta:=", ["All"], "Phi:=", ["0deg"], "Freq:=", [f"{TARGET_F}GHz"]],
         ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]]),

        ("B: Far Fields + Rectangular Plot + Sweep2",
         "Far Fields", "Rectangular Plot", "Setup1 : Sweep2",
         ["Context:=", ff_name],
         ["Theta:=", ["All"], "Phi:=", ["0deg"], "Freq:=", [f"{TARGET_F}GHz"]],
         ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]]),

        ("C: Far Fields + Radiation Pattern + LastAdaptive (no context)",
         "Far Fields", "Radiation Pattern", "Setup1 : LastAdaptive",
         [],
         ["Theta:=", ["All"], "Phi:=", ["0deg"], "Freq:=", [f"{TARGET_F}GHz"]],
         ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]]),

        ("D: Far Fields + Radiation Pattern + LastAdaptive + rETotal",
         "Far Fields", "Radiation Pattern", "Setup1 : LastAdaptive",
         ["Context:=", ff_name],
         ["Theta:=", ["All"], "Phi:=", ["0deg"], "Freq:=", [f"{TARGET_F}GHz"]],
         ["X Component:=", "Theta", "Y Component:=", ["rETotal"]]),

        ("E: Far Fields + Data Table + Sweep2",
         "Far Fields", "Data Table", "Setup1 : Sweep2",
         ["Context:=", ff_name],
         ["Theta:=", ["All"], "Phi:=", ["0deg"], "Freq:=", [f"{TARGET_F}GHz"]],
         ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]]),

        ("F: Modal Solution Data + Radiation Pattern",
         "Modal Solution Data", "Radiation Pattern", "Setup1 : LastAdaptive",
         ["Context:=", ff_name],
         ["Theta:=", ["All"], "Phi:=", ["0deg"], "Freq:=", [f"{TARGET_F}GHz"]],
         ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]]),

        ("G: Far Fields + Radiation Pattern + Theta -180 to 180",
         "Far Fields", "Radiation Pattern", "Setup1 : Sweep2",
         ["Context:=", ff_name],
         ["Theta:=", ["All"], "Phi:=", ["All"], "Freq:=", [f"{TARGET_F}GHz"]],
         ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]]),

        ("H: Far Fields + 3D Polar Plot + Sweep2",
         "Far Fields", "3D Polar Plot", "Setup1 : Sweep2",
         ["Context:=", ff_name],
         ["Theta:=", ["All"], "Phi:=", ["All"], "Freq:=", [f"{TARGET_F}GHz"]],
         ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]]),
    ]

    for i, (desc, rtype, dtype, sctx, ctx, fam, trace) in enumerate(attempts):
        rn = f"Test_{chr(65+i)}"
        try:
            oReport.CreateReport(rn, rtype, dtype, sctx, ctx, fam, trace, [])
            print(f"  {desc}: SUCCESS!")
            # 尝试导出
            csv_out = os.path.join(PROJECT_DIR, f"test_{chr(97+i)}.csv")
            try:
                oReport.ExportToFile(rn, csv_out)
                sz = os.path.getsize(csv_out) if os.path.exists(csv_out) else 0
                print(f"    Exported {csv_out} ({sz} bytes)")
            except Exception as ex:
                print(f"    Export failed: {ex}")
            # 尝试导出图片
            img_out = os.path.join(PROJECT_DIR, f"test_{chr(97+i)}.jpg")
            try:
                oReport.ExportImageToFile(rn, img_out, 1920, 1080)
                sz = os.path.getsize(img_out) if os.path.exists(img_out) else 0
                print(f"    Image {img_out} ({sz} bytes)")
            except Exception as ex:
                print(f"    Image export failed: {ex}")
        except Exception as e:
            err_str = str(e)
            # 只打印关键错误码
            if "2147024382" in err_str:
                print(f"  {desc}: FILE_NOT_FOUND")
            elif "2147352567" in err_str:
                print(f"  {desc}: COM_EXCEPTION - {err_str[:80]}")
            else:
                print(f"  {desc}: FAILED - {err_str[:100]}")

    # 7. 尝试 RadField 模块直接导出
    print("\n=== RadField module export methods ===")

    # ExportRadiationParametersToFile
    try:
        out_txt = os.path.join(PROJECT_DIR, "rad_params.txt")
        oRadField.ExportRadiationParametersToFile(
            "Setup1 : Sweep2", ff_name, out_txt)
        sz = os.path.getsize(out_txt) if os.path.exists(out_txt) else 0
        print(f"  ExportRadiationParametersToFile: {sz} bytes")
    except Exception as e:
        print(f"  ExportRadiationParametersToFile: {e}")

    # ExportRadFieldsToFile (try .csv)
    try:
        out_csv = os.path.join(PROJECT_DIR, "rad_fields.csv")
        oRadField.ExportRadFieldsToFile("Setup1 : Sweep2", ff_name, out_csv)
        print(f"  ExportRadFieldsToFile: OK")
    except Exception as e:
        print(f"  ExportRadFieldsToFile: {e}")

    # 8. 尝试 Solutions 模块
    print("\n=== Solutions module ===")
    try:
        valid_setups = oSolutions.GetValidISSolutions()
        print(f"  Valid solutions: {list(valid_setups)}")
    except Exception as e:
        print(f"  GetValidISSolutions: {e}")

    try:
        data = oSolutions.GetSolutionDataPerVariation(
            "Far Fields", "Setup1 : Sweep2",
            ["Context:=", ff_name],
            ["Theta:=", ["All"], "Phi:=", ["0deg"]],
            ["GainTotal"])
        print(f"  GetSolutionDataPerVariation: {data}")
    except Exception as e:
        print(f"  GetSolutionDataPerVariation: {e}")

    # 9. 用 VBScript/IronPython 方式尝试
    print("\n=== VBScript approach ===")
    try:
        vbs = f'''
Dim oModule
Set oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport "VBS_Test", "Far Fields", "Radiation Pattern", _
    "Setup1 : Sweep2", _
    Array("Context:=", "{ff_name}"), _
    Array("Theta:=", Array("All"), "Phi:=", Array("0deg"), "Freq:=", Array("{TARGET_F}GHz")), _
    Array("X Component:=", "Theta", "Y Component:=", Array("GainTotal")), _
    Array()
'''
        # Can't directly run VBS through COM, but let's try ExecuteScriptCode
        try:
            oDesign.GetDesktopObject().RunScript(vbs)
        except: pass
    except Exception as e:
        print(f"  VBS: {e}")

    print("\n=== Done ===")
    hfss.release_desktop(close_projects=True)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Fatal: {e}")
        traceback.print_exc()
