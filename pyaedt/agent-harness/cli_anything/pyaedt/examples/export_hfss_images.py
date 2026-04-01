#!/usr/bin/env python3
"""
export_hfss_images.py
打开已有 HFSS 工程，以图形模式导出所有结果图片（S11、方向图、3D模型截图）。
需要 AEDT GUI 渲染，因此 non_graphical=False。
"""

import os, sys, time, subprocess, traceback

# ============================================================================
# 配置
# ============================================================================
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
PROJ_NAME = "Printed_Dipole_Balun_2217"
PROJ_FILE = os.path.join(PROJECT_DIR, PROJ_NAME + ".aedt")
DESIGN_NAME = "Dipole_Balun"
TARGET_F = 2.217  # GHz

# 输出图片分辨率
IMG_W, IMG_H = 1920, 1080


def kill_aedt():
    try:
        subprocess.run(["taskkill", "/F", "/IM", "ansysedt.exe"],
                       capture_output=True, timeout=15)
        time.sleep(3)
    except Exception:
        try:
            import psutil
            for p in psutil.process_iter(['name']):
                if p.info['name'] and 'ansysedt' in p.info['name'].lower():
                    p.kill()
            time.sleep(3)
        except Exception:
            pass


def main():
    print("=" * 60)
    print("HFSS 工程结果图片导出")
    print("=" * 60)

    if not os.path.exists(PROJ_FILE):
        print(f"错误: 工程文件不存在: {PROJ_FILE}")
        return

    kill_aedt()
    time.sleep(2)

    # ------------------------------------------------------------------
    # PyAEDT 兼容性补丁 (AEDT 2019.1)
    # ------------------------------------------------------------------
    try:
        from pyaedt import desktop as _dm
        _orig = _dm.Desktop.__init__

        def _patched(self, *a, **k):
            _orig(self, *a, **k)
            if not hasattr(self, 'student_version'):
                self.student_version = False

        _dm.Desktop.__init__ = _patched
    except Exception:
        pass

    try:
        import pyaedt.application.Design as _dd
        _orig2 = _dd.DesignSettings.__init__

        def _patched2(self, *a, **k):
            try:
                _orig2(self, *a, **k)
            except AttributeError:
                pass

        _dd.DesignSettings.__init__ = _patched2
    except Exception:
        pass

    # ------------------------------------------------------------------
    # 以图形模式打开已有工程
    # ------------------------------------------------------------------
    from pyaedt import Hfss

    print(f"\n打开工程 (图形模式): {PROJ_FILE}")
    hfss = Hfss(
        projectname=PROJ_FILE,
        designname=DESIGN_NAME,
        solution_type="DrivenModal",
        non_graphical=False,       # 图形模式 - 必须用于图片导出
        new_desktop_session=True,
        specified_version="2019.1",
    )

    oDesign = hfss.odesign
    oEditor = oDesign.SetActiveEditor("3D Modeler")
    oReport = oDesign.GetModule("ReportSetup")
    oDesktop = hfss.odesktop

    print("  工程已打开!")
    time.sleep(3)  # 等待 GUI 完全渲染

    exported = []

    # ------------------------------------------------------------------
    # 1. 导出 S11 报告图片
    # ------------------------------------------------------------------
    print("\n[1] 导出 S11 报告图片...")

    # 确保 S11_Plot 报告存在，若无则创建
    s_expr = "dB(S(Port1,Port1))"
    try:
        report_names = oReport.GetAllReportNames()
        print(f"  现有报告: {list(report_names)}")
    except Exception:
        report_names = []

    if "S11_Plot" not in report_names:
        print("  创建 S11_Plot 报告...")
        try:
            oReport.CreateReport(
                "S11_Plot", "Modal Solution Data", "Rectangular Plot",
                "Setup1 : Sweep1", [],
                ["Freq:=", ["All"]],
                ["X Component:=", "Freq", "Y Component:=", [s_expr]], [])
        except Exception as e:
            print(f"  创建失败: {e}")

    s11_img = os.path.join(PROJECT_DIR, "S11_Plot_HFSS.jpg")
    for method_name, export_fn in [
        ("ExportImageToFile", lambda: oReport.ExportImageToFile(
            "S11_Plot", s11_img, IMG_W, IMG_H)),
        ("ExportToFile(.bmp)", lambda: oReport.ExportToFile(
            "S11_Plot", s11_img.replace(".jpg", ".bmp"))),
    ]:
        try:
            export_fn()
            out = s11_img if "Image" in method_name else s11_img.replace(".jpg", ".bmp")
            if os.path.exists(out):
                print(f"  S11 图片已导出 ({method_name}): {out}")
                exported.append(out)
                break
        except Exception as e:
            print(f"  {method_name} 失败: {e}")

    # ------------------------------------------------------------------
    # 2. 导出 3D 增益方向图 (Rectangular Plot — AEDT 2019.1 兼容)
    # ------------------------------------------------------------------
    print("\n[2] 导出增益报告 (Rectangular Plot)...")
    # 先删除旧的 broken 报告
    for old_rpt in ["Gain_3D"]:
        if old_rpt in report_names:
            try:
                oReport.DeleteReports([old_rpt])
            except Exception:
                pass

    for ctx in ["Setup1 : Sweep2", "Setup1 : LastAdaptive", "Setup1 : Sweep1"]:
        try:
            oReport.CreateReport(
                "Gain_3D", "Far Fields", "Rectangular Plot", ctx,
                ["Context:=", "InfSphere1"],
                ["Theta:=", ["All"], "Phi:=", ["All"],
                 "Freq:=", [f"{TARGET_F}GHz"]],
                ["X Component:=", "Theta",
                 "Y Component:=", ["GainTotal"]], [])
            gain3d_csv = os.path.join(PROJECT_DIR, "gain_3d_2217.csv")
            oReport.ExportToFile("Gain_3D", gain3d_csv)
            print(f"  Gain_3D 报告已创建 + CSV导出 ({ctx})")
            break
        except Exception as e:
            print(f"  Gain_3D 创建失败 ({ctx}): {e}")

    gain_img = os.path.join(PROJECT_DIR, "Gain_3D_HFSS.jpg")
    try:
        oReport.ExportImageToFile("Gain_3D", gain_img, IMG_W, IMG_H)
        if os.path.exists(gain_img):
            print(f"  3D增益图片: {gain_img}")
            exported.append(gain_img)
    except Exception as e:
        print(f"  3D增益导出失败: {e}")

    # ------------------------------------------------------------------
    # 3. 导出 E/H 面方向图 (Rectangular Plot — AEDT 2019.1 兼容)
    # ------------------------------------------------------------------
    print("\n[3] 导出 E/H 面辐射方向图 (Rectangular Plot)...")
    csv_files = {}
    for pn, phi_val in [("E_Plane", "0deg"), ("H_Plane", "90deg")]:
        # 先删除旧报告
        if pn in report_names:
            try:
                oReport.DeleteReports([pn])
            except Exception:
                pass

        for ctx in ["Setup1 : Sweep2", "Setup1 : LastAdaptive", "Setup1 : Sweep1"]:
            try:
                oReport.CreateReport(
                    pn, "Far Fields", "Rectangular Plot", ctx,
                    ["Context:=", "InfSphere1"],
                    ["Theta:=", ["All"], "Phi:=", [phi_val],
                     "Freq:=", [f"{TARGET_F}GHz"]],
                    ["X Component:=", "Theta",
                     "Y Component:=", ["GainTotal"]], [])
                csv_path = os.path.join(PROJECT_DIR, f"{pn.lower()}_2217.csv")
                oReport.ExportToFile(pn, csv_path)
                csv_files[pn] = csv_path
                print(f"  {pn} 报告已创建 + CSV导出 ({ctx})")
                break
            except Exception as e:
                print(f"  {pn} 创建失败 ({ctx}): {e}")

        pat_img = os.path.join(PROJECT_DIR, f"{pn}_HFSS.jpg")
        try:
            oReport.ExportImageToFile(pn, pat_img, IMG_W, IMG_H)
            if os.path.exists(pat_img):
                print(f"  {pn} 图片: {pat_img}")
                exported.append(pat_img)
        except Exception as e:
            print(f"  {pn} 导出失败: {e}")

    # ------------------------------------------------------------------
    # 3b. 用 matplotlib 将 CSV 数据绘制为极坐标方向图和增益图
    # ------------------------------------------------------------------
    print("\n[3b] 生成 matplotlib 辐射方向图 & 增益图...")
    try:
        import csv as csv_mod
        import numpy as np
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt

        e_csv = csv_files.get("E_Plane", os.path.join(PROJECT_DIR, "e_plane_2217.csv"))
        h_csv = csv_files.get("H_Plane", os.path.join(PROJECT_DIR, "h_plane_2217.csv"))

        if os.path.exists(e_csv) and os.path.exists(h_csv):
            def read_ff_csv(path):
                theta, gain = [], []
                with open(path) as f:
                    rd = csv_mod.reader(f)
                    next(rd)  # skip header
                    for row in rd:
                        if len(row) >= 2:
                            try:
                                theta.append(float(row[0]))
                                gain.append(float(row[1]))  # linear
                            except ValueError:
                                pass
                return np.array(theta), np.array(gain)

            te, ge_lin = read_ff_csv(e_csv)
            th, gh_lin = read_ff_csv(h_csv)

            if len(ge_lin) > 0 and len(gh_lin) > 0:
                ge_dbi = 10.0 * np.log10(np.maximum(ge_lin, 1e-12))
                gh_dbi = 10.0 * np.log10(np.maximum(gh_lin, 1e-12))
                me = np.max(ge_dbi)
                mh = np.max(gh_dbi)
                print(f"  E-plane 最大增益: {me:.2f} dBi")
                print(f"  H-plane 最大增益: {mh:.2f} dBi")

                # 极坐标方向图
                r_floor = -30
                ge_r = np.clip(ge_dbi - r_floor, 0, None)
                gh_r = np.clip(gh_dbi - r_floor, 0, None)
                fig, (a1, a2) = plt.subplots(
                    1, 2, subplot_kw={'projection': 'polar'}, figsize=(14, 6))
                a1.plot(np.radians(te), ge_r, 'b-', lw=2)
                a1.fill(np.radians(te), ge_r, alpha=0.15, color='b')
                a1.set_title(f'E-Plane (phi=0°)\nMax Gain={me:.2f} dBi', pad=20)
                a1.set_theta_zero_location('N')
                a1.set_theta_direction(-1)
                a2.plot(np.radians(th), gh_r, 'r-', lw=2)
                a2.fill(np.radians(th), gh_r, alpha=0.15, color='r')
                a2.set_title(f'H-Plane (phi=90°)\nMax Gain={mh:.2f} dBi', pad=20)
                a2.set_theta_zero_location('N')
                a2.set_theta_direction(-1)
                fig.suptitle(f'Radiation Pattern @ {TARGET_F} GHz', fontsize=14)
                rad_png = os.path.join(PROJECT_DIR, "radiation_2217.png")
                fig.savefig(rad_png, dpi=150, bbox_inches='tight')
                plt.close(fig)
                print(f"  辐射方向图: {rad_png}")
                exported.append(rad_png)

                # 增益图 (笛卡尔坐标)
                fig2, ax2 = plt.subplots(figsize=(10, 6))
                ax2.plot(te, ge_dbi, 'b-', lw=2, label='E-Plane (phi=0°)')
                ax2.plot(th, gh_dbi, 'r--', lw=2, label='H-Plane (phi=90°)')
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
                print(f"  增益图: {gain_png}")
                exported.append(gain_png)
    except Exception as e:
        print(f"  matplotlib 绘图失败: {e}")

    # ------------------------------------------------------------------
    # 4. 导出 3D 模型截图
    # ------------------------------------------------------------------
    print("\n[4] 导出 3D 模型截图...")
    model_img = os.path.join(PROJECT_DIR, "Model_3D_HFSS.jpg")
    for method_name, export_fn in [
        ("oEditor.ExportModelImageToFile", lambda: (
            oEditor.FitAll(),
            time.sleep(1),
            oEditor.ExportModelImageToFile(model_img, IMG_W, IMG_H,
                ["NAME:SaveImageParams",
                 "ShowAxis:=", "True",
                 "ShowGrid:=", "True",
                 "ShowRuler:=", "True"])
        )),
        ("oDesign.ExportImage", lambda: (
            oEditor.FitAll(),
            time.sleep(1),
            oDesign.ExportImage(model_img, IMG_W, IMG_H)
        )),
        ("oDesktop.ExportImage (bmp)", lambda: (
            oEditor.FitAll(),
            time.sleep(1),
            oDesktop.ExportImage(model_img.replace(".jpg", ".bmp"), IMG_W, IMG_H)
        )),
    ]:
        try:
            export_fn()
            out = model_img if ".bmp" not in method_name else model_img.replace(".jpg", ".bmp")
            if os.path.exists(out):
                print(f"  3D模型截图 ({method_name}): {out}")
                exported.append(out)
                break
            else:
                print(f"  {method_name}: 调用成功但未生成文件")
        except Exception as e:
            print(f"  {method_name} 失败: {e}")

    # ------------------------------------------------------------------
    # 5. 汇总输出
    # ------------------------------------------------------------------
    print("\n" + "=" * 60)
    print("导出完成!")
    print("=" * 60)
    if exported:
        print(f"成功导出 {len(exported)} 个图片:")
        for p in exported:
            sz = os.path.getsize(p) / 1024
            print(f"  ✓ {p}  ({sz:.0f} KB)")
    else:
        print("  未能导出任何 HFSS 图片。")
        print("  提示: matplotlib 生成的 S11 图片在:")
        s11_png = os.path.join(PROJECT_DIR, "s11_balun_2217.png")
        if os.path.exists(s11_png):
            print(f"  ✓ {s11_png}")

    # 列出所有已有的结果文件
    print(f"\n所有结果文件 ({PROJECT_DIR}):")
    for f in sorted(os.listdir(PROJECT_DIR)):
        fp = os.path.join(PROJECT_DIR, f)
        if os.path.isfile(fp):
            ext = os.path.splitext(f)[1].lower()
            if ext in ('.png', '.jpg', '.bmp', '.csv', '.s1p', '.txt'):
                sz = os.path.getsize(fp) / 1024
                print(f"  {f}  ({sz:.1f} KB)")

    hfss.release_desktop(close_projects=True)
    print("\nAEDT 已关闭.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n致命错误: {e}")
        traceback.print_exc()
