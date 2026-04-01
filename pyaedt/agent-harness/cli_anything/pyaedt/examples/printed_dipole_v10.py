#!/usr/bin/env python3
"""
printed_dipole_v10.py - 印刷偶极子天线 v10 自动优化
  基于 v9 模板 (Validation=1, 仿真成功)
  内建优化循环: 每次迭代新建完整 AEDT 项目
  起始 arm=108mm (v9: 90mm→2.66GHz 缩放外推)
  目标: f0=2.217GHz, S11<-10dB
"""
import numpy as np, os, sys, time, csv, traceback, subprocess, gc

os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.makedirs(PROJECT_DIR, exist_ok=True)
LOG = os.path.join(PROJECT_DIR, "dipole_v10.log")

class Logger:
    def __init__(self, p):
        self.t = sys.stdout; self.f = open(p, "w", encoding="utf-8")
    def write(self, m):
        try: self.t.write(m)
        except UnicodeEncodeError: self.t.write(m.encode('ascii','replace').decode())
        self.f.write(m); self.f.flush()
    def flush(self): self.t.flush(); self.f.flush()
sys.stdout = Logger(LOG); sys.stderr = sys.stdout

# ============================================================================
# 优化配置
# ============================================================================
TARGET_F = 2.217       # GHz
FREQ_TOL = 0.02        # GHz (允许误差 ±20MHz)
S11_TARGET = -10.0     # dB
MAX_ITER = 5
ARM_START = 108.0      # mm (从 v9 缩放: 90*(2.66/2.217)≈108)

# 物理参数
c = 3e8; f0 = TARGET_F * 1e9; eps_r = 4.4
lambda_0 = c / f0; lambda_g = lambda_0 / np.sqrt(eps_r)

sub_L = 120.0; sub_W = 60.0; sub_H = 1.6; cu_t = 0.035
dipole_w = 3.0; gap = 1.0
balun_L = lambda_g * 1e3 / 4.0; balun_segs = 10
balun_w_start = 1.0; balun_w_end = 3.0
pad = lambda_0 * 1e3 / 4.0
f_min = 1.5e9; f_max = 3.0e9; n_pts = 151

print("=" * 70)
print("PRINTED DIPOLE v10 - AUTO OPTIMIZATION")
print("=" * 70)
print(f"  Target: {TARGET_F} GHz, Tolerance: ±{FREQ_TOL*1e3:.0f} MHz")
print(f"  Starting arm: {ARM_START} mm")
print(f"  Max iterations: {MAX_ITER}")
print(f"  balun_L={balun_L:.2f}mm, pad={pad:.2f}mm")
print()

# ============================================================================
# 优化历史跟踪
# ============================================================================
history = []  # [(arm, f_res, s11_min)]


def kill_aedt():
    """强制关闭所有 AEDT 进程"""
    try:
        subprocess.run(["taskkill", "/F", "/IM", "ansysedt.exe"],
                       capture_output=True, timeout=15)
        time.sleep(3)
    except:
        pass


def run_single_simulation(arm_length, iteration):
    """运行单次仿真，返回 (f_res, s11_min, bw_lo, bw_hi, bw_pct, success)"""
    
    print(f"\n{'='*70}")
    print(f"ITERATION {iteration}: arm = {arm_length:.1f} mm")
    print(f"{'='*70}")
    
    # 确保之前的 AEDT 已关闭
    kill_aedt()
    time.sleep(2)
    
    # 重新应用补丁 (每次迭代都需要，因为模块可能被重载)
    try:
        from pyaedt import desktop as _dm
        _o = _dm.Desktop.__init__
        def _p(self, *a, **k):
            _o(self, *a, **k)
            self.student_version = getattr(self, 'student_version', False)
        _dm.Desktop.__init__ = _p
    except: pass
    try:
        import pyaedt.application.Design as _dd
        _o2 = _dd.DesignSettings.__init__
        def _p2(s, a):
            try: _o2(s, a)
            except AttributeError:
                s._app = a; s.design_settings = None; s.manipulate_inputs = None
        _dd.DesignSettings.__init__ = _p2
    except: pass
    
    from pyaedt import Hfss
    
    # 几何计算
    z_top = sub_H
    balun_y_top = -dipole_w / 2
    balun_y_bot = balun_y_top - balun_L
    airbox_ymin = -(sub_W / 2 + pad)
    sub_y_start = airbox_ymin
    sub_y_end = sub_W / 2
    sub_total_y = sub_y_end - sub_y_start
    feed_y_start = airbox_ymin
    feed_y_end = balun_y_bot
    feed_len = feed_y_end - feed_y_start
    air_y_max = sub_y_end + pad
    air_y_total = air_y_max - airbox_ymin
    
    # 检查 arm 长度不超过基板
    if arm_length * 2 + gap > sub_L:
        # 加长基板
        eff_sub_L = arm_length * 2 + gap + 20
        print(f"  NOTE: sub_L extended to {eff_sub_L:.0f}mm for arm={arm_length:.1f}mm")
    else:
        eff_sub_L = sub_L
    
    proj_name = f"Printed_Dipole_v10_iter{iteration}"
    design_name = f"Dipole_iter{iteration}"
    
    result = {'f_res': 0, 's11_min': 0, 'bw_lo': 0, 'bw_hi': 0, 'bw_pct': 0, 'success': False}
    
    try:
        # 1. 启动 AEDT
        print("  Launching AEDT...")
        hfss = Hfss(
            projectname=os.path.join(PROJECT_DIR, proj_name),
            designname=design_name,
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
        
        # 2. 几何建模
        print("  Creating geometry...")
        create_box("Substrate", -eff_sub_L/2, sub_y_start, 0, eff_sub_L, sub_total_y, sub_H, "FR4_epoxy")
        create_box("Ground", -eff_sub_L/2, sub_y_start, -cu_t, eff_sub_L, sub_total_y, cu_t, "copper")
        create_box("Dipole_Left", -(gap/2+arm_length), -dipole_w/2, z_top, arm_length, dipole_w, cu_t, "copper")
        create_box("Dipole_Right", gap/2, -dipole_w/2, z_top, arm_length, dipole_w, cu_t, "copper")
        
        for i in range(balun_segs):
            fs = i / balun_segs; fe = (i + 1) / balun_segs
            avg_w = (balun_w_end - (balun_w_end - balun_w_start) * fs +
                     balun_w_end - (balun_w_end - balun_w_start) * fe) / 2
            seg_len = balun_L / balun_segs
            seg_y = balun_y_top - seg_len * (i + 1)
            create_box(f"Balun_{i}", -avg_w/2, seg_y, z_top, avg_w, seg_len, cu_t, "copper")
        
        create_box("Feed_Line", -balun_w_start/2, feed_y_start, z_top,
                    balun_w_start, feed_len, cu_t, "copper")
        create_box("Air", -(eff_sub_L/2+pad), airbox_ymin, -pad,
                    eff_sub_L+2*pad, air_y_total, sub_H+2*pad, "vacuum", True)
        
        oBnd.AssignRadiation(["NAME:Rad1", "Objects:=", ["Air"],
                              "IsFssReference:=", False, "IsForPML:=", False])
        print("  Geometry OK")
        
        # 3. WavePort
        print("  Setting up port...")
        air_faces = oEditor.GetFaceIDs("Air")
        fi = []
        for fid in air_faces:
            try:
                cc = oEditor.GetFaceCenter(int(fid))
                fi.append((int(fid), float(cc[1])))
            except: pass
        fi.sort(key=lambda x: x[1])
        ymin_face = fi[0][0]
        
        oBnd.AutoIdentifyPorts(["NAME:Faces", ymin_face], True,
            ["NAME:ReferenceConductors", "Ground"], "Port1", True)
        terms = oBnd.GetExcitationsOfType("Terminal")
        terminal = str(terms[0]) if terms else None
        print(f"  Terminal: {terminal}")
        
        if not terminal:
            print("  ERROR: No terminals!")
            hfss.save_project()
            hfss.release_desktop(close_projects=True)
            return result
        
        # 4. Far Field Setup
        oRadField = oDesign.GetModule("RadField")
        ff_name = None
        try:
            oRadField.InsertFarFieldSphereSetup(
                ["NAME:FF1", "UseCustomRadiationSurface:=", False,
                 "ThetaStart:=", "-180deg", "ThetaStop:=", "180deg", "ThetaStep:=", "2deg",
                 "PhiStart:=", "0deg", "PhiStop:=", "360deg", "PhiStep:=", "2deg",
                 "UseLocalCS:=", False])
            ff_name = "FF1"
            print(f"  Far field setup: {ff_name}")
        except Exception as e:
            print(f"  Far field setup failed: {e}")
        
        # 5. Analysis Setup
        oAnalysis = oDesign.GetModule("AnalysisSetup")
        oAnalysis.InsertSetup("HfssDriven",
            ["NAME:Setup1", "Frequency:=", f"{TARGET_F}GHz",
             "MaxDeltaS:=", 0.02, "MaximumPasses:=", 15,
             "MinimumPasses:=", 2, "MinimumConvergedPasses:=", 2,
             "PercentRefinement:=", 30, "IsEnabled:=", True,
             "BasisOrder:=", 1, "UseIterativeSolver:=", False,
             "DoLambdaRefine:=", True, "DoMaterialLambdaRefine:=", True,
             "SetLambdaTarget:=", False, "Target:=", 0.3333])
        oAnalysis.InsertFrequencySweep("Setup1",
            ["NAME:Sweep1", "IsEnabled:=", True, "SetupType:=", "LinearCount",
             "StartValue:=", f"{f_min/1e9}GHz", "StopValue:=", f"{f_max/1e9}GHz",
             "Count:=", n_pts, "Type:=", "Discrete", "SaveFields:=", True,
             "SaveRadFields:=", True])
        print("  Setup OK")
        
        # 6. Validate & Simulate
        hfss.save_project()
        v = oDesign.ValidateDesign()
        print(f"  Validation: {v}")
        
        if v != 1:
            print("  ERROR: Validation failed!")
            hfss.save_project()
            hfss.release_desktop(close_projects=True)
            return result
        
        t0 = time.time()
        print("  Simulating...")
        oDesign.Analyze("Setup1")
        ts = time.time() - t0
        print(f"  Simulation OK! {ts:.0f}s ({ts/60:.1f}min)")
        
        # 7. Extract S11
        oSolutions = oDesign.GetModule("Solutions")
        oReport = oDesign.GetModule("ReportSetup")
        
        tag = f"v10_i{iteration}"
        s1p = os.path.join(PROJECT_DIR, f"dipole_{tag}.s1p")
        s11_csv = os.path.join(PROJECT_DIR, f"s11_{tag}.csv")
        
        try:
            oSolutions.ExportNetworkData("", ["Setup1:Sweep1"], 3, s1p, ["All"], True, 50)
        except Exception as e:
            print(f"  s1p err: {e}")
        
        oReport.CreateReport(f"S11_{tag}", "Terminal Solution Data", "Rectangular Plot",
            "Setup1 : Sweep1", [], ["Freq:=", ["All"]],
            ["X Component:=", "Freq", "Y Component:=",
             [f"dB(St({terminal},{terminal}))"]], [])
        oReport.ExportToFile(f"S11_{tag}", s11_csv)
        print(f"  S11 CSV exported")
        
        # 8. Parse S11
        freqs, s11 = [], []
        with open(s11_csv) as fh:
            rd = csv.reader(fh); next(rd)
            for row in rd:
                if len(row) >= 2:
                    try: freqs.append(float(row[0])); s11.append(float(row[1]))
                    except: pass
        freqs = np.array(freqs); s11 = np.array(s11)
        
        mi = np.argmin(s11)
        f_res = freqs[mi]; s11_min = s11[mi]
        b10 = np.where(s11 < -10)[0]
        if len(b10) >= 2:
            bw_lo = freqs[b10[0]]; bw_hi = freqs[b10[-1]]
            bw = bw_hi - bw_lo; bw_pct = bw / f_res * 100
        else:
            bw_lo = bw_hi = f_res; bw = bw_pct = 0
        
        fe = abs(f_res - TARGET_F) / TARGET_F * 100
        print(f"  Resonance: {f_res:.4f} GHz")
        print(f"  S11 min:   {s11_min:.2f} dB")
        if bw > 0:
            print(f"  -10dB BW:  {bw_lo:.3f}~{bw_hi:.3f} GHz ({bw_pct:.2f}%)")
        print(f"  Freq err:  {fe:.2f}%")
        
        result = {
            'f_res': f_res, 's11_min': s11_min,
            'bw_lo': bw_lo, 'bw_hi': bw_hi, 'bw_pct': bw_pct,
            'success': True, 'freqs': freqs, 's11': s11,
            'terminal': terminal, 'ff_name': ff_name
        }
        
        # 9. 如果是最终迭代或已收敛，提取辐射方向图
        converged = (abs(f_res - TARGET_F) < FREQ_TOL and s11_min < S11_TARGET)
        if converged and ff_name:
            print("  Extracting radiation patterns...")
            e_csv = os.path.join(PROJECT_DIR, "e_plane_v10.csv")
            h_csv = os.path.join(PROJECT_DIR, "h_plane_v10.csv")
            for pn, phi, cp in [("E_Plane", "0deg", e_csv), ("H_Plane", "90deg", h_csv)]:
                for sw in ["Setup1 : Sweep1", "Setup1 : LastAdaptive"]:
                    try:
                        oReport.CreateReport(pn, "Far Fields", "Radiation Pattern", sw,
                            ["Context:=", ff_name],
                            ["Theta:=", ["All"], "Phi:=", [phi],
                             "Freq:=", [f"{TARGET_F}GHz"]],
                            ["X Component:=", "Theta", "Y Component:=", ["GainTotal"]], [])
                        oReport.ExportToFile(pn, cp)
                        print(f"    {pn}: OK"); break
                    except Exception as e:
                        print(f"    {pn} ({sw}): {e}")
            result['e_csv'] = e_csv
            result['h_csv'] = h_csv
        
        # S11 plot for this iteration
        try:
            import matplotlib; matplotlib.use('Agg')
            import matplotlib.pyplot as plt
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.plot(freqs, s11, 'b-', lw=2, label='S11')
            ax.axhline(-10, color='r', ls='--', lw=1, label='-10dB')
            ax.axvline(TARGET_F, color='g', ls=':', lw=1, label=f'Target {TARGET_F}')
            ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'Res {f_res:.3f}')
            if bw > 0:
                ax.axvspan(bw_lo, bw_hi, alpha=0.15, color='green',
                           label=f'BW {bw*1e3:.0f}MHz ({bw_pct:.1f}%)')
            ax.set_xlabel('Frequency (GHz)'); ax.set_ylabel('S11 (dB)')
            ax.set_title(f'v10 iter{iteration} (arm={arm_length:.1f}mm)')
            ax.legend(); ax.grid(True, alpha=0.3)
            ax.set_xlim(f_min/1e9, f_max/1e9)
            fig.savefig(os.path.join(PROJECT_DIR, f"s11_{tag}.png"), dpi=150, bbox_inches='tight')
            plt.close(fig)
            print(f"  Plot saved")
        except Exception as e:
            print(f"  Plot err: {e}")
        
        hfss.save_project()
        hfss.release_desktop(close_projects=True)
        
    except Exception as e:
        print(f"  EXCEPTION: {e}")
        traceback.print_exc()
        try: hfss.release_desktop(close_projects=True)
        except: pass
    
    # 确保 AEDT 完全关闭
    kill_aedt()
    gc.collect()
    time.sleep(5)
    
    return result


# ============================================================================
# 主优化循环
# ============================================================================
current_arm = ARM_START
best_arm = current_arm
best_err = 999.0
best_s11 = 0
best_result = None

for it in range(1, MAX_ITER + 1):
    res = run_single_simulation(current_arm, it)
    
    if not res['success']:
        print(f"\n  Iteration {it} FAILED, retrying with same arm...")
        continue
    
    f_res = res['f_res']
    s11_min = res['s11_min']
    freq_err = abs(f_res - TARGET_F)
    
    history.append((current_arm, f_res, s11_min))
    
    # 更新最佳
    if freq_err < best_err:
        best_err = freq_err
        best_arm = current_arm
        best_s11 = s11_min
        best_result = res
    
    print(f"\n  History so far:")
    for arm, fr, s in history:
        print(f"    arm={arm:.1f}mm -> f={fr:.4f}GHz, S11={s:.2f}dB")
    
    # 检查收敛
    if freq_err < FREQ_TOL and s11_min < S11_TARGET:
        print(f"\n  *** CONVERGED at arm={current_arm:.1f}mm ***")
        print(f"  f_res={f_res:.4f}GHz (err={freq_err*1e3:.0f}MHz)")
        print(f"  S11={s11_min:.2f}dB")
        break
    
    # 计算下一个 arm
    if len(history) >= 2:
        # 用最近两个点线性插值
        a1, f1, _ = history[-2]
        a2, f2, _ = history[-1]
        if abs(f2 - f1) > 0.001:
            slope = (a2 - a1) / (f2 - f1)  # mm/GHz
            next_arm = a2 + slope * (TARGET_F - f2)
        else:
            # 频率变化太小，用简单缩放
            next_arm = current_arm * (f_res / TARGET_F)
    else:
        # 第一次迭代，用简单缩放
        next_arm = current_arm * (f_res / TARGET_F)
    
    # 限制变化幅度 (防止过冲)
    max_change = current_arm * 0.3
    if abs(next_arm - current_arm) > max_change:
        next_arm = current_arm + np.sign(next_arm - current_arm) * max_change
    
    # 限制合理范围
    next_arm = max(30.0, min(150.0, next_arm))
    next_arm = round(next_arm, 1)
    
    print(f"\n  Next arm: {next_arm:.1f}mm")
    current_arm = next_arm

# ============================================================================
# 最终结果汇总
# ============================================================================
print("\n" + "=" * 70)
print("OPTIMIZATION COMPLETE")
print("=" * 70)
print(f"  Iterations: {len(history)}")
print(f"  Best arm:   {best_arm:.1f} mm")
print(f"  Best f_res: {history[-1][1] if history else 'N/A'} GHz")
print(f"  Best S11:   {best_s11:.2f} dB")
print()
print("  Full history:")
for i, (arm, fr, s) in enumerate(history, 1):
    err = abs(fr - TARGET_F) / TARGET_F * 100
    print(f"    [{i}] arm={arm:.1f}mm -> f={fr:.4f}GHz (err={err:.2f}%), S11={s:.2f}dB")

# 生成最终 S11 汇总图
try:
    import matplotlib; matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    
    fig, ax = plt.subplots(figsize=(12, 7))
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']
    
    for i, (arm, fr, s) in enumerate(history):
        tag = f"v10_i{i+1}"
        csv_path = os.path.join(PROJECT_DIR, f"s11_{tag}.csv")
        if os.path.exists(csv_path):
            fs, ss = [], []
            with open(csv_path) as fh:
                rd = csv.reader(fh); next(rd)
                for row in rd:
                    if len(row) >= 2:
                        try: fs.append(float(row[0])); ss.append(float(row[1]))
                        except: pass
            c_idx = i % len(colors)
            ax.plot(fs, ss, color=colors[c_idx], lw=2,
                    label=f'arm={arm:.1f}mm (f={fr:.3f}, S11={s:.1f}dB)')
    
    ax.axhline(-10, color='r', ls='--', lw=1, alpha=0.7, label='-10dB')
    ax.axvline(TARGET_F, color='k', ls=':', lw=1.5, label=f'Target {TARGET_F}GHz')
    ax.set_xlabel('Frequency (GHz)', fontsize=12)
    ax.set_ylabel('S11 (dB)', fontsize=12)
    ax.set_title('Printed Dipole v10 - Optimization History', fontsize=14)
    ax.legend(fontsize=9); ax.grid(True, alpha=0.3)
    ax.set_xlim(f_min/1e9, f_max/1e9)
    fig.savefig(os.path.join(PROJECT_DIR, "s11_v10_summary.png"), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print("  Summary plot saved: s11_v10_summary.png")
except Exception as e:
    print(f"  Summary plot err: {e}")

# 辐射方向图 (如果有数据)
try:
    e_csv = os.path.join(PROJECT_DIR, "e_plane_v10.csv")
    h_csv = os.path.join(PROJECT_DIR, "h_plane_v10.csv")
    if os.path.exists(e_csv) and os.path.exists(h_csv):
        def read_pat(p):
            t, g = [], []
            with open(p) as f:
                rd = csv.reader(f); next(rd)
                for row in rd:
                    if len(row) >= 2:
                        try: t.append(float(row[0])); g.append(float(row[1]))
                        except: pass
            return np.array(t), np.array(g)
        
        te, ge = read_pat(e_csv)
        th, gh = read_pat(h_csv)
        if len(ge) > 0 and len(gh) > 0:
            me = np.max(ge); mh = np.max(gh)
            max_gain = max(me, mh)
            print(f"  E-plane gain: {me:.2f} dBi")
            print(f"  H-plane gain: {mh:.2f} dBi")
            print(f"  Max gain:     {max_gain:.2f} dBi")
            
            fig, (a1, a2) = plt.subplots(1, 2, subplot_kw={'projection': 'polar'}, figsize=(14, 6))
            a1.plot(np.radians(te), ge, 'b-', lw=2)
            a1.set_title(f'E-Plane (phi=0)\nGain={me:.2f}dBi', pad=20)
            a1.set_theta_zero_location('N'); a1.set_theta_direction(-1)
            a2.plot(np.radians(th), gh, 'r-', lw=2)
            a2.set_title(f'H-Plane (phi=90)\nGain={mh:.2f}dBi', pad=20)
            a2.set_theta_zero_location('N'); a2.set_theta_direction(-1)
            fig.suptitle(f'Printed Dipole v10 @ {TARGET_F} GHz', fontsize=14)
            fig.savefig(os.path.join(PROJECT_DIR, "radiation_v10.png"), dpi=150, bbox_inches='tight')
            plt.close(fig)
            print("  Radiation plot saved")
    else:
        print("  No radiation pattern data (extraction may have failed)")
        print("  Theoretical gain for half-wave printed dipole: ~2.15 dBi")
except Exception as e:
    print(f"  Radiation err: {e}")

print("\nDONE")
