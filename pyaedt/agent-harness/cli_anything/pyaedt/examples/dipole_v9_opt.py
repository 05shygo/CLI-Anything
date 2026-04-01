#!/usr/bin/env python3
"""
dipole_v9_opt.py - 基于 v9 项目进行臂长优化
连接到已有 AEDT 进程, 在 Printed_Dipole_v9 项目中迭代修改臂长
目标: 谐振频率 2.217 GHz, S11 < -10 dB
"""
import numpy as np, os, sys, time, csv, traceback

os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "dipole_v9_opt.log")

class Logger:
    def __init__(self, p):
        self.t = sys.stdout; self.f = open(p, "w", encoding="utf-8")
    def write(self, m):
        try: self.t.write(m)
        except UnicodeEncodeError: self.t.write(m.encode('ascii','replace').decode())
        self.f.write(m); self.f.flush()
    def flush(self): self.t.flush(); self.f.flush()
sys.stdout = Logger(LOG); sys.stderr = sys.stdout

# Patches
try:
    from pyaedt import desktop as _dm
    _o = _dm.Desktop.__init__
    def _p(self, *a, **k): _o(self, *a, **k); self.student_version = getattr(self, 'student_version', False)
    _dm.Desktop.__init__ = _p
except: pass
try:
    import pyaedt.application.Design as _dd
    _o2 = _dd.DesignSettings.__init__
    def _p2(s, a):
        try: _o2(s, a)
        except AttributeError: s._app=a; s.design_settings=None; s.manipulate_inputs=None
    _dd.DesignSettings.__init__ = _p2
except: pass

from pyaedt import Hfss

print("="*70)
print("Dipole v9 Optimization")
print("="*70)

# 连接到已有 AEDT
hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "Printed_Dipole_v9"),
    designname="Dipole_v9",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)
print(f"  Connected to {hfss.project_name}")

hfss.modeler.model_units = "mm"
oDesign = hfss.odesign
oEditor = oDesign.SetActiveEditor("3D Modeler")
oBnd = oDesign.GetModule("BoundarySetup")
oReport = oDesign.GetModule("ReportSetup")

# 获取 terminal name
terms = oBnd.GetExcitationsOfType("Terminal")
terminal = str(terms[0]) if terms else "Feed_Line_T1"
print(f"  Terminal: {terminal}")

# 参数 (与 v9 一致)
f0 = 2.217e9
cu_t = 0.035; z_top = 1.6; dipole_w = 3.0; gap = 1.0

def create_box(name, x, y, z, dx, dy, dz, mat, solve_inside=None):
    if solve_inside is None:
        solve_inside = mat.lower() not in ("copper","pec","aluminum")
    oEditor.CreateBox(
        ["NAME:BoxParameters",
         "XPosition:=",f"{x}mm","YPosition:=",f"{y}mm","ZPosition:=",f"{z}mm",
         "XSize:=",f"{dx}mm","YSize:=",f"{dy}mm","ZSize:=",f"{dz}mm"],
        ["NAME:Attributes",
         "Name:=",name,"Flags:=","","Color:=","(143 175 131)",
         "Transparency:=",0,"PartCoordinateSystem:=","Global","UDMId:=","",
         "MaterialValue:=",f'"{mat}"',"SurfaceMaterialValue:=",'""',
         "SolveInside:=",solve_inside,"IsMaterialEditable:=",True,
         "UseMaterialAppearance:=",False,"IsLightweight:=",False])

def parse_csv(p):
    f,s=[],[]
    with open(p) as fh:
        rd=csv.reader(fh); next(rd)
        for row in rd:
            if len(row)>=2:
                try: f.append(float(row[0])); s.append(float(row[1]))
                except: pass
    return np.array(f), np.array(s)

def analyze(freqs, s11, target):
    mi = np.argmin(s11)
    f_res=freqs[mi]; s11_min=s11[mi]
    b10 = np.where(s11<-10)[0]
    if len(b10)>=2:
        bw_lo=freqs[b10[0]]; bw_hi=freqs[b10[-1]]
        bw=bw_hi-bw_lo; bw_pct=bw/f_res*100
    else: bw_lo=bw_hi=f_res; bw=bw_pct=0
    fe = abs(f_res-target)/target*100
    return dict(f_res=f_res, s11_min=s11_min, bw_lo=bw_lo, bw_hi=bw_hi,
                bw=bw, bw_pct=bw_pct, freq_err=fe)

# v9 结果: arm=90mm, f_res=2.66GHz
# 频率缩放: new_arm = 90 * (2.66/2.217) = 108.0
current_arm = 90.0
current_f_res = 2.66
target_f = f0 / 1e9

# 迭代优化
best_arm = current_arm; best_s11 = -16.77; best_f = 2.66
best_freqs = np.array([]); best_s11_data = np.array([])
best_bw = 0; best_bw_pct = 0; best_bw_lo = 0; best_bw_hi = 0

for it in range(8):
    # 频率缩放
    new_arm = round(current_arm * (current_f_res / target_f), 1)
    
    if abs(new_arm - current_arm) < 0.5:
        print(f"\nIter {it+1}: converged (delta < 0.5mm)")
        break
    if new_arm > 150 or new_arm < 30:
        print(f"\nIter {it+1}: arm {new_arm}mm out of range [30,150]")
        break
    
    print(f"\nIter {it+1}: arm {current_arm:.1f} -> {new_arm:.1f} mm")
    
    # 删除旧臂
    try:
        oEditor.Delete(["NAME:Selections", "Selections:=", "Dipole_Left,Dipole_Right"])
    except Exception as e:
        print(f"  Delete error: {e}"); break
    
    # 创建新臂
    create_box("Dipole_Left", -(gap/2+new_arm), -dipole_w/2, z_top, new_arm, dipole_w, cu_t, "copper")
    create_box("Dipole_Right", gap/2, -dipole_w/2, z_top, new_arm, dipole_w, cu_t, "copper")
    print(f"  New arms OK (total={2*new_arm+gap:.1f}mm)")
    
    hfss.save_project()
    
    # 仿真
    t0 = time.time()
    print("  Analyzing...")
    try:
        oDesign.Analyze("Setup1")
        print(f"  Done ({time.time()-t0:.0f}s)")
    except Exception as e:
        print(f"  FAILED: {e}"); break
    
    # 提取 S11
    csv_path = os.path.join(PROJECT_DIR, f"s11_v9_opt{it+1}.csv")
    rname = f"S11_Opt{it+1}"
    try:
        oReport.CreateReport(rname, "Terminal Solution Data", "Rectangular Plot",
            "Setup1 : Sweep1", [], ["Freq:=", ["All"]],
            ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal},{terminal}))"]],
            [])
        oReport.ExportToFile(rname, csv_path)
        
        freqs, s11 = parse_csv(csv_path)
        r = analyze(freqs, s11, target_f)
        
        print(f"  f_res={r['f_res']:.4f}GHz, S11={r['s11_min']:.2f}dB, err={r['freq_err']:.2f}%")
        if r['bw'] > 0:
            print(f"  -10dB BW: {r['bw_lo']:.3f}~{r['bw_hi']:.3f}GHz ({r['bw_pct']:.1f}%)")
        
        current_arm = new_arm
        current_f_res = r['f_res']
        
        # 更新 best
        if r['s11_min'] < best_s11 or (r['freq_err'] < abs(best_f - target_f)/target_f*100):
            best_arm = new_arm; best_s11 = r['s11_min']; best_f = r['f_res']
            best_freqs = freqs; best_s11_data = s11
            best_bw = r['bw']; best_bw_pct = r['bw_pct']
            best_bw_lo = r['bw_lo']; best_bw_hi = r['bw_hi']
        
        if r['freq_err'] <= 2.0 and r['s11_min'] < -10:
            print(f"  ** CONVERGED! **")
            break
        
        if r['freq_err'] <= 2.0 and r['s11_min'] > -10:
            print(f"  Freq OK but S11={r['s11_min']:.2f}dB (> -10dB)")
            # 微调
            for delta in [-1, +1, -2, +2, -3, +3]:
                trial = round(new_arm + delta, 1)
                print(f"  Fine: arm={trial:.1f}")
                try:
                    oEditor.Delete(["NAME:Selections", "Selections:=", "Dipole_Left,Dipole_Right"])
                    create_box("Dipole_Left", -(gap/2+trial), -dipole_w/2, z_top, trial, dipole_w, cu_t, "copper")
                    create_box("Dipole_Right", gap/2, -dipole_w/2, z_top, trial, dipole_w, cu_t, "copper")
                    hfss.save_project()
                    oDesign.Analyze("Setup1")
                    fc = os.path.join(PROJECT_DIR, f"s11_v9_fine_{trial}.csv")
                    fn = f"S11_Fine_{trial}"
                    oReport.CreateReport(fn, "Terminal Solution Data", "Rectangular Plot",
                        "Setup1 : Sweep1", [], ["Freq:=", ["All"]],
                        ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal},{terminal}))"]],
                        [])
                    oReport.ExportToFile(fn, fc)
                    ff, ss = parse_csv(fc)
                    rf = analyze(ff, ss, target_f)
                    print(f"    f={rf['f_res']:.4f}, S11={rf['s11_min']:.2f}")
                    if rf['s11_min'] < best_s11:
                        best_arm = trial; best_s11 = rf['s11_min']; best_f = rf['f_res']
                        best_freqs = ff; best_s11_data = ss
                        best_bw = rf['bw']; best_bw_pct = rf['bw_pct']
                        best_bw_lo = rf['bw_lo']; best_bw_hi = rf['bw_hi']
                    if rf['s11_min'] < -10 and rf['freq_err'] <= 3:
                        current_arm = trial; current_f_res = rf['f_res']
                        print(f"    ** Found match! **")
                        break
                except Exception as ef:
                    print(f"    Error: {ef}"); continue
            break
            
    except Exception as e:
        print(f"  Extract error: {e}"); traceback.print_exc(); break

# 最终 S11 图
print()
print("="*70)
print("FINAL RESULTS")
print("="*70)
print(f"  Best arm:    {best_arm:.1f}mm (total {2*best_arm+gap:.1f}mm)")
print(f"  Resonance:   {best_f:.4f} GHz")
print(f"  S11 min:     {best_s11:.2f} dB")
if best_bw > 0:
    print(f"  -10dB BW:    {best_bw_lo:.3f}~{best_bw_hi:.3f} GHz ({best_bw*1e3:.0f}MHz, {best_bw_pct:.1f}%)")
print(f"  Freq err:    {abs(best_f-target_f)/target_f*100:.2f}%")

if len(best_freqs) > 0:
    try:
        import matplotlib; matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.plot(best_freqs, best_s11_data, 'b-', lw=2, label=f'S11 (arm={best_arm:.1f}mm)')
        ax.axhline(-10, color='r', ls='--', lw=1, label='-10 dB')
        ax.axvline(target_f, color='g', ls=':', lw=1, label=f'Target {target_f:.3f} GHz')
        ax.axvline(best_f, color='orange', ls='-.', lw=1, label=f'Res {best_f:.3f} GHz')
        if best_bw > 0:
            ax.axvspan(best_bw_lo, best_bw_hi, alpha=0.15, color='green',
                       label=f'BW {best_bw*1e3:.0f}MHz ({best_bw_pct:.1f}%)')
        ax.set_xlabel('Frequency (GHz)'); ax.set_ylabel('S11 (dB)')
        ax.set_title(f'Printed Dipole Optimized (arm={best_arm:.1f}mm)')
        ax.legend(); ax.grid(True, alpha=0.3); ax.set_xlim(1.5, 3.0)
        fig.savefig(os.path.join(PROJECT_DIR, "s11_v9_opt_final.png"), dpi=150, bbox_inches='tight')
        plt.close(fig)
        print(f"  S11 plot saved: s11_v9_opt_final.png")
    except Exception as e:
        print(f"  Plot error: {e}")

# 远场方向图 (如果有远场设置)
print()
print("Attempting radiation patterns...")
oRadField = oDesign.GetModule("RadField")
ff_name = None
try:
    existing = oRadField.GetSetupNames()
    print(f"  FF setups: {existing}")
    if existing:
        ff_name = str(existing[0])
except: pass

if ff_name:
    e_csv = os.path.join(PROJECT_DIR, "e_plane_v9.csv")
    h_csv = os.path.join(PROJECT_DIR, "h_plane_v9.csv")
    for pn, phi, cp in [("E_Plane_v9","0deg",e_csv), ("H_Plane_v9","90deg",h_csv)]:
        for sr in ["Setup1 : Sweep1", "Setup1 : LastAdaptive"]:
            try:
                oReport.CreateReport(pn, "Far Fields", "Radiation Pattern", sr,
                    ["Context:=", ff_name],
                    ["Theta:=",["All"],"Phi:=",[phi],"Freq:=",[f"{target_f}GHz"]],
                    ["X Component:=","Theta","Y Component:=",["GainTotal"]],[])
                oReport.ExportToFile(pn, cp)
                print(f"  {pn} OK ({sr})")
                break
            except Exception as e:
                print(f"  {pn} ({sr}): {e}")
    
    # Read & plot
    def read_pat(p):
        th,g=[],[]
        with open(p) as f:
            rd=csv.reader(f); next(rd)
            for row in rd:
                if len(row)>=2:
                    try: th.append(float(row[0])); g.append(float(row[1]))
                    except: pass
        return np.array(th), np.array(g)
    
    try:
        if os.path.exists(e_csv) and os.path.exists(h_csv):
            te,ge = read_pat(e_csv); th,gh = read_pat(h_csv)
            if len(ge)>0 and len(gh)>0:
                mg = max(np.max(ge), np.max(gh))
                print(f"  E max gain: {np.max(ge):.2f} dBi")
                print(f"  H max gain: {np.max(gh):.2f} dBi")
                print(f"  Antenna gain: {mg:.2f} dBi")
                fig,(ax1,ax2) = plt.subplots(1,2,subplot_kw={'projection':'polar'},figsize=(14,6))
                ax1.plot(np.radians(te),ge,'b-',lw=2); ax1.set_title(f'E-Plane\n{np.max(ge):.2f}dBi',pad=20)
                ax2.plot(np.radians(th),gh,'r-',lw=2); ax2.set_title(f'H-Plane\n{np.max(gh):.2f}dBi',pad=20)
                fig.suptitle(f'Dipole@{best_f:.3f}GHz, arm={best_arm:.1f}mm',fontsize=14)
                fig.savefig(os.path.join(PROJECT_DIR,"radiation_v9.png"),dpi=150,bbox_inches='tight')
                plt.close(fig)
                print(f"  Pattern plot saved")
    except Exception as e:
        print(f"  Pattern error: {e}")
else:
    print("  No FF setup found. Estimating gain ~2.15 dBi (half-wave dipole nominal)")

hfss.save_project()
print()
print("OPTIMIZATION COMPLETE")
print("="*70)
