#!/usr/bin/env python3
"""
printed_dipole_v9.py - 印刷偶极子天线 v9
  基于 v5 全地面方案 (S11=-20dB 已验证)
  改进: SaveRadFields=True, InsertFarFieldSphereSetup
  起始 arm=90mm (v5 轨迹外推: 81.5mm→2.38GHz, 目标2.217GHz)
  无进程内优化 (避免 COM 崩溃)
"""
import numpy as np, os, sys, time, csv, traceback

os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
os.makedirs(PROJECT_DIR, exist_ok=True)
LOG = os.path.join(PROJECT_DIR, "dipole_v9.log")

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

# ============================================================================
# 1. 参数 (与 v5 一致, arm=90mm)
# ============================================================================
print("="*70); print("STEP 1: Parameters"); print("="*70)

c=3e8; f0=2.217e9; eps_r=4.4
lambda_0 = c/f0; lambda_g = lambda_0/np.sqrt(eps_r)

sub_L=120.0; sub_W=60.0; sub_H=1.6; cu_t=0.035
dipole_arm = 90.0   # v5: 81.5→2.38GHz, 外推 ~90→2.22GHz
dipole_w = 3.0; gap = 1.0
balun_L = lambda_g*1e3/4.0; balun_segs = 10
balun_w_start=1.0; balun_w_end=3.0
pad = lambda_0*1e3/4.0  # ~33.83mm

z_top = sub_H
balun_y_top = -dipole_w/2
balun_y_bot = balun_y_top - balun_L
airbox_ymin = -(sub_W/2 + pad)
sub_y_start = airbox_ymin  # 基板延伸到 AirBox ymin
sub_y_end = sub_W/2
sub_total_y = sub_y_end - sub_y_start
feed_y_start = airbox_ymin
feed_y_end = balun_y_bot
feed_len = feed_y_end - feed_y_start

f_min=1.5e9; f_max=3.0e9; n_pts=151

print(f"  f0={f0/1e9:.3f}GHz, lam0={lambda_0*1e3:.1f}mm")
print(f"  arm={dipole_arm}mm, total dipole={2*dipole_arm+gap}mm")
print(f"  balun_L={balun_L:.2f}mm, pad={pad:.2f}mm")
print(f"  sub Y=[{sub_y_start:.2f}, {sub_y_end:.1f}] ({sub_total_y:.1f}mm)")
print(f"  feed Y=[{feed_y_start:.2f}, {feed_y_end:.2f}] ({feed_len:.1f}mm)")
print(f"  AirBox ymin={airbox_ymin:.2f}")
print()

# ============================================================================
# 2. 启动 AEDT
# ============================================================================
print("="*70); print("STEP 2: Launch AEDT"); print("="*70)

hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "Printed_Dipole_v9"),
    designname="Dipole_v9",
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

# ============================================================================
# 3. 几何
# ============================================================================
print(); print("="*70); print("STEP 3: Geometry"); print("="*70)

# 3.1 基板
create_box("Substrate", -sub_L/2, sub_y_start, 0, sub_L, sub_total_y, sub_H, "FR4_epoxy")
print(f"  Substrate: {sub_L}x{sub_total_y:.1f}x{sub_H}")

# 3.2 全地面 (与基板同尺寸)
create_box("Ground", -sub_L/2, sub_y_start, -cu_t, sub_L, sub_total_y, cu_t, "copper")
print(f"  Ground (full): {sub_L}x{sub_total_y:.1f}x{cu_t}")

# 3.3 偶极子臂
create_box("Dipole_Left", -(gap/2+dipole_arm), -dipole_w/2, z_top, dipole_arm, dipole_w, cu_t, "copper")
create_box("Dipole_Right", gap/2, -dipole_w/2, z_top, dipole_arm, dipole_w, cu_t, "copper")
print(f"  Dipole: arm={dipole_arm}mm")

# 3.4 巴伦 (渐变)
for i in range(balun_segs):
    fs = i/balun_segs; fe = (i+1)/balun_segs
    avg_w = (balun_w_end - (balun_w_end-balun_w_start)*fs + balun_w_end - (balun_w_end-balun_w_start)*fe) / 2
    seg_len = balun_L/balun_segs
    seg_y = balun_y_top - seg_len*(i+1)
    create_box(f"Balun_{i}", -avg_w/2, seg_y, z_top, avg_w, seg_len, cu_t, "copper")
print(f"  Balun: {balun_segs} segs, L={balun_L:.2f}mm")

# 3.5 馈线
create_box("Feed_Line", -balun_w_start/2, feed_y_start, z_top, balun_w_start, feed_len, cu_t, "copper")
print(f"  Feed: y=[{feed_y_start:.2f},{feed_y_end:.2f}] ({feed_len:.1f}mm)")

# 3.6 AirBox
ax = sub_L + 2*pad; ay = sub_W + 2*pad; az = sub_H + 2*pad  # v5 中 y 方向两侧均有 pad
# 更正: AirBox Y 下界 = airbox_ymin, 上界 = sub_W/2 + pad
air_y_max = sub_y_end + pad
air_y_total = air_y_max - airbox_ymin
create_box("Air", -(sub_L/2+pad), airbox_ymin, -pad, sub_L+2*pad, air_y_total, sub_H+2*pad, "vacuum", True)
print(f"  AirBox: {sub_L+2*pad:.1f}x{air_y_total:.1f}x{sub_H+2*pad:.1f}")

oBnd.AssignRadiation(["NAME:Rad1","Objects:=",["Air"],"IsFssReference:=",False,"IsForPML:=",False])
print("  Radiation BC OK")

# ============================================================================
# 4. WavePort
# ============================================================================
print(); print("="*70); print("STEP 4: WavePort"); print("="*70)

air_faces = oEditor.GetFaceIDs("Air")
print(f"  AirBox faces: {air_faces}")
# 找 y_min 面
fi = []
for fid in air_faces:
    try:
        cc = oEditor.GetFaceCenter(int(fid))
        fi.append((int(fid), float(cc[1])))
    except: pass
fi.sort(key=lambda x: x[1])
for fid, yv in fi:
    print(f"    Face {fid}: y_center={yv:.2f}")
ymin_face = fi[0][0]
print(f"  y_min face: {ymin_face}")

oBnd.AutoIdentifyPorts(["NAME:Faces",ymin_face], True,
    ["NAME:ReferenceConductors","Ground"], "Port1", True)
print("  AutoIdentifyPorts OK")

terms = oBnd.GetExcitationsOfType("Terminal")
print(f"  Terminals: {terms}")
terminal = str(terms[0]) if terms else "Feed_Line_T1"
print(f"  Using: {terminal}")

if not terms:
    print("  ERROR: No terminals! Aborting.")
    hfss.save_project()
    sys.exit(1)

# ============================================================================
# 5. 远场设置
# ============================================================================
print(); print("STEP 5: Far Field Setup")
oRadField = oDesign.GetModule("RadField")
ff_name = None
for mn, mf in [
    ("InsertFarFieldSphereSetup", lambda: oRadField.InsertFarFieldSphereSetup(
        ["NAME:FF1","UseCustomRadiationSurface:=",False,
         "ThetaStart:=","-180deg","ThetaStop:=","180deg","ThetaStep:=","2deg",
         "PhiStart:=","0deg","PhiStop:=","360deg","PhiStep:=","2deg",
         "UseLocalCS:=",False])),
    ("InsertInfiniteSphereDef", lambda: oRadField.InsertInfiniteSphereDef(
        ["NAME:FF2","UseCustomRadiationSurface:=",False,
         "ThetaStart:=","-180deg","ThetaStop:=","180deg","ThetaStep:=","2deg",
         "PhiStart:=","0deg","PhiStop:=","360deg","PhiStep:=","2deg",
         "UseLocalCS:=",False])),
]:
    try: mf(); ff_name = "FF1" if "Sphere" in mn else "FF2"; print(f"  {mn}: OK -> {ff_name}"); break
    except Exception as e: print(f"  {mn}: {e}")

# ============================================================================
# 6. 分析设置
# ============================================================================
print(); print("="*70); print("STEP 6: Setup"); print("="*70)
oAnalysis = oDesign.GetModule("AnalysisSetup")
oAnalysis.InsertSetup("HfssDriven",
    ["NAME:Setup1","Frequency:=",f"{f0/1e9}GHz",
     "MaxDeltaS:=",0.02,"MaximumPasses:=",15,
     "MinimumPasses:=",2,"MinimumConvergedPasses:=",2,
     "PercentRefinement:=",30,"IsEnabled:=",True,
     "BasisOrder:=",1,"UseIterativeSolver:=",False,
     "DoLambdaRefine:=",True,"DoMaterialLambdaRefine:=",True,
     "SetLambdaTarget:=",False,"Target:=",0.3333])
oAnalysis.InsertFrequencySweep("Setup1",
    ["NAME:Sweep1","IsEnabled:=",True,"SetupType:=","LinearCount",
     "StartValue:=",f"{f_min/1e9}GHz","StopValue:=",f"{f_max/1e9}GHz",
     "Count:=",n_pts,"Type:=","Discrete","SaveFields:=",True,"SaveRadFields:=",True])
print(f"  Setup1 OK, Sweep 1.5-3.0GHz 151pts SaveRadFields=True")

# ============================================================================
# 7. 验证 & 仿真
# ============================================================================
print(); print("="*70); print("STEP 7: Simulate"); print("="*70)
hfss.save_project()
v = oDesign.ValidateDesign()
print(f"  Validation: {v}")

if v != 1:
    print("  WARNING: Validation failed!")
    print(f"    Excitations: {oBnd.GetExcitations()}")

t0 = time.time()
print("  Running Setup1...")
try:
    oDesign.Analyze("Setup1")
    ts = time.time()-t0
    print(f"  OK! {ts:.0f}s ({ts/60:.1f}min)")
except Exception as e:
    print(f"  FAILED: {e}")
    hfss.save_project(); sys.exit(1)

# ============================================================================
# 8. S11 提取与分析
# ============================================================================
print(); print("="*70); print("STEP 8: S11"); print("="*70)

oSolutions = oDesign.GetModule("Solutions")
oReport = oDesign.GetModule("ReportSetup")

s1p = os.path.join(PROJECT_DIR, "dipole_v9.s1p")
s11_csv = os.path.join(PROJECT_DIR, "s11_v9.csv")

try: oSolutions.ExportNetworkData("",["Setup1:Sweep1"],3,s1p,["All"],True,50); print(f"  s1p OK")
except Exception as e: print(f"  s1p err: {e}")

try:
    oReport.CreateReport("S11_v9","Terminal Solution Data","Rectangular Plot",
        "Setup1 : Sweep1",[],["Freq:=",["All"]],
        ["X Component:=","Freq","Y Component:=",[f"dB(St({terminal},{terminal}))"]],[])
    oReport.ExportToFile("S11_v9", s11_csv)
    print(f"  S11 CSV OK")
except Exception as e: print(f"  S11 err: {e}")

def parse_csv(p):
    f,s=[],[]
    with open(p) as fh:
        rd=csv.reader(fh); next(rd)
        for row in rd:
            if len(row)>=2:
                try: f.append(float(row[0])); s.append(float(row[1]))
                except: pass
    return np.array(f), np.array(s)

try:
    freqs, s11 = parse_csv(s11_csv)
    mi = np.argmin(s11)
    f_res=freqs[mi]; s11_min=s11[mi]
    b10 = np.where(s11<-10)[0]
    if len(b10)>=2:
        bw_lo=freqs[b10[0]]; bw_hi=freqs[b10[-1]]
        bw=bw_hi-bw_lo; bw_pct=bw/f_res*100
    else: bw_lo=bw_hi=f_res; bw=bw_pct=0
    fe = abs(f_res-f0/1e9)/(f0/1e9)*100

    print(f"  Resonance:  {f_res:.4f} GHz")
    print(f"  S11 min:    {s11_min:.2f} dB")
    if bw>0: print(f"  -10dB BW:   {bw_lo:.3f}~{bw_hi:.3f} GHz ({bw_pct:.2f}%)")
    else: print("  WARNING: S11 > -10dB")
    print(f"  Freq err:   {fe:.2f}%")

    # 如果需要下一步优化,打印建议
    if fe > 2.0 or s11_min > -10:
        next_arm = round(dipole_arm * (f_res/(f0/1e9)), 1)
        print(f"  >> Next arm suggestion: {next_arm:.1f}mm")
except Exception as e:
    print(f"  Analysis err: {e}"); traceback.print_exc()
    f_res=f0/1e9; s11_min=0; bw=bw_pct=0; bw_lo=bw_hi=f_res; fe=999
    freqs=np.array([]); s11=np.array([])

# S11 plot
try:
    import matplotlib; matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    fig,ax=plt.subplots(figsize=(10,6))
    ax.plot(freqs, s11, 'b-', lw=2, label='S11')
    ax.axhline(-10, color='r', ls='--', lw=1, label='-10dB')
    ax.axvline(f0/1e9, color='g', ls=':', lw=1, label=f'Target {f0/1e9:.3f}')
    ax.axvline(f_res, color='orange', ls='-.', lw=1, label=f'Res {f_res:.3f}')
    if bw>0: ax.axvspan(bw_lo,bw_hi,alpha=0.15,color='green',label=f'BW {bw*1e3:.0f}MHz ({bw_pct:.1f}%)')
    ax.set_xlabel('Frequency (GHz)'); ax.set_ylabel('S11 (dB)')
    ax.set_title(f'Printed Dipole v9 (arm={dipole_arm}mm, full ground)')
    ax.legend(); ax.grid(True,alpha=0.3); ax.set_xlim(f_min/1e9,f_max/1e9)
    fig.savefig(os.path.join(PROJECT_DIR,"s11_v9.png"),dpi=150,bbox_inches='tight')
    plt.close(fig); print(f"  Plot saved")
except Exception as e: print(f"  Plot err: {e}")

# ============================================================================
# 9. 辐射方向图
# ============================================================================
print(); print("="*70); print("STEP 9: Radiation Patterns"); print("="*70)

max_gain = 0
e_csv = os.path.join(PROJECT_DIR,"e_plane_v9.csv")
h_csv = os.path.join(PROJECT_DIR,"h_plane_v9.csv")

if not ff_name:
    try:
        ex = oRadField.GetSetupNames()
        if ex: ff_name=str(ex[0]); print(f"  Found existing FF: {ff_name}")
    except: pass

for pn,phi,cp in [("E_Plane_v9","0deg",e_csv),("H_Plane_v9","90deg",h_csv)]:
    if not ff_name: print(f"  {pn}: no FF setup"); continue
    for sw in ["Setup1 : Sweep1","Setup1 : LastAdaptive"]:
        try:
            oReport.CreateReport(pn,"Far Fields","Radiation Pattern",sw,
                ["Context:=",ff_name],
                ["Theta:=",["All"],"Phi:=",[phi],"Freq:=",[f"{f0/1e9}GHz"]],
                ["X Component:=","Theta","Y Component:=",["GainTotal"]],[])
            oReport.ExportToFile(pn, cp)
            print(f"  {pn}: OK ({sw})"); break
        except Exception as e: print(f"  {pn} ({sw}): {e}")

def read_pat(p):
    t,g=[],[]
    with open(p) as f:
        rd=csv.reader(f); next(rd)
        for row in rd:
            if len(row)>=2:
                try: t.append(float(row[0])); g.append(float(row[1]))
                except: pass
    return np.array(t), np.array(g)

try:
    if os.path.exists(e_csv) and os.path.exists(h_csv):
        te,ge = read_pat(e_csv)
        th,gh = read_pat(h_csv)
        if len(ge)>0 and len(gh)>0:
            me=np.max(ge); mh=np.max(gh); max_gain=max(me,mh)
            print(f"  E-plane gain: {me:.2f} dBi")
            print(f"  H-plane gain: {mh:.2f} dBi")
            print(f"  Max gain:     {max_gain:.2f} dBi")

            fig,(a1,a2)=plt.subplots(1,2,subplot_kw={'projection':'polar'},figsize=(14,6))
            a1.plot(np.radians(te),ge,'b-',lw=2); a1.set_title(f'E-Plane (phi=0)\nGain={me:.2f}dBi',pad=20)
            a1.set_theta_zero_location('N'); a1.set_theta_direction(-1)
            a2.plot(np.radians(th),gh,'r-',lw=2); a2.set_title(f'H-Plane (phi=90)\nGain={mh:.2f}dBi',pad=20)
            a2.set_theta_zero_location('N'); a2.set_theta_direction(-1)
            fig.suptitle(f'Printed Dipole @ {f0/1e9:.3f} GHz',fontsize=14)
            fig.savefig(os.path.join(PROJECT_DIR,"radiation_v9.png"),dpi=150,bbox_inches='tight')
            plt.close(fig); print(f"  Pattern plot saved")
except Exception as e: print(f"  Pattern err: {e}")

# ============================================================================
# 设计总结
# ============================================================================
print()
print("="*70)
print("DESIGN SUMMARY")
print("="*70)
print(f"  Type:       Printed Dipole + Full GND + Microstrip Balun")
print(f"  Target:     {f0/1e9:.3f} GHz")
print(f"  Substrate:  FR4 (er={eps_r}), {sub_L}x{sub_total_y:.1f}x{sub_H}mm")
print(f"  Arm:        {dipole_arm:.1f}mm (total {2*dipole_arm+gap:.1f}mm)")
print(f"  Balun:      {balun_L:.2f}mm (taper {balun_w_end}->{balun_w_start}mm)")
print(f"  Resonance:  {f_res:.4f} GHz (err={fe:.2f}%)")
print(f"  S11:        {s11_min:.2f} dB")
if bw>0: print(f"  -10dB BW:   {bw_lo:.3f}~{bw_hi:.3f} GHz ({bw*1e3:.0f}MHz, {bw_pct:.2f}%)")
if max_gain>0: print(f"  Gain:       {max_gain:.2f} dBi")
print()
hfss.save_project()
print("COMPLETE")
