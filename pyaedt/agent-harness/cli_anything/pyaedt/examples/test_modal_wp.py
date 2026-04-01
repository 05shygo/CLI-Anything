"""Test: DrivenModal + WavePort - no terminals needed, fix geometry."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "modal_wp_test.log")

class DL:
    def __init__(self, p):
        self.t = sys.stdout
        self.f = open(p, "w", encoding="utf-8")
    def write(self, m):
        try: self.t.write(m)
        except: pass
        self.f.write(m); self.f.flush()
    def flush(self):
        self.t.flush(); self.f.flush()
sys.stdout = DL(LOG)
sys.stderr = sys.stdout

try:
    from pyaedt import desktop as _dm
    _o = _dm.Desktop.__init__
    def _p(s, *a, **k):
        _o(s, *a, **k)
        if not hasattr(s, 'student_version'): s.student_version = False
    _dm.Desktop.__init__ = _p
except: pass
try:
    import pyaedt.application.Design as _dd
    _o2 = _dd.DesignSettings.__init__
    def _p2(s, app):
        try: _o2(s, app)
        except AttributeError:
            s._app = app; s.design_settings = None; s.manipulate_inputs = None
    _dd.DesignSettings.__init__ = _p2
except: pass

from pyaedt import Hfss
import traceback

# DrivenModal - no terminals needed!
hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "ModalWP_Test"),
    designname="Test1",
    solution_type="DrivenModal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)
hfss.modeler.model_units = "mm"
oDesign = hfss.odesign
oEditor = oDesign.SetActiveEditor("3D Modeler")
oDesktop = hfss.odesktop

def create_box(name, x, y, z, dx, dy, dz, mat, si=None):
    if si is None:
        si = mat.lower() not in ("copper","pec","aluminum")
    oEditor.CreateBox(
        ["NAME:BoxParameters",
         "XPosition:=", str(x), "YPosition:=", str(y), "ZPosition:=", str(z),
         "XSize:=", str(dx), "YSize:=", str(dy), "ZSize:=", str(dz)],
        ["NAME:Attributes", "Name:=", name, "Flags:=", "", "Color:=", "(143 175 131)",
         "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
         "MaterialValue:=", '"' + mat + '"', "SurfaceMaterialValue:=", '""',
         "SolveInside:=", si, "IsMaterialEditable:=", True,
         "UseMaterialAppearance:=", False, "IsLightweight:=", False])

print("Creating simple microstrip line geometry...")
# Simple microstrip line for testing: substrate + ground + trace
sub_L, sub_W, sub_H = 30.0, 20.0, 1.6
cu_t = 0.035
trace_w = 3.0  # microstrip width for ~50 Ohm on FR4

# Substrate
create_box("Substrate", -sub_L/2, -sub_W/2, 0, sub_L, sub_W, sub_H, "FR4_epoxy")
# Ground (bottom)
create_box("Ground", -sub_L/2, -sub_W/2, -cu_t, sub_L, sub_W, cu_t, "copper")
# Trace (top) - extends from y=-sub_W/2 to y=+sub_W/2 (full length)
create_box("Trace", -trace_w/2, -sub_W/2, sub_H, trace_w, sub_W, cu_t, "copper")
print("  Geometry OK")

# WavePort at y=-sub_W/2 (edge of substrate)
# Port spans wider than trace and taller than substrate
port_w = 5 * trace_w  # 15mm wide in x
port_h = 5 * sub_H    # 8mm tall in z
port_y = -sub_W / 2

oEditor.CreateRectangle(
    ["NAME:RectangleParameters",
     "IsCovered:=", True,
     "XStart:=", str(-port_w/2), "YStart:=", str(port_y), "ZStart:=", str(-port_h/5),
     "Width:=", str(port_w), "Height:=", str(port_h),
     "WhichAxis:=", "Y"],
    ["NAME:Attributes", "Name:=", "WPRect", "Flags:=", "", "Color:=", "(0 0 255)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '""', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  WPRect created")

oBnd = oDesign.GetModule("BoundarySetup")

# WavePort - DrivenModal uses Modes, no terminals
oBnd.AssignWavePort(
    ["NAME:WP1",
     "Objects:=", ["WPRect"],
     "NumModes:=", 1,
     "RenormalizeAllTerminals:=", True,
     "DoDeembed:=", False,
     ["NAME:Modes",
      ["NAME:Mode1",
       "ModeNum:=", 1,
       "UseIntLine:=", True,
       ["NAME:IntLine",
        "Start:=", ["0", str(port_y), str(sub_H + cu_t)],
        "End:=",   ["0", str(port_y), str(-cu_t)]],
       "CharImp:=", "Zpi"]]])
print("  WavePort WP1 assigned")

# AirBox - bottom face at y=-sub_W/2 (where port is)
pad = 10.0
create_box("AirBox",
           -(sub_L/2 + pad), -sub_W/2, -(pad),
           sub_L + 2*pad, sub_W + pad, sub_H + 2*pad,
           "vacuum", si=True)

oBnd.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["AirBox"],
     "IsFssReference:=", False,
     "IsForPML:=", False])
print("  Radiation boundary on AirBox")

# Analysis setup
oAnalysis = oDesign.GetModule("AnalysisSetup")
oAnalysis.InsertSetup("HfssDriven",
    ["NAME:Setup1",
     "Frequency:=", "2.217GHz",
     "MaxDeltaS:=", 0.05,
     "MaximumPasses:=", 6,
     "MinimumPasses:=", 2,
     "MinimumConvergedPasses:=", 1,
     "PercentRefinement:=", 30,
     "IsEnabled:=", True,
     "BasisOrder:=", 1,
     "UseIterativeSolver:=", False,
     "DoLambdaRefine:=", True,
     "DoMaterialLambdaRefine:=", True,
     "SetLambdaTarget:=", False,
     "Target:=", 0.3333])
print("  Setup1 created")

hfss.save_project()
print("  Saved")

# Validate
print("\n--- Validation ---")
try:
    v = oDesign.ValidateDesign()
    print(f"  Result: {v}")
except Exception as e:
    print(f"  Error: {e}")

try:
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        print(f"  MSG: {m}")
except: pass

if True:  # always try
    print("\n--- Analyze ---")
    t0 = time.time()
    try:
        oDesign.Analyze("Setup1")
        print(f"  Analyze COMPLETED in {time.time()-t0:.1f}s!")
    except Exception as e:
        print(f"  Analyze FAILED ({time.time()-t0:.1f}s): {e}")
        traceback.print_exc()

    try:
        msgs = oDesktop.GetMessages("", "", 2)
        for m in msgs:
            print(f"  MSG: {m}")
    except: pass

print("\nDone.")
try: hfss.release_desktop()
except: pass
