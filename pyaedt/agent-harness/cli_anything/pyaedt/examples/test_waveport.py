"""Test: Complete workflow with WavePort - validate and analyze."""
import os, sys
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "waveport_test.log")

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

# Simple microstrip test with WavePort at substrate edge
# The feed line extends to the substrate/airbox edge where the WavePort is placed
hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "WP_Test"),
    designname="WPTest",
    solution_type="DrivenTerminal",
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

print("Creating geometry...")
sub_L, sub_W, sub_H = 60.0, 40.0, 1.6
cu_t = 0.035
pad = 20.0  # air padding

# Substrate
create_box("Substrate", -sub_L/2, -sub_W/2, 0, sub_L, sub_W, sub_H, "FR4_epoxy")

# Ground plane (bottom of substrate)
create_box("Ground", -sub_L/2, -sub_W/2, -cu_t, sub_L, sub_W, cu_t, "copper")

# A simple dipole-like trace on top
trace_w = 2.0
trace_l = 25.0
create_box("Trace", -trace_l, -trace_w/2, sub_H, 2*trace_l, trace_w, cu_t, "copper")

# Feed line from center to substrate bottom edge
feed_w = 1.5
feed_y_end = -sub_W/2  # ends at substrate edge
create_box("FeedLine", -feed_w/2, feed_y_end, sub_H, feed_w, sub_W/2, cu_t, "copper")
print("  Geometry OK")

# WavePort at substrate bottom edge (y = -sub_W/2)
# The port rectangle spans from ground bottom to above trace top in the xz plane
# Port width should be big enough (5x trace width typical for waveport)
port_w = 6 * feed_w  # 9mm width in x
port_h = sub_H + 6 * sub_H  # height in z
# Center the port around the feed
px_start = -port_w / 2
pz_start = -port_h / 6  # extend slightly below ground

print(f"\nCreating WavePort at y={feed_y_end}...")
print(f"  Port rect: x=[{px_start}, {px_start+port_w}], z=[{pz_start}, {pz_start+port_h}]")

oEditor.CreateRectangle(
    ["NAME:RectangleParameters",
     "IsCovered:=", True,
     "XStart:=", str(px_start), "YStart:=", str(feed_y_end), "ZStart:=", str(pz_start),
     "Width:=", str(port_w), "Height:=", str(port_h),
     "WhichAxis:=", "Y"],
    ["NAME:Attributes", "Name:=", "WPRect", "Flags:=", "", "Color:=", "(0 0 255)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '""', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  WPRect created")

oBnd = oDesign.GetModule("BoundarySetup")

# Assign WavePort
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
        "Start:=", ["0", str(feed_y_end), str(sub_H + cu_t)],
        "End:=",   ["0", str(feed_y_end), str(-cu_t)]],
       "CharImp:=", "Zpi"]]])
print("  WavePort WP1 assigned")

# AirBox - make the bottom face (y=-sub_W/2-pad) coincide with port
air_x = sub_L/2 + pad
air_y_neg = sub_W/2 + pad  # but make bottom at -sub_W/2 (matching port)
air_y_pos = sub_W/2 + pad
air_z_top = pad
air_z_bot = pad

# Actually, the WavePort must be ON the airbox face. So airbox y_min = -sub_W/2
# But then there's no padding on the feed side...
# The standard approach: airbox y_min = feed_y_end = -sub_W/2
# WavePort is on this face
create_box("AirBox",
           -air_x, -sub_W/2, -air_z_bot,
           2*air_x, sub_W + pad, sub_H + air_z_top + air_z_bot,
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
     "MaxDeltaS:=", 0.02,
     "MaximumPasses:=", 10,
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

# Save
hfss.save_project()
print("\n  Project saved")

# Validate
print("\n--- Validation ---")
try:
    v = oDesign.ValidateDesign()
    print(f"  Result: {v}")
except Exception as e:
    print(f"  Error: {e}")

# Check messages
try:
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        print(f"  MSG: {m}")
except: pass

# Try Analyze
print("\n--- Analyze ---")
try:
    oDesign.Analyze("Setup1")
    print("  Analyze COMPLETED!")
except Exception as e:
    print(f"  Analyze FAILED: {e}")
    traceback.print_exc()

# Check messages after analyze
try:
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        if "error" in str(m).lower() or "warning" in str(m).lower():
            print(f"  MSG: {m}")
except: pass

print("\nDone.")
try: hfss.release_desktop()
except: pass
