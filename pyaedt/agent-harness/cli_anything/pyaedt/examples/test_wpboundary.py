"""Test: WavePort at airbox boundary - standard microstrip simulation approach.
The feed line extends to the airbox edge. WavePort is defined on that face."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "waveport_boundary_test.log")

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

hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "WPBoundary"),
    designname="T1",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)
hfss.modeler.model_units = "mm"
oDesign = hfss.odesign
oEditor = oDesign.SetActiveEditor("3D Modeler")
oDesktop = hfss.odesktop
oBnd = oDesign.GetModule("BoundarySetup")

# Geometry: simple microstrip line with feed extending to airbox boundary
# AirBox: y from -15 to 15, x from -15 to 15, z from -10 to 15
# Sub, Gnd, Trace all start at y=-15 (airbox boundary)
AIR_Y_MIN = -15.0
AIR_Y_MAX = 15.0
AIR_X_MIN = -15.0
AIR_X_MAX = 15.0
AIR_Z_MIN = -10.0
AIR_Z_MAX = 15.0

SUB_W = 10.0  # substrate x-dimension
SUB_L = 10.0  # substrate y-dimension
SUB_H = 1.6
TRACE_W = 3.0
CU_T = 0.035

print("=== Creating geometry ===")
# Sub: 10x10x1.6, centered, starting at y=AIR_Y_MIN
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", f"{-SUB_W/2}mm", "YPosition:=", f"{AIR_Y_MIN}mm", "ZPosition:=", "0mm",
     "XSize:=", f"{SUB_W}mm", "YSize:=", f"{SUB_L}mm", "ZSize:=", f"{SUB_H}mm"],
    ["NAME:Attributes", "Name:=", "Sub", "Flags:=", "", "Color:=", "(0 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"FR4_epoxy"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Gnd: extends from y=AIR_Y_MIN (airbox edge)
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", f"{-SUB_W/2}mm", "YPosition:=", f"{AIR_Y_MIN}mm", "ZPosition:=", f"{-CU_T}mm",
     "XSize:=", f"{SUB_W}mm", "YSize:=", f"{SUB_L}mm", "ZSize:=", f"{CU_T}mm"],
    ["NAME:Attributes", "Name:=", "Gnd", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Trace: extends from y=AIR_Y_MIN (airbox edge)
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", f"{-TRACE_W/2}mm", "YPosition:=", f"{AIR_Y_MIN}mm", "ZPosition:=", f"{SUB_H}mm",
     "XSize:=", f"{TRACE_W}mm", "YSize:=", f"{SUB_L}mm", "ZSize:=", f"{CU_T}mm"],
    ["NAME:Attributes", "Name:=", "Trace", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# WavePort sheet at airbox boundary y=AIR_Y_MIN
# Standard: wider and taller than needed (6*trace_w wide, well above trace)
wp_w = 6 * TRACE_W  # 18mm wide
wp_z_start = -CU_T - 1.0  # 1mm below ground bottom
wp_z_end = SUB_H + CU_T + 3.0  # 3mm above trace top
wp_h = wp_z_end - wp_z_start
oEditor.CreateRectangle(
    ["NAME:RectangleParameters",
     "IsCovered:=", True,
     "XStart:=", f"{-wp_w/2}mm", "YStart:=", f"{AIR_Y_MIN}mm", "ZStart:=", f"{wp_z_start}mm",
     "Width:=", f"{wp_w}mm", "Height:=", f"{wp_h}mm",
     "WhichAxis:=", "Y"],
    ["NAME:Attributes", "Name:=", "WPRect", "Flags:=", "", "Color:=", "(255 0 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '""', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# AirBox
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", f"{AIR_X_MIN}mm", "YPosition:=", f"{AIR_Y_MIN}mm", "ZPosition:=", f"{AIR_Z_MIN}mm",
     "XSize:=", f"{AIR_X_MAX-AIR_X_MIN}mm", "YSize:=", f"{AIR_Y_MAX-AIR_Y_MIN}mm",
     "ZSize:=", f"{AIR_Z_MAX-AIR_Z_MIN}mm"],
    ["NAME:Attributes", "Name:=", "Air", "Flags:=", "", "Color:=", "(143 175 131)",
     "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"vacuum"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  All geometry OK")

# Radiation boundary - on ALL faces EXCEPT the WavePort face
oBnd.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["Air"],
     "IsFssReference:=", False,
     "IsForPML:=", False])
print("  Radiation boundary OK")

# Get WPRect face ID
wp_faces = oEditor.GetFaceIDs("WPRect")
print(f"  WPRect face IDs: {wp_faces}")
face_id = int(wp_faces[0])

# WavePort via AutoIdentifyPorts
print(f"\n=== AutoIdentifyPorts on WPRect face {face_id} (WavePort) ===")
try:
    oBnd.AutoIdentifyPorts(
        ["NAME:Faces", face_id],
        True,  # IsWavePort=True
        ["NAME:ReferenceConductors", "Gnd"],
        "Port1",
        True)
    print("  AutoIdentifyPorts OK!")
except Exception as e:
    print(f"  AutoIdentifyPorts FAILED: {e}")

# Check excitations
exc = oBnd.GetExcitations()
print(f"  Excitations: {exc}")
try:
    terms = oBnd.GetExcitationsOfType("Terminal")
    print(f"  Terminals: {terms}")
except: pass
try:
    wports = oBnd.GetExcitationsOfType("Wave Port")
    print(f"  Wave Ports: {wports}")
except: pass

# Also try: AutoIdentifyPorts as LumpedPort with PEC-bounded sheet
print(f"\n=== Method 2: LumpedPort with PEC Sidewalls ===")
# Create a port sheet with PEC sidewalls connecting Trace to Gnd
# This creates a "closed" port cross-section
lp_y = -3.0  # port location (inside the structure)
lp_w = TRACE_W  # 3mm, same as trace
# Create two PEC walls on sides of port
# Left wall: x=-1.5, z from 0 to 1.6, y at lp_y, thickness 0.1mm in y
# Actually in AEDT, PEC boundary on a face is simpler
# Let me try just using integration line with AutoIdentify on a narrower port

# Create a narrow port sheet (1mm wide, centered) between trace and ground
oEditor.CreateRectangle(
    ["NAME:RectangleParameters",
     "IsCovered:=", True,
     "XStart:=", "-0.5mm", "YStart:=", f"{lp_y}mm", "ZStart:=", "0mm",
     "Width:=", "1mm", "Height:=", f"{SUB_H}mm",
     "WhichAxis:=", "Y"],
    ["NAME:Attributes", "Name:=", "NarrowPort", "Flags:=", "", "Color:=", "(0 255 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '""', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

np_faces = oEditor.GetFaceIDs("NarrowPort")
print(f"  NarrowPort faces: {np_faces}")
np_fid = int(np_faces[0])

try:
    oBnd.AutoIdentifyPorts(
        ["NAME:Faces", np_fid],
        False,  # Lumped
        ["NAME:ReferenceConductors", "Gnd"],
        "Port2",
        True)
    print("  NarrowPort AutoIdentify OK!")
except Exception as e:
    print(f"  NarrowPort AutoIdentify FAILED: {e}")

exc2 = oBnd.GetExcitations()
print(f"  All Excitations now: {exc2}")

# Setup
oAnalysis = oDesign.GetModule("AnalysisSetup")
oAnalysis.InsertSetup("HfssDriven",
    ["NAME:Setup1",
     "Frequency:=", "2.217GHz",
     "MaxDeltaS:=", 0.05,
     "MaximumPasses:=", 3,
     "MinimumPasses:=", 1,
     "MinimumConvergedPasses:=", 1,
     "PercentRefinement:=", 30,
     "IsEnabled:=", True,
     "BasisOrder:=", 1,
     "UseIterativeSolver:=", False,
     "DoLambdaRefine:=", True,
     "DoMaterialLambdaRefine:=", True,
     "SetLambdaTarget:=", False,
     "Target:=", 0.3333])

hfss.save_project()

# Validate
print("\n=== VALIDATION ===")
v = oDesign.ValidateDesign()
print(f"  Result: {v}")

msgs = oDesktop.GetMessages("", "", 2)
for m in msgs:
    s = str(m)
    if any(k in s.lower() for k in ["terminal", "port", "error", "non-conducting"]):
        print(f"  MSG: {m}")

if v:
    print("\n=== ANALYZE ===")
    t0 = time.time()
    try:
        oDesign.Analyze("Setup1")
        elapsed = time.time() - t0
        print(f"  *** SUCCESS in {elapsed:.1f}s ***")
    except Exception as e:
        elapsed = time.time() - t0
        print(f"  FAILED ({elapsed:.1f}s): {e}")
        # Check messages after failure
        msgs = oDesktop.GetMessages("", "", 2)
        for m in msgs:
            s = str(m).lower()
            if any(k in s for k in ["terminal", "port", "error", "non-conducting", "process"]):
                print(f"  POST-MSG: {m}")

print("\nDone.")
try: hfss.release_desktop()
except: pass
