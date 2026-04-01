"""Minimal test: Check if AEDT 2019.1 can solve at ALL with a simple waveguide."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "minimal_test.log")

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

# ========== Test 1: DrivenTerminal with explicit unit suffixes ==========
print("=" * 60)
print("TEST: DrivenTerminal + LumpedPort with mm suffixes")
print("=" * 60)

hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "MinTest"),
    designname="LPortTest",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)
hfss.modeler.model_units = "mm"
oDesign = hfss.odesign
oEditor = oDesign.SetActiveEditor("3D Modeler")
oDesktop = hfss.odesktop

# Create a SIMPLE coax-like setup with explicit mm units
# AirBox with PEC shell, inner conductor, lumped port between them

# Outer conductor (PEC box)
print("Creating geometry with mm suffixes...")
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-10mm", "YPosition:=", "-10mm", "ZPosition:=", "0mm",
     "XSize:=", "20mm", "YSize:=", "20mm", "ZSize:=", "30mm"],
    ["NAME:Attributes", "Name:=", "AirBox", "Flags:=", "", "Color:=", "(143 175 131)",
     "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"vacuum"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  AirBox OK")

# Simple substrate
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", "-5mm", "ZPosition:=", "10mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "1.6mm"],
    ["NAME:Attributes", "Name:=", "SubFR4", "Flags:=", "", "Color:=", "(0 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"FR4_epoxy"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  SubFR4 OK")

# Ground plane  
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", "-5mm", "ZPosition:=", "9.965mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "GndPlane", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  GndPlane OK")

# Trace on top of substrate
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-1.5mm", "YPosition:=", "-5mm", "ZPosition:=", "11.6mm",
     "XSize:=", "3mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "MsTrace", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  MsTrace OK")

# Port sheet at y=-5mm (substrate edge) in xz plane
# Spans from ground bottom to trace top
oEditor.CreateRectangle(
    ["NAME:RectangleParameters",
     "IsCovered:=", True,
     "XStart:=", "-1.5mm", "YStart:=", "-5mm", "ZStart:=", "9.965mm",
     "Width:=", "3mm", "Height:=", "1.67mm",
     "WhichAxis:=", "Y"],
    ["NAME:Attributes", "Name:=", "PortSheet", "Flags:=", "", "Color:=", "(0 0 255)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '""', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  PortSheet OK")

# List objects
try:
    print(f"\n  Solids: {oEditor.GetObjectsInGroup('Solids')}")
    print(f"  Sheets: {oEditor.GetObjectsInGroup('Sheets')}")
except: pass

oBnd = oDesign.GetModule("BoundarySetup")

# Assign lumped port
print("\nAssigning LumpedPort...")
try:
    oBnd.AssignLumpedPort(
        ["NAME:Port1",
         "Objects:=", ["PortSheet"],
         "RenormalizeAllTerminals:=", True,
         "DoDeembed:=", False,
         ["NAME:Modes",
          ["NAME:Mode1",
           "ModeNum:=", 1,
           "UseIntLine:=", True,
           ["NAME:IntLine",
            "Start:=", ["0mm", "-5mm", "11.635mm"],
            "End:=",   ["0mm", "-5mm", "9.965mm"]],
           "CharImp:=", "Zpi"]]])
    print("  Port1 assigned!")
except Exception as e:
    print(f"  Port1 FAILED: {e}")

# Check excitations and terminals
try:
    exc = oBnd.GetExcitations()
    print(f"  Excitations: {exc}")
except: pass

# Radiation boundary
oBnd.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["AirBox"],
     "IsFssReference:=", False,
     "IsForPML:=", False])
print("  Radiation boundary OK")

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
print("  Setup1 OK")

hfss.save_project()
print("  Saved")

# Validate
print("\n--- Validation ---")
v = oDesign.ValidateDesign()
print(f"  Result: {v}")

try:
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        print(f"  MSG: {m}")
except: pass

# Only analyze if validation passed
if v != 0:
    print("\n--- Analyze ---")
    t0 = time.time()
    try:
        oDesign.Analyze("Setup1")
        print(f"  Analyze COMPLETED in {time.time()-t0:.1f}s!")
    except Exception as e:
        print(f"  Analyze FAILED ({time.time()-t0:.1f}s): {e}")
else:
    print("\n  Validation FAILED, checking if we can analyze anyway...")
    # Check messages
    try:
        msgs = oDesktop.GetMessages("", "", 2)
        for m in msgs:
            print(f"  MSG: {m}")
    except: pass
    
    # Try analyze anyway
    print("\n--- Analyze (forced) ---")
    t0 = time.time()
    try:
        oDesign.Analyze("Setup1")
        print(f"  Analyze COMPLETED in {time.time()-t0:.1f}s!")
    except Exception as e:
        print(f"  Analyze FAILED ({time.time()-t0:.1f}s): {e}")

print("\n--- Final Messages ---")
try:
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        print(f"  {m}")
except: pass

print("\nDone.")
try: hfss.release_desktop()
except: pass
