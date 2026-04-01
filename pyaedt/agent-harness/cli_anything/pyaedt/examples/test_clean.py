"""Clean test: Subtract-based port topology + LumpedPort only, no extra ports."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "clean_test.log")

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

hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "CleanTest"),
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

print("Creating geometry...")
# Sub
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", "-5mm", "ZPosition:=", "0mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "1.6mm"],
    ["NAME:Attributes", "Name:=", "Sub", "Flags:=", "", "Color:=", "(0 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"FR4_epoxy"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Ground
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", "-5mm", "ZPosition:=", "-0.035mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Gnd", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Trace
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-1.5mm", "YPosition:=", "-5mm", "ZPosition:=", "1.6mm",
     "XSize:=", "3mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Trace", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

print("  Solids OK")

# Port rect - extends slightly into conductor volumes to ensure topological connection
oEditor.CreateRectangle(
    ["NAME:RectangleParameters",
     "IsCovered:=", True,
     "XStart:=", "-1.5mm", "YStart:=", "-5mm", "ZStart:=", "-0.04mm",
     "Width:=", "3mm", "Height:=", "1.68mm",
     "WhichAxis:=", "Y"],
    ["NAME:Attributes", "Name:=", "PRect", "Flags:=", "", "Color:=", "(0 0 255)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '""', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  PRect OK (extends 0.005mm into conductors)")

# Subtract PRect from both conductors (keeps PRect intact)
print("  Subtracting PRect from conductors...")
oEditor.Subtract(
    ["NAME:Selections",
     "Blank Parts:=", "Gnd", "Tool Parts:=", "PRect"],
    ["NAME:SubtractParameters",
     "KeepOriginals:=", True])
print("    Subtracted from Gnd")

oEditor.Subtract(
    ["NAME:Selections",
     "Blank Parts:=", "Trace", "Tool Parts:=", "PRect"],
    ["NAME:SubtractParameters",
     "KeepOriginals:=", True])
print("    Subtracted from Trace")

# AirBox
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-15mm", "YPosition:=", "-15mm", "ZPosition:=", "-10mm",
     "XSize:=", "30mm", "YSize:=", "30mm", "ZSize:=", "25mm"],
    ["NAME:Attributes", "Name:=", "Air", "Flags:=", "", "Color:=", "(143 175 131)",
     "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"vacuum"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  AirBox OK")

oBnd = oDesign.GetModule("BoundarySetup")

# Assign Lumped Port
print("\nAssigning LumpedPort...")
oBnd.AssignLumpedPort(
    ["NAME:Port1",
     "Objects:=", ["PRect"],
     "RenormalizeAllTerminals:=", True,
     "DoDeembed:=", False,
     ["NAME:Modes",
      ["NAME:Mode1",
       "ModeNum:=", 1,
       "UseIntLine:=", True,
       ["NAME:IntLine",
        "Start:=", ["0mm", "-5mm", "1.635mm"],
        "End:=",   ["0mm", "-5mm", "-0.035mm"]],
       "CharImp:=", "Zpi"]]])
print("  Port1 assigned!")

# Check excitations
try:
    exc = oBnd.GetExcitations()
    print(f"  Excitations: {exc}")
except: pass

# Radiation boundary
oBnd.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["Air"],
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
print("  Project saved")

# Validate
print("\n--- VALIDATION ---")
v = oDesign.ValidateDesign()
print(f"  Result: {v}")

msgs = oDesktop.GetMessages("", "", 2)
for m in msgs:
    print(f"  MSG: {m}")

# Analyze
print("\n--- ANALYZE ---")
t0 = time.time()
try:
    oDesign.Analyze("Setup1")
    elapsed = time.time() - t0
    print(f"  *** COMPLETED in {elapsed:.1f}s! ***")
except Exception as e:
    elapsed = time.time() - t0
    print(f"  FAILED ({elapsed:.1f}s): {e}")

msgs = oDesktop.GetMessages("", "", 2)
for m in msgs:
    if "error" in str(m).lower() or "complet" in str(m).lower() or "pass" in str(m).lower():
        print(f"  MSG: {m}")

print("\nDone.")
try: hfss.release_desktop()
except: pass
