"""Test: Conductors extend beyond port sheet so it CUTS THROUGH them.
Key insight: AEDT terminal detection requires the port sheet to cut
through conductors, not just coincide with their end faces."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "cut_test.log")

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
    projectname=os.path.join(PROJECT_DIR, "CutTest"),
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

PORT_Y = -3.0  # port location in middle of the conductors

print("Creating geometry...")
# Sub: extends from y=-5 to y=5
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", "-5mm", "ZPosition:=", "0mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "1.6mm"],
    ["NAME:Attributes", "Name:=", "Sub", "Flags:=", "", "Color:=", "(0 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"FR4_epoxy"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Ground: extends from y=-5 to y=5 (port at y=-3 cuts through)
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", "-5mm", "ZPosition:=", "-0.035mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Gnd", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Trace: extends from y=-5 to y=5 (port at y=-3 cuts through)
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

# Port rect at y=PORT_Y, spanning from below Gnd to above Trace
# It CUTS THROUGH both conductors since they extend on both sides of PORT_Y
pz_start = -0.035  # bottom of Gnd
pz_end = 1.635     # top of Trace (1.6 + 0.035)
port_h = pz_end - pz_start  # 1.67 mm
oEditor.CreateRectangle(
    ["NAME:RectangleParameters",
     "IsCovered:=", True,
     "XStart:=", "-1.5mm", "YStart:=", f"{PORT_Y}mm", "ZStart:=", f"{pz_start}mm",
     "Width:=", "3mm", "Height:=", f"{port_h}mm",
     "WhichAxis:=", "Y"],
    ["NAME:Attributes", "Name:=", "PRect", "Flags:=", "", "Color:=", "(0 0 255)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '""', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print(f"  Port sheet at y={PORT_Y}mm (cuts through Gnd and Trace)")

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

# Assign LumpedPort
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
        "Start:=", [f"0mm", f"{PORT_Y}mm", "1.635mm"],
        "End:=",   [f"0mm", f"{PORT_Y}mm", "-0.035mm"]],
       "CharImp:=", "Zpi"]]])
print("  Port1 assigned!")

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
print("  Saved")

# Validate
print("\n--- VALIDATION ---")
v = oDesign.ValidateDesign()
print(f"  Result: {v} (0=fail, non-zero=pass)")

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
    
    # Check S-params
    oReport = oDesign.GetModule("ReportSetup")
    oReport.CreateReport(
        "S11", "Terminal Solution Data", "Rectangular Plot",
        "Setup1 : LastAdaptive",
        ["Domain:=", "Sweep"],
        ["Freq:=", ["All"]],
        ["X Component:=", "Freq",
         "Y Component:=", ["dB(St(Port1,Port1))"]])
    s11_csv = os.path.join(PROJECT_DIR, "cut_test_s11.csv")
    oReport.ExportToFile("S11", s11_csv)
    print(f"  S11 exported: {s11_csv}")
    
except Exception as e:
    elapsed = time.time() - t0
    print(f"  FAILED ({elapsed:.1f}s): {e}")

# Check batch log
print("\n--- Batch log messages ---")
try:
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        if "error" in str(m).lower() or "terminal" in str(m).lower() or "pass" in str(m).lower():
            print(f"  {m}")
except: pass

print("\nDone.")
try: hfss.release_desktop()
except: pass
