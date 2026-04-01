"""CLEAN TEST: Single AutoIdentifyPorts on PRect face -> Lumped Port with terminal.
This IS the correct method for DrivenTerminal in AEDT 2019.1!"""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "single_autoid_test.log")

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
    projectname=os.path.join(PROJECT_DIR, "SingleAutoID"),
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

print("=== Creating geometry ===")
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

# Gnd
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

# Port sheet at y=-3 (cuts through both conductors)
pz_start = -0.035
pz_end = 1.635
port_h = pz_end - pz_start
oEditor.CreateRectangle(
    ["NAME:RectangleParameters",
     "IsCovered:=", True,
     "XStart:=", "-1.5mm", "YStart:=", "-3mm", "ZStart:=", f"{pz_start}mm",
     "Width:=", "3mm", "Height:=", f"{port_h}mm",
     "WhichAxis:=", "Y"],
    ["NAME:Attributes", "Name:=", "PRect", "Flags:=", "", "Color:=", "(0 0 255)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '""', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

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
print("  All geometry OK")

# Radiation boundary
oBnd.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["Air"],
     "IsFssReference:=", False,
     "IsForPML:=", False])
print("  Radiation boundary OK")

# Get PRect face ID
prect_faces = oEditor.GetFaceIDs("PRect")
print(f"  PRect face IDs: {prect_faces}")
face_id = int(prect_faces[0])

# SINGLE AutoIdentifyPorts call - Lumped Port
print(f"\n=== AutoIdentifyPorts on PRect face {face_id} (Lumped) ===")
oBnd.AutoIdentifyPorts(
    ["NAME:Faces", face_id],
    False,  # IsWavePort=False -> Lumped Port
    ["NAME:ReferenceConductors", "Gnd"],
    "Port1",
    True)
print("  AutoIdentifyPorts OK!")

# Check
exc = oBnd.GetExcitations()
print(f"  Excitations: {exc}")
terms = oBnd.GetExcitationsOfType("Terminal")
print(f"  Terminals: {terms}")
ports = oBnd.GetExcitationsOfType("Lumped Port")
print(f"  Lumped Ports: {ports}")

# Analysis setup
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
print("  Setup1 OK")

hfss.save_project()
print("  Saved")

# Validate
print("\n=== VALIDATION ===")
v = oDesign.ValidateDesign()
print(f"  Result: {v}")

msgs = oDesktop.GetMessages("", "", 2)
for m in msgs:
    print(f"  MSG: {m}")

# Analyze
if v:
    print("\n=== ANALYZE ===")
    t0 = time.time()
    try:
        oDesign.Analyze("Setup1")
        elapsed = time.time() - t0
        print(f"  *** SIMULATION SUCCEEDED in {elapsed:.1f}s! ***")
        
        # Try S-param extraction
        oReport = oDesign.GetModule("ReportSetup")
        # In DrivenTerminal, terminal names are like PRect_T1
        # S-param ref: St(PRect_T1,PRect_T1)  
        terminal_name = str(terms[0]) if terms else "PRect_T1"
        print(f"  Using terminal: {terminal_name}")
        
        oReport.CreateReport(
            "S11", "Terminal Solution Data", "Rectangular Plot",
            "Setup1 : LastAdaptive",
            ["Domain:=", "Sweep"],
            ["Freq:=", ["All"]],
            ["X Component:=", "Freq",
             "Y Component:=", [f"dB(St({terminal_name},{terminal_name}))"]])
        s11_csv = os.path.join(PROJECT_DIR, "single_autoid_s11.csv")
        oReport.ExportToFile("S11", s11_csv)
        print(f"  S11 exported: {s11_csv}")
        
    except Exception as e:
        elapsed = time.time() - t0
        print(f"  FAILED ({elapsed:.1f}s): {e}")
else:
    print("  Validation failed, checking batch log...")
    # Still try analyze to get error details
    try:
        oDesign.Analyze("Setup1")
    except Exception as e:
        print(f"  Analyze error: {e}")

print("\n=== Batch log check ===")
msgs = oDesktop.GetMessages("", "", 2)
for m in msgs:
    s = str(m).lower()
    if any(k in s for k in ["terminal", "port", "error", "pass", "converge"]):
        print(f"  {m}")

print("\nDone.")
try: hfss.release_desktop()
except: pass
