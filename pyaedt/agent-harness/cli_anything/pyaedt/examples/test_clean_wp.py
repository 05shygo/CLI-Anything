"""CLEAN: Single WavePort on airbox y_min face where conductors exist.
From the airbox_face test, face index 2 (the 3rd face) is y_min where
both Gnd and Trace start. AutoIdentifyPorts on that face creates
a valid terminal."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "clean_wp_test.log")

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
    projectname=os.path.join(PROJECT_DIR, "CleanWP"),
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

AIR_YMIN = -15.0

print("=== Creating geometry ===")
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", f"{AIR_YMIN}mm", "ZPosition:=", "0mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "1.6mm"],
    ["NAME:Attributes", "Name:=", "Sub", "Flags:=", "", "Color:=", "(0 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"FR4_epoxy"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", f"{AIR_YMIN}mm", "ZPosition:=", "-0.035mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Gnd", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-1.5mm", "YPosition:=", f"{AIR_YMIN}mm", "ZPosition:=", "1.6mm",
     "XSize:=", "3mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Trace", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-15mm", "YPosition:=", f"{AIR_YMIN}mm", "ZPosition:=", "-10mm",
     "XSize:=", "30mm", "YSize:=", "30mm", "ZSize:=", "25mm"],
    ["NAME:Attributes", "Name:=", "Air", "Flags:=", "", "Color:=", "(143 175 131)",
     "Transparency:=", 0.8, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"vacuum"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  Geometry OK")

oBnd.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["Air"],
     "IsFssReference:=", False,
     "IsForPML:=", False])
print("  Radiation boundary OK")

# Get Air face IDs and find the one with conductors
air_faces = oEditor.GetFaceIDs("Air")
print(f"  Air faces: {air_faces}")

# Try each face until we find one with conductors (terminal created)
port_created = False
for i, fid_str in enumerate(air_faces):
    fid = int(fid_str)
    oBnd.AutoIdentifyPorts(
        ["NAME:Faces", fid],
        True,  # WavePort
        ["NAME:ReferenceConductors", "Gnd"],
        "Port1",
        True)
    
    terms = oBnd.GetExcitationsOfType("Terminal")
    if terms and len(terms) > 0:
        print(f"  ** Face {fid} (index {i}): Terminal found: {terms} **")
        port_created = True
        break
    else:
        # No conductor on this face, delete the empty port
        try:
            oBnd.DeleteBoundary(["Port1"])
            print(f"  Face {fid} (index {i}): no conductors, deleted")
        except:
            print(f"  Face {fid} (index {i}): no conductors, delete failed")

if not port_created:
    print("  NO VALID PORT FOUND!")
    sys.exit(1)

# Check state
exc = oBnd.GetExcitations()
print(f"  Excitations: {exc}")
terms = oBnd.GetExcitationsOfType("Terminal")
print(f"  Terminals: {terms}")
wports = oBnd.GetExcitationsOfType("Wave Port")
print(f"  Wave Ports: {wports}")

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

# Validate
print("\n=== VALIDATION ===")
v = oDesign.ValidateDesign()
print(f"  Result: {v}")

msgs = oDesktop.GetMessages("", "", 2)
for m in msgs:
    s = str(m).lower()
    if any(k in s for k in ["terminal", "port", "error", "non-conduct", "pass"]):
        print(f"  MSG: {m}")

# ANALYZE
print("\n=== ANALYZE ===")
t0 = time.time()
try:
    oDesign.Analyze("Setup1")
    elapsed = time.time() - t0
    print(f"  *** SIMULATION SUCCEEDED in {elapsed:.1f}s! ***")
    
    # Extract S-params
    terminal_name = str(terms[0])
    oReport = oDesign.GetModule("ReportSetup")
    oReport.CreateReport(
        "S11", "Terminal Solution Data", "Rectangular Plot",
        "Setup1 : LastAdaptive",
        ["Domain:=", "Sweep"],
        ["Freq:=", ["All"]],
        ["X Component:=", "Freq",
         "Y Component:=", [f"dB(St({terminal_name},{terminal_name}))"]])
    s11_csv = os.path.join(PROJECT_DIR, "clean_wp_s11.csv")
    oReport.ExportToFile("S11", s11_csv)
    print(f"  S11 exported: {s11_csv}")
    
except Exception as e:
    elapsed = time.time() - t0
    print(f"  FAILED ({elapsed:.1f}s): {e}")
    
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        s = str(m).lower()
        if any(k in s for k in ["error", "terminal", "non-conduct", "process", "pass"]):
            print(f"  POST-MSG: {m}")

print("\nDone.")
try: hfss.release_desktop()
except: pass
