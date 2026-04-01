"""ALL-IN-ONE: Geometry + Port + Setup + Sweep + Simulate + Extract.
Key discoveries:
- AutoIdentifyPorts on airbox y_min face (index 2) creates WavePort with terminal
- Simulation WORKS in DrivenTerminal mode
- Need to find correct data extraction method for AEDT 2019.1
"""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "allinone_test.log")

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
    projectname=os.path.join(PROJECT_DIR, "AllInOne"),
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

# ========== STEP 1: GEOMETRY ==========
print("=== STEP 1: Geometry ===")
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

# ========== STEP 2: BOUNDARIES & PORT ==========
print("\n=== STEP 2: Port & Boundaries ===")
oBnd.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["Air"],
     "IsFssReference:=", False,
     "IsForPML:=", False])

air_faces = oEditor.GetFaceIDs("Air")
ymin_face = int(air_faces[2])
oBnd.AutoIdentifyPorts(
    ["NAME:Faces", ymin_face],
    True, ["NAME:ReferenceConductors", "Gnd"], "Port1", True)

terms = oBnd.GetExcitationsOfType("Terminal")
print(f"  Terminals: {terms}")
terminal_name = str(terms[0]) if terms else "Trace_T1"
print(f"  Port OK, terminal: {terminal_name}")

# ========== STEP 3: SETUP & SWEEP ==========
print("\n=== STEP 3: Setup & Sweep ===")
oAnalysis = oDesign.GetModule("AnalysisSetup")
oAnalysis.InsertSetup("HfssDriven",
    ["NAME:Setup1",
     "Frequency:=", "2.217GHz",
     "MaxDeltaS:=", 0.02,
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

# Frequency sweep 1.5 - 3.0 GHz, step 0.01
oAnalysis.InsertFrequencySweep("Setup1",
    ["NAME:Sweep1",
     "IsEnabled:=", True,
     "SetupType:=", "LinearCount",
     "StartValue:=", "1.5GHz",
     "StopValue:=", "3GHz",
     "Count:=", 151,
     "Type:=", "Discrete",
     "SaveFields:=", False,
     "ExtrapToDC:=", False])
print("  Setup1 + Sweep1 OK")

hfss.save_project()

# ========== STEP 4: VALIDATE & SIMULATE ==========
print("\n=== STEP 4: Validate & Simulate ===")
v = oDesign.ValidateDesign()
print(f"  Validation: {v}")
if not v:
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        print(f"  ERROR: {m}")
    print("  VALIDATION FAILED - aborting")
    sys.exit(1)

t0 = time.time()
oDesign.Analyze("Setup1")
elapsed = time.time() - t0
print(f"  *** Simulation COMPLETED in {elapsed:.1f}s ***")

# ========== STEP 5: EXTRACT DATA ==========
print("\n=== STEP 5: Extract Data ===")
oReport = oDesign.GetModule("ReportSetup")
oSolutions = oDesign.GetModule("Solutions")

# Method A: ExportNetworkData for touchstone file
print("\n--- Method A: ExportNetworkData (touchstone) ---")
s1p_path = os.path.join(PROJECT_DIR, "allinone.s1p")
for fmt_args in [
    # Format 1: 10 args
    ("", ["Setup1:Sweep1"], True, 50, "S", -1, -1, "s1p", s1p_path, ["All"]),
    # Format 2: setup name format variants
    ("", ["Setup1 : Sweep1"], True, 50, "S", -1, -1, "s1p", s1p_path, ["All"]),
    # Format 3: last adaptive
    ("", ["Setup1:LastAdaptive"], True, 50, "S", -1, -1, "s1p", s1p_path, ["All"]),
]:
    try:
        oDesign.ExportNetworkData(*fmt_args)
        if os.path.exists(s1p_path):
            print(f"  ExportNetworkData OK: {s1p_path}")
            with open(s1p_path, 'r') as f:
                lines = f.readlines()
            print(f"  Lines: {len(lines)}")
            for line in lines[:10]:
                print(f"    {line.rstrip()}")
            break
    except Exception as e:
        print(f"  ExportNetworkData failed: {e}")

# Method B: Try different CreateReport formats
print("\n--- Method B: CreateReport ---")
report_ok = False
for i, args in enumerate([
    # 8-arg format
    ("S11", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep1",
     [], ["Domain:=", "Sweep"], ["Freq:=", ["All"]],
     ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal_name},{terminal_name}))"]]),
    # 8-arg with NAME context
    ("S11", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep1",
     ["NAME:Context"], ["Domain:=", "Sweep"], ["Freq:=", ["All"]],
     ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal_name},{terminal_name}))"]]),
    # 7-arg
    ("S11", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep1",
     ["Domain:=", "Sweep"], ["Freq:=", ["All"]],
     ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal_name},{terminal_name}))"]]),
    # 7-arg LastAdaptive
    ("S11", "Terminal Solution Data", "Rectangular Plot", "Setup1 : LastAdaptive",
     ["Domain:=", "Sweep"], ["Freq:=", ["All"]],
     ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal_name},{terminal_name}))"]]),
]):
    try:
        oReport.CreateReport(*args)
        print(f"  Format {i}: OK!")
        report_ok = True
        break
    except Exception as e:
        print(f"  Format {i}: {e}")

if report_ok:
    csv_path = os.path.join(PROJECT_DIR, "allinone_s11.csv")
    try:
        oReport.ExportToFile("S11", csv_path)
        print(f"  Exported to {csv_path}")
        with open(csv_path, 'r') as f:
            content = f.read()
        print(f"  Content: {content[:500]}")
    except Exception as e:
        print(f"  Export failed: {e}")

# Method C: ExportSolution
print("\n--- Method C: oSolutions methods ---")
for method in dir(oSolutions):
    if 'export' in method.lower() or 'get' in method.lower() or 'solution' in method.lower():
        print(f"  Method: {method}")

# Try ListValuesOfVariable
print("\n--- Method D: GetValidISolutionList ---")
try:
    sols = oDesign.GetValidISolutionList()
    print(f"  Solutions: {sols}")
except Exception as e:
    print(f"  GetValidISolutionList: {e}")

# Try direct export
print("\n--- Method E: oSolutions.ExportNetworkData ---")
try:
    oSolutions.ExportNetworkData(
        "", "Setup1:Sweep1", s1p_path, "Terminal", 50, True)
    print(f"  OK!")
except Exception as e:
    print(f"  Failed: {e}")

try:
    oSolutions.ExportNetworkData(
        [], "Setup1:Sweep1", s1p_path, True, 50, "S", -1, -1, "s1p")
    print(f"  Alt format OK!")
except Exception as e:
    print(f"  Alt format: {e}")

print("\nDone.")
try: hfss.release_desktop()
except: pass
