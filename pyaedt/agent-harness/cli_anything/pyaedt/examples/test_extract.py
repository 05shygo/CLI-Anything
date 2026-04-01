"""Test data extraction from a FRESH simulation.
Run geometry+port+setup+analysis+extraction all in one go."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "extract_test.log")

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
    projectname=os.path.join(PROJECT_DIR, "ExtractTest"),
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

# Simple geometry
print("=== Geometry ===")
for name, args, mat, solve in [
    ("Sub", ["-5mm",f"{AIR_YMIN}mm","0mm","10mm","10mm","1.6mm"], "FR4_epoxy", True),
    ("Gnd", ["-5mm",f"{AIR_YMIN}mm","-0.035mm","10mm","10mm","0.035mm"], "copper", False),
    ("Trace",["-1.5mm",f"{AIR_YMIN}mm","1.6mm","3mm","10mm","0.035mm"], "copper", False),
    ("Air", ["-15mm",f"{AIR_YMIN}mm","-10mm","30mm","30mm","25mm"], "vacuum", True),
]:
    oEditor.CreateBox(
        ["NAME:BoxParameters",
         "XPosition:=",args[0],"YPosition:=",args[1],"ZPosition:=",args[2],
         "XSize:=",args[3],"YSize:=",args[4],"ZSize:=",args[5]],
        ["NAME:Attributes","Name:=",name,"Flags:=","","Color:=","(128 128 128)",
         "Transparency:=",0,"PartCoordinateSystem:=","Global","UDMId:=","",
         "MaterialValue:=",f'"{mat}"',"SurfaceMaterialValue:=",'""',
         "SolveInside:=",solve,"IsMaterialEditable:=",True,
         "UseMaterialAppearance:=",False,"IsLightweight:=",False])

oBnd.AssignRadiation(["NAME:Rad1","Objects:=",["Air"],"IsFssReference:=",False,"IsForPML:=",False])

air_faces = oEditor.GetFaceIDs("Air")
oBnd.AutoIdentifyPorts(["NAME:Faces",int(air_faces[2])],True,["NAME:ReferenceConductors","Gnd"],"Port1",True)
terms = oBnd.GetExcitationsOfType("Terminal")
terminal = str(terms[0]) if terms else "Trace_T1"
print(f"  Terminal: {terminal}")

# Setup with 11-point sweep (fewer points for faster test)
oAnalysis = oDesign.GetModule("AnalysisSetup")
oAnalysis.InsertSetup("HfssDriven",
    ["NAME:Setup1","Frequency:=","2.217GHz","MaxDeltaS:=",0.05,
     "MaximumPasses:=",3,"MinimumPasses:=",1,"MinimumConvergedPasses:=",1,
     "PercentRefinement:=",30,"IsEnabled:=",True,"BasisOrder:=",1,
     "UseIterativeSolver:=",False,"DoLambdaRefine:=",True,
     "DoMaterialLambdaRefine:=",True,"SetLambdaTarget:=",False,"Target:=",0.3333])
oAnalysis.InsertFrequencySweep("Setup1",
    ["NAME:Sweep1","IsEnabled:=",True,"SetupType:=","LinearCount",
     "StartValue:=","1.5GHz","StopValue:=","3GHz","Count:=",16,
     "Type:=","Discrete","SaveFields:=",False,"ExtrapToDC:=",False])

hfss.save_project()

# Simulate
print("\n=== Simulate ===")
v = oDesign.ValidateDesign()
print(f"  Validation: {v}")
t0 = time.time()
oDesign.Analyze("Setup1")
print(f"  Done in {time.time()-t0:.1f}s")

# ========== DATA EXTRACTION ==========
print("\n=== Extract Data ===")
oSolutions = oDesign.GetModule("Solutions")
oReport = oDesign.GetModule("ReportSetup")
s1p_path = os.path.join(PROJECT_DIR, "extract_test.s1p")

# Method 1: oSolutions.ExportNetworkData with correct format
print("\n--- M1: oSolutions.ExportNetworkData ---")
formats_to_try = [
    # 7 args (minimal)
    ("", ["Setup1:Sweep1"], 3, s1p_path, ["All"], True, 50),
    # 8 args (with matrix type)
    ("", ["Setup1:Sweep1"], 3, s1p_path, ["All"], True, 50, "S"),
    # 9 args
    ("", ["Setup1:Sweep1"], 3, s1p_path, ["All"], True, 50, "S", -1),
    # 10 args
    ("", ["Setup1:Sweep1"], 3, s1p_path, ["All"], True, 50, "S", -1, 15),
    # 11 args
    ("", ["Setup1:Sweep1"], 3, s1p_path, ["All"], True, 50, "S", -1, 15, 8),
    # With "Last Adaptive" instead of Sweep1
    ("", ["Setup1:LastAdaptive"], 3, s1p_path, ["All"], True, 50),
    # 2 for tab format
    ("", ["Setup1:Sweep1"], 2, s1p_path.replace(".s1p",".tab"), ["All"], True, 50),
]

for i, args in enumerate(formats_to_try):
    try:
        oSolutions.ExportNetworkData(*args)
        outpath = args[3]
        if os.path.exists(outpath):
            print(f"  Format {i}: SUCCESS! File: {outpath}")
            with open(outpath, 'r') as f:
                lines = f.readlines()
            print(f"  Lines: {len(lines)}")
            for line in lines[:15]:
                print(f"    {line.rstrip()}")
            break
        else:
            print(f"  Format {i}: call OK but no file created")
    except Exception as e:
        print(f"  Format {i}: {e}")

# Method 2: CreateReport with 8 args (needed in AEDT 2019.1)
print("\n--- M2: CreateReport (8 args) ---")
cr_formats = [
    # Standard 8-arg format
    ("S11_A", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep1",
     [], ["Freq:=", ["All"]], 
     ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal},{terminal}))"]], []),
    # with Domain context
    ("S11_B", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep1",
     ["Domain:=", "Sweep"], ["Freq:=", ["All"]],
     ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal},{terminal}))"]], []),
    # Without context, without trailing empty
    ("S11_C", "Terminal Solution Data", "Rectangular Plot", "Setup1 : Sweep1",
     ["NAME:Context"], ["Freq:=", ["All"]],
     ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal},{terminal}))"]], []),
    # LastAdaptive
    ("S11_D", "Terminal Solution Data", "Rectangular Plot", "Setup1 : LastAdaptive",
     [], ["Freq:=", ["All"]],
     ["X Component:=", "Freq", "Y Component:=", [f"dB(St({terminal},{terminal}))"]], []),
]
for i, args in enumerate(cr_formats):
    try:
        oReport.CreateReport(*args)
        print(f"  CR Format {i}: OK!")
        csv_path = os.path.join(PROJECT_DIR, f"extract_{args[0]}.csv")
        try:
            oReport.ExportToFile(args[0], csv_path)
            print(f"  Exported: {csv_path}")
            with open(csv_path, 'r') as f:
                content = f.read()
            print(f"  Content:\n{content[:500]}")
        except Exception as e:
            print(f"  Export failed: {e}")
        break
    except Exception as e:
        print(f"  CR Format {i}: {e}")

# Method 3: GetSolveRangeInfo
print("\n--- M3: Solution info ---")
try:
    info = oSolutions.GetSolveRangeInfo("Setup1:Sweep1")
    print(f"  SolveRangeInfo: {info}")
except Exception as e:
    print(f"  GetSolveRangeInfo: {e}")

try:
    ndata = oSolutions.GetNetworkDataSolution("Setup1:Sweep1")
    print(f"  NetworkDataSolution: {ndata}")
except Exception as e:
    print(f"  GetNetworkDataSolution: {e}")

# Method 4: Try ExportForSpice
print("\n--- M4: ExportForSpice ---")
try:
    spice_path = os.path.join(PROJECT_DIR, "extract_test.sp")
    oSolutions.ExportForSpice("", "Setup1:Sweep1", spice_path)
    if os.path.exists(spice_path):
        print(f"  ExportForSpice OK!")
except Exception as e:
    print(f"  ExportForSpice: {e}")

print("\nDone.")
try: hfss.release_desktop()
except: pass
