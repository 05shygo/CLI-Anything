"""Test: Use AutoIdentifyPorts with face ID - the CORRECT method for DrivenTerminal.
PyAEDT source shows _create_port_terminal uses AutoIdentify boundary type,
passing a single face ID with IsWavePort and ReferenceConductors."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "autoid_face_test.log")

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
    projectname=os.path.join(PROJECT_DIR, "AutoIDFace"),
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
# Sub: 10x10x1.6
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", "-5mm", "ZPosition:=", "0mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "1.6mm"],
    ["NAME:Attributes", "Name:=", "Sub", "Flags:=", "", "Color:=", "(0 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"FR4_epoxy"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Gnd: 10x10x0.035 on bottom
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", "-5mm", "ZPosition:=", "-0.035mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Gnd", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Trace: 3x10x0.035 on top
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-1.5mm", "YPosition:=", "-5mm", "ZPosition:=", "1.6mm",
     "XSize:=", "3mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Trace", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Port rectangle at y=-3, XZ plane, spanning from Gnd bottom to Trace top
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

# Radiation boundary
oBnd.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["Air"],
     "IsFssReference:=", False,
     "IsForPML:=", False])
print("  Geometry created")

# Get face IDs of PRect (port sheet)
print("\n=== Getting face IDs ===")
prect_faces = oEditor.GetFaceIDs("PRect")
print(f"  PRect faces: {prect_faces}")

trace_faces = oEditor.GetFaceIDs("Trace")
print(f"  Trace faces: {trace_faces}")

gnd_faces = oEditor.GetFaceIDs("Gnd")
print(f"  Gnd faces: {gnd_faces}")

# Try AutoIdentifyPorts with different faces
# Test 1: PRect face (the port sheet itself)
print("\n=== Method 1: AutoIdentifyPorts on PRect face (IsWavePort=False for Lumped) ===")
for i, fid in enumerate(prect_faces):
    fid_int = int(fid)
    print(f"  Trying face {fid_int}...")
    try:
        oBnd.AutoIdentifyPorts(
            ["NAME:Faces", fid_int],
            False,  # IsWavePort=False -> Lumped Port
            ["NAME:ReferenceConductors", "Gnd"],
            f"LP_{i}",
            True)
        print(f"    OK!")
    except Exception as e:
        print(f"    Failed: {e}")

# Test 2: PRect face with IsWavePort=True 
print("\n=== Method 2: AutoIdentifyPorts on PRect face (IsWavePort=True for Wave) ===")
for i, fid in enumerate(prect_faces):
    fid_int = int(fid)
    print(f"  Trying face {fid_int} as WavePort...")
    try:
        oBnd.AutoIdentifyPorts(
            ["NAME:Faces", fid_int],
            True,  # IsWavePort=True -> Wave Port
            ["NAME:ReferenceConductors", "Gnd"],
            f"WP_{i}",
            True)
        print(f"    OK!")
    except Exception as e:
        print(f"    Failed: {e}")

# Test 3: Try first Trace face
print("\n=== Method 3: AutoIdentifyPorts on first Trace face ===")
if trace_faces:
    fid_int = int(trace_faces[0])
    print(f"  Trying Trace face {fid_int}...")
    try:
        oBnd.AutoIdentifyPorts(
            ["NAME:Faces", fid_int],
            False,
            ["NAME:ReferenceConductors", "Gnd"],
            "TPort",
            True)
        print(f"    OK!")
    except Exception as e:
        print(f"    Failed: {e}")

# Check what excitations were created
print("\n=== Excitations ===")
try:
    exc = oBnd.GetExcitations()
    print(f"  GetExcitations: {exc}")
except Exception as e:
    print(f"  GetExcitations failed: {e}")

# Try GetExcitationsOfType
for etype in ["Terminal", "Lumped Port", "Wave Port", "LumpedPort", "WavePort"]:
    try:
        exc = oBnd.GetExcitationsOfType(etype)
        print(f"  GetExcitationsOfType('{etype}'): {exc}")
    except Exception as e:
        print(f"  GetExcitationsOfType('{etype}'): {e}")

# Try GetBoundaries
try:
    bnds = oBnd.GetBoundaries()
    print(f"  GetBoundaries: {bnds}")
except Exception as e:
    print(f"  GetBoundaries failed: {e}")

# Setup and validate
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

print("\n=== Validation ===")
v = oDesign.ValidateDesign()
print(f"  Result: {v}")

msgs = oDesktop.GetMessages("", "", 2)
for m in msgs:
    if any(k in str(m).lower() for k in ["terminal", "port", "error", "excit"]):
        print(f"  MSG: {m}")

if v:
    print("\n=== ANALYZE ===")
    try:
        oDesign.Analyze("Setup1")
        print("  *** SUCCESS! ***")
    except Exception as e:
        print(f"  FAILED: {e}")

print("\nDone.")
try: hfss.release_desktop()
except: pass
