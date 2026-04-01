"""Test: AutoIdentifyPorts on airbox faces to find the boundary face where
conductors (Gnd, Trace) exist. Standard WavePort at airbox boundary."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "airbox_face_test.log")

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
    projectname=os.path.join(PROJECT_DIR, "AirboxFace"),
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

# Geometry: microstrip with feed at airbox boundary y=-15
AIR_YMIN = -15.0

print("=== Creating geometry ===")
# Substrate: y from AIR_YMIN to AIR_YMIN+10mm
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", f"{AIR_YMIN}mm", "ZPosition:=", "0mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "1.6mm"],
    ["NAME:Attributes", "Name:=", "Sub", "Flags:=", "", "Color:=", "(0 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"FR4_epoxy"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Ground: y from AIR_YMIN (at boundary)
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", f"{AIR_YMIN}mm", "ZPosition:=", "-0.035mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Gnd", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Trace: y from AIR_YMIN (at boundary)
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-1.5mm", "YPosition:=", f"{AIR_YMIN}mm", "ZPosition:=", "1.6mm",
     "XSize:=", "3mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Trace", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# AirBox: y from -15 to 15
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

# Radiation boundary on Air
oBnd.AssignRadiation(
    ["NAME:Rad1",
     "Objects:=", ["Air"],
     "IsFssReference:=", False,
     "IsForPML:=", False])
print("  Radiation boundary OK")

# Get airbox face IDs
air_faces = oEditor.GetFaceIDs("Air")
print(f"\n  Air face IDs: {air_faces} ({len(air_faces)} faces)")

# Try AutoIdentifyPorts on EACH airbox face
# Use separate designs to avoid port overlap
print("\n=== Testing each Air face for WavePort ===")
success_faces = []
for i, fid_str in enumerate(air_faces):
    fid = int(fid_str)
    print(f"\n  Face {fid} (#{i+1}/{len(air_faces)}):")
    try:
        port_name = f"WP_{i}"
        oBnd.AutoIdentifyPorts(
            ["NAME:Faces", fid],
            True,  # WavePort
            ["NAME:ReferenceConductors", "Gnd"],
            port_name,
            True)
        print(f"    AutoIdentifyPorts: OK")
        
        # Check if it created excitations
        exc = oBnd.GetExcitations()
        print(f"    Excitations: {exc}")
        
        terms = oBnd.GetExcitationsOfType("Terminal")
        print(f"    Terminals: {terms}")
        
        if terms and len(terms) > 0:
            success_faces.append((fid, port_name, terms))
            print(f"    *** TERMINAL FOUND! ***")
        
        # Check messages for "No conductors"
        msgs = oDesktop.GetMessages("", "", 2)
        for m in msgs:
            if "conductor" in str(m).lower() or "port" in str(m).lower():
                print(f"    MSG: {m}")
        
    except Exception as e:
        print(f"    FAILED: {e}")

print(f"\n\n=== SUMMARY ===")
print(f"  Faces with terminals: {success_faces}")

if success_faces:
    # Now delete all ports except the first successful one
    # and try to simulate
    print(f"\n=== Using first successful face: {success_faces[0]} ===")
    
    # Delete all existing boundaries except Rad1
    try:
        all_exc = oBnd.GetExcitations()
        print(f"  Current excitations: {all_exc}")
    except: pass
    
    # Delete all ports/terminals
    for i in range(len(air_faces)):
        try:
            oBnd.DeleteBoundary([f"WP_{i}"])
        except: pass
    # Delete auto-created terminals
    try:
        terms_all = oBnd.GetExcitationsOfType("Terminal")
        for t in terms_all:
            try:
                oBnd.DeleteBoundary([str(t)])
            except: pass
    except: pass
    
    # Re-create only on the successful face
    fid, pname, _ = success_faces[0]
    print(f"  Creating WavePort on face {fid}...")
    oBnd.AutoIdentifyPorts(
        ["NAME:Faces", fid],
        True,
        ["NAME:ReferenceConductors", "Gnd"],
        "Port1",
        True)
    
    exc = oBnd.GetExcitations()
    print(f"  Excitations after: {exc}")
    terms = oBnd.GetExcitationsOfType("Terminal")
    print(f"  Terminals after: {terms}")
    
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
    
    v = oDesign.ValidateDesign()
    print(f"\n  Validation: {v}")
    
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        if any(k in str(m).lower() for k in ["terminal", "port", "error", "conduct"]):
            print(f"  MSG: {m}")
    
    if v:
        print("\n=== ANALYZE ===")
        t0 = time.time()
        try:
            oDesign.Analyze("Setup1")
            elapsed = time.time() - t0
            print(f"  *** SIMULATION SUCCEEDED in {elapsed:.1f}s! ***")
        except Exception as e:
            elapsed = time.time() - t0
            print(f"  FAILED ({elapsed:.1f}s): {e}")
            msgs = oDesktop.GetMessages("", "", 2)
            for m in msgs:
                if any(k in str(m).lower() for k in ["error", "terminal", "conduct", "pass"]):
                    print(f"  POST-MSG: {m}")

print("\nDone.")
try: hfss.release_desktop()
except: pass
