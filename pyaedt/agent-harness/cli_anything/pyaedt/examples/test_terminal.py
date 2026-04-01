"""Test: Manually add terminals after port creation in DrivenTerminal."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "terminal_test.log")

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
    projectname=os.path.join(PROJECT_DIR, "TermTest"),
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

# Create microstrip geometry
print("Creating geometry...")
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", "-5mm", "ZPosition:=", "0mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "1.6mm"],
    ["NAME:Attributes", "Name:=", "Sub", "Flags:=", "", "Color:=", "(0 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"FR4_epoxy"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-5mm", "YPosition:=", "-5mm", "ZPosition:=", "-0.035mm",
     "XSize:=", "10mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Gnd", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-1.5mm", "YPosition:=", "-5mm", "ZPosition:=", "1.6mm",
     "XSize:=", "3mm", "YSize:=", "10mm", "ZSize:=", "0.035mm"],
    ["NAME:Attributes", "Name:=", "Trace", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Port rect
oEditor.CreateRectangle(
    ["NAME:RectangleParameters",
     "IsCovered:=", True,
     "XStart:=", "-1.5mm", "YStart:=", "-5mm", "ZStart:=", "-0.035mm",
     "Width:=", "3mm", "Height:=", "1.67mm",
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

print("  Geometry OK")

# Get edge IDs
print("\nEdge analysis:")
try:
    prect_edges = oEditor.GetEdgeIDsOfObject("PRect")
    print(f"  PRect edges: {prect_edges}")
except Exception as e:
    print(f"  PRect edges error: {e}")

try:
    trace_edges = oEditor.GetEdgeIDsOfObject("Trace")
    print(f"  Trace edges: {trace_edges}")
except Exception as e:
    print(f"  Trace edges error: {e}")

try:
    gnd_edges = oEditor.GetEdgeIDsOfObject("Gnd")
    print(f"  Gnd edges: {gnd_edges}")
except Exception as e:
    print(f"  Gnd edges error: {e}")

# Get face IDs
try:
    prect_faces = oEditor.GetFaceIDs("PRect")
    print(f"  PRect faces: {prect_faces}")
except Exception as e:
    print(f"  PRect faces error: {e}")

# Check vertex positions of port rect to verify alignment
try:
    for fid in prect_faces:
        verts = oEditor.GetVertexIDsOfFace(int(fid))
        print(f"  Face {fid} vertices: {verts}")
        for vid in verts:
            pos = oEditor.GetVertexPosition(int(vid))
            print(f"    Vertex {vid}: {pos}")
except Exception as e:
    print(f"  Vertex error: {e}")

oBnd = oDesign.GetModule("BoundarySetup")

# --- Method A: Create port, then try to edit it to add terminals ---
print("\n--- Method A: Create port then add terminal ---")
try:
    oBnd.AssignLumpedPort(
        ["NAME:LP1",
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
    print("  Port LP1 created")
except Exception as e:
    print(f"  Port LP1 failed: {e}")

# Try editing the port to add terminals
print("  Trying EditLumpedPort...")
try:
    oBnd.EditLumpedPort("LP1",
        ["NAME:LP1",
         "Objects:=", ["PRect"],
         "RenormalizeAllTerminals:=", True,
         "DoDeembed:=", False,
         ["NAME:Terminals",
          ["NAME:LP1_T1",
           "TerminalType:=", "Auto",
           "SignalLine:=", "Trace",
           "ReferenceRE:=", "GND"]]])
    print("  EditLumpedPort succeeded!")
except Exception as e:
    print(f"  EditLumpedPort failed: {e}")

# Try AssignTerminal
print("  Trying AssignTerminal...")
try:
    edge_id = int(prect_edges[0]) if prect_edges else 0
    oBnd.AssignTerminal(
        ["NAME:T1",
         "Edges:=", [edge_id],
         "ParentBndID:=", "LP1",
         "TerminalResistance:=", "50ohm"])
    print(f"  AssignTerminal succeeded with edge {edge_id}!")
except Exception as e:
    print(f"  AssignTerminal failed: {e}")

# Try AutoAssignTerminals
print("  Trying different terminal commands...")
for cmd in ["AutoIdentifyTerminals", "AssignAutoTerminal", "SetTerminals"]:
    try:
        func = getattr(oBnd, cmd)
        print(f"    {cmd} exists")
    except:
        print(f"    {cmd} does not exist")

# Try to list all methods of oBnd
print("\n  BoundarySetup module methods:")
try:
    # Use dir() on COM object - might not work but worth trying
    import win32com.client
    # Try to enumerate type library
    print("  (Listing COM methods may not be available)")
    
    # Try various known methods
    for method in ["AssignLumpedPort", "AssignWavePort", "AssignTerminal", 
                   "EditLumpedPort", "AutoIdentifyPorts", "GetExcitations",
                   "GetBoundaries", "AssignLumpedRLC", "DeleteBoundary",
                   "GetExcitationsOfType", "GetBoundariesOfType",
                   "AutoAssignTerminals", "SetTerminals",
                   "EditTerminal", "AutoIdentifyTerminals"]:
        try:
            getattr(oBnd, method)
            print(f"    {method} - EXISTS")
        except:
            print(f"    {method} - N/A")
except: pass

# --- Method B: Try creating port with explicit terminal in creation ---
print("\n--- Method B: Port with explicit terminal ---")
try:
    oBnd.AssignLumpedPort(
        ["NAME:LP2",
         "Objects:=", ["PRect"],
         "RenormalizeAllTerminals:=", True,
         "DoDeembed:=", False,
         ["NAME:Terminals",
          ["NAME:LP2_T1",
           "Objects:=", ["Trace"],
           "ParentBndID:=", "LP2",
           "TerminalResistance:=", "50ohm"]]])
    print("  Method B succeeded!")
except Exception as e:
    print(f"  Method B failed: {e}")

# --- Method C: Try changing solution type to DrivenModal and re-assign port ---
print("\n--- Method C: Change to DrivenModal ---")
try:
    oDesign.SetSolutionType("DrivenModal")
    print("  Changed to DrivenModal")
    
    # Now try to validate
    v = oDesign.ValidateDesign()
    print(f"  Validation: {v}")
    
    # Check messages
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        if "LP1" in str(m) or "error" in str(m).lower():
            print(f"  MSG: {m}")
except Exception as e:
    print(f"  Method C failed: {e}")

# --- Method D: SetSolutionType after port creation ---
# If we can change from DrivenTerminal to DrivenModal AFTER port creation,
# the port might work in modal mode
print("\n--- Method D: Analyze in DrivenModal ---")
try:
    # Add radiation boundary
    oBnd.AssignRadiation(
        ["NAME:Rad1",
         "Objects:=", ["Air"],
         "IsFssReference:=", False,
         "IsForPML:=", False])
    print("  Radiation boundary OK")
except: 
    print("  Radiation already exists or failed")

try:
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
except:
    print("  Setup already exists or failed")

hfss.save_project()
print("  Saved")

# Validate in DrivenModal
v = oDesign.ValidateDesign()
print(f"  Validation (DrivenModal): {v}")
msgs = oDesktop.GetMessages("", "", 2)
for m in msgs:
    print(f"  MSG: {m}")

if True:
    print("\n  Analyzing in DrivenModal mode...")
    t0 = time.time()
    try:
        oDesign.Analyze("Setup1")
        print(f"  ANALYZE COMPLETED in {time.time()-t0:.1f}s!")
    except Exception as e:
        print(f"  ANALYZE FAILED ({time.time()-t0:.1f}s): {e}")

    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        if "error" in str(m).lower() or "complet" in str(m).lower():
            print(f"  MSG: {m}")

print("\nDone.")
try: hfss.release_desktop()
except: pass
