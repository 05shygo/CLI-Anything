"""Test: Port assignment in DrivenTerminal mode with different approaches."""
import os, sys
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "port_test2.log")

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

# Create fresh test project with DrivenTerminal
hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "Port_Test2"),
    designname="TestPort",
    solution_type="DrivenTerminal",
    non_graphical=True,
    new_desktop_session=False,
    specified_version="2019.1",
)
hfss.modeler.model_units = "mm"
oDesign = hfss.odesign
oEditor = oDesign.SetActiveEditor("3D Modeler")

# Create a simple substrate + ground + trace
print("Creating test geometry...")
oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-30", "YPosition:=", "-15", "ZPosition:=", "0",
     "XSize:=", "60", "YSize:=", "30", "ZSize:=", "1.6"],
    ["NAME:Attributes", "Name:=", "Sub", "Flags:=", "", "Color:=", "(143 175 131)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"FR4_epoxy"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-30", "YPosition:=", "-15", "ZPosition:=", "-0.035",
     "XSize:=", "60", "YSize:=", "30", "ZSize:=", "0.035"],
    ["NAME:Attributes", "Name:=", "GND", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-0.5", "YPosition:=", "-10", "ZPosition:=", "1.6",
     "XSize:=", "1", "YSize:=", "8", "ZSize:=", "0.035"],
    ["NAME:Attributes", "Name:=", "Trace", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])

# Port rect spanning from GND to Trace
oEditor.CreateRectangle(
    ["NAME:RectangleParameters",
     "IsCovered:=", True,
     "XStart:=", "-0.5", "YStart:=", "-10", "ZStart:=", "-0.035",
     "Width:=", "1", "Height:=", "1.67",
     "WhichAxis:=", "Y"],
    ["NAME:Attributes", "Name:=", "PortRect", "Flags:=", "", "Color:=", "(0 0 255)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '""', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", True, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  Geometry OK")

oBnd = oDesign.GetModule("BoundarySetup")
oDesktop = hfss.odesktop

# Test 1: DrivenTerminal with Modes syntax (this worked before for creating port, but no terminals)
print("\n--- Test 1: AssignLumpedPort with Modes (DrivenTerminal) ---")
try:
    oBnd.AssignLumpedPort(
        ["NAME:P1",
         "Objects:=", ["PortRect"],
         "RenormalizeAllTerminals:=", True,
         "DoDeembed:=", False,
         ["NAME:Modes",
          ["NAME:Mode1",
           "ModeNum:=", 1,
           "UseIntLine:=", True,
           ["NAME:IntLine",
            "Start:=", ["0", "-10", "-0.035"],
            "End:=",   ["0", "-10", "1.635"]],
           "CharImp:=", "Zpi"]]])
    print("  Port P1 created!")
    
    # Check excitations
    try:
        exc = oBnd.GetExcitations()
        print(f"  Excitations: {exc}")
    except: pass
    
    # Check for messages
    try:
        msgs = oDesktop.GetMessages("", "", 2)
        for m in msgs:
            print(f"  MSG: {m}")
    except: pass
    
    # Try to get terminals
    try:
        terms = oBnd.GetExcitationsOfType("Terminal")
        print(f"  Terminals: {terms}")
    except Exception as e:
        print(f"  GetExcitationsOfType error: {e}")
    
    oBnd.DeleteBoundary(["P1"])
    print("  Deleted P1")
except Exception as e:
    print(f"  Test 1 FAILED: {e}")

# Test 2: DrivenTerminal with Terminals syntax
print("\n--- Test 2: AssignLumpedPort with Terminals ---")
try:
    oBnd.AssignLumpedPort(
        ["NAME:P2",
         "Objects:=", ["PortRect"],
         "RenormalizeAllTerminals:=", True,
         "DoDeembed:=", False,
         ["NAME:Terminals",
          ["NAME:P2_T1",
           "TerminalType:=", "Auto",
           "SignalLine:=", "Trace",
           "ReferenceObjects:=", ["GND"],
           "UseIntLine:=", True,
           ["NAME:IntLine",
            "Start:=", ["0", "-10", "1.635"],
            "End:=",   ["0", "-10", "-0.035"]]]]])
    print("  Test 2 = SUCCESS")
    oBnd.DeleteBoundary(["P2"])
except Exception as e:
    print(f"  Test 2 FAILED: {e}")

# Test 3: AutoIdentifyPorts
print("\n--- Test 3: AutoIdentifyPorts ---")
try:
    # Get face IDs of the port rect
    face_ids = oEditor.GetFaceIDs("PortRect")
    print(f"  PortRect faces: {face_ids}")
    
    # Try auto-identify
    if face_ids:
        fid = int(face_ids[0]) if hasattr(face_ids[0], '__int__') else face_ids[0]
        oBnd.AutoIdentifyPorts(
            ["NAME:Faces", fid],
            True,
            ["NAME:ReferenceConductors", "GND"],
            "P3",
            True)
        print("  Test 3 = SUCCESS")
        
        exc = oBnd.GetExcitations()
        print(f"  Excitations: {exc}")
except Exception as e:
    print(f"  Test 3 FAILED: {e}")
    traceback.print_exc()

# Test 4: LumpedPort with Faces instead of Objects
print("\n--- Test 4: LumpedPort with Faces ---")
try:
    face_ids = oEditor.GetFaceIDs("PortRect")
    fid = int(face_ids[0]) if hasattr(face_ids[0], '__int__') else face_ids[0]
    oBnd.AssignLumpedPort(
        ["NAME:P4",
         "Faces:=", [fid],
         "RenormalizeAllTerminals:=", True,
         "DoDeembed:=", False,
         ["NAME:Modes",
          ["NAME:Mode1",
           "ModeNum:=", 1,
           "UseIntLine:=", True,
           ["NAME:IntLine",
            "Start:=", ["0", "-10", "-0.035"],
            "End:=",   ["0", "-10", "1.635"]],
           "CharImp:=", "Zpi"]]])
    print("  Test 4 = SUCCESS")
except Exception as e:
    print(f"  Test 4 FAILED: {e}")

# Test 5: LumpedPort with face of the Trace object (end face)
print("\n--- Test 5: LumpedPort on Trace face ---")
try:
    face_ids_trace = oEditor.GetFaceIDs("Trace")
    print(f"  Trace faces: {face_ids_trace}")
    
    # The face at y=-10 is the end face we want
    # Try each face to see which one works
    for fi in face_ids_trace:
        fid = int(fi) if hasattr(fi, '__int__') else fi
        try:
            oBnd.AssignLumpedPort(
                ["NAME:P5",
                 "Faces:=", [fid],
                 "RenormalizeAllTerminals:=", True,
                 "DoDeembed:=", False,
                 ["NAME:Modes",
                  ["NAME:Mode1",
                   "ModeNum:=", 1,
                   "UseIntLine:=", False,
                   "CharImp:=", "Zpi"]]])
            print(f"  Face {fid} = SUCCESS")
            oBnd.DeleteBoundary(["P5"])
            break
        except:
            pass
    else:
        print("  All trace faces failed")
except Exception as e:
    print(f"  Test 5 FAILED: {e}")
    traceback.print_exc()

# Test 6: Try WavePort
print("\n--- Test 6: WavePort ---")
try:
    oBnd.AssignWavePort(
        ["NAME:WP1",
         "Objects:=", ["PortRect"],
         "NumModes:=", 1,
         "RenormalizeAllTerminals:=", True,
         "DoDeembed:=", False,
         ["NAME:Modes",
          ["NAME:Mode1",
           "ModeNum:=", 1,
           "UseIntLine:=", True,
           ["NAME:IntLine",
            "Start:=", ["0", "-10", "-0.035"],
            "End:=",   ["0", "-10", "1.635"]],
           "CharImp:=", "Zpi"]]])
    print("  Test 6 = SUCCESS")
except Exception as e:
    print(f"  Test 6 FAILED: {e}")

# Test 7: LumpedRLC
print("\n--- Test 7: LumpedRLC ---")
try:
    oBnd.AssignLumpedRLC(
        ["NAME:LRLC1",
         "Objects:=", ["PortRect"],
         "RLC Type:=", "Serial",
         "Resistance:=", "50ohm",
         "Inductance:=", "0nH",
         "Capacitance:=", "0pF",
         "UseResist:=", True,
         "UseInduct:=", False,
         "UseCap:=", False])
    print("  Test 7 = SUCCESS")
except Exception as e:
    print(f"  Test 7 FAILED: {e}")

print("\nCheck validation status:")
try:
    v = oDesign.ValidateDesign()
    print(f"  Validation: {v}")
except: pass

try:
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        print(f"  {m}")
except: pass

print("\nDone.")
try: hfss.release_desktop()
except: pass
