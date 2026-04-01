"""Test: AutoIdentifyTerminals + VBS script approach for AEDT 2019.1."""
import os, sys, time
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "autoid_test.log")

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
    projectname=os.path.join(PROJECT_DIR, "AutoIDTest"),
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

# Geometry
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

# Port rectangle - make it slightly wider/taller than needed to ensure overlap
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

oBnd = oDesign.GetModule("BoundarySetup")

# ---- Test 1: AutoIdentifyTerminals on port face ----
print("\n--- Test 1: AssignLumpedPort then AutoIdentifyTerminals ---")
try:
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
    print("  Port1 created")
except Exception as e:
    print(f"  Port1 failed: {e}")

# Get face ID
prect_faces = oEditor.GetFaceIDs("PRect")
print(f"  PRect faces: {prect_faces}")
fid = int(prect_faces[0])

# Try AutoIdentifyTerminals
print("  Trying AutoIdentifyTerminals...")
for fmt in range(5):
    try:
        if fmt == 0:
            # Format 1: face list
            oBnd.AutoIdentifyTerminals(
                ["NAME:Faces", fid],
                True,
                ["NAME:ReferenceConductors", "Gnd"],
                "Port1",
                True)
        elif fmt == 1:
            # Format 2: just port name
            oBnd.AutoIdentifyTerminals("Port1")
        elif fmt == 2:
            # Format 3: objects
            oBnd.AutoIdentifyTerminals(
                ["NAME:Objects", "PRect"],
                True,
                ["NAME:ReferenceConductors", "Gnd"])
        elif fmt == 3:
            # Format 4: port + conductor info
            oBnd.AutoIdentifyTerminals(
                "Port1",
                ["NAME:ReferenceConductors", "Gnd"])
        elif fmt == 4:
            # Format 5: simple
            oBnd.AutoIdentifyTerminals(
                ["NAME:Faces", fid],
                False)
        print(f"  Format {fmt} = SUCCESS")
        break
    except Exception as e:
        print(f"  Format {fmt} FAILED: {e}")

# Check excitations
try:
    exc = oBnd.GetExcitations()
    print(f"  Excitations: {exc}")
except: pass

# ---- Test 2: RunScript with VBS ----
print("\n--- Test 2: VBS script approach ---")
vbs_path = os.path.join(PROJECT_DIR, "test_port.vbs")
vbs_code = '''
Dim oDesign
Set oDesign = GetActiveDesign()

Dim oModule
Set oModule = oDesign.GetModule("BoundarySetup")

' Try to auto-identify terminals on existing port
Dim oEditor
Set oEditor = oDesign.SetActiveEditor("3D Modeler")

' Check what objects exist
Dim solids
solids = oEditor.GetObjectsInGroup("Solids")

Dim sheets
sheets = oEditor.GetObjectsInGroup("Sheets")

' Create a lumped port with terminal definition
oModule.AssignLumpedPort Array("NAME:Port2", _
    "Objects:=", Array("PRect"), _
    "RenormalizeAllTerminals:=", True, _
    "DoDeembed:=", False, _
    Array("NAME:Modes", _
        Array("NAME:Mode1", _
            "ModeNum:=", 1, _
            "UseIntLine:=", True, _
            Array("NAME:IntLine", _
                "Start:=", Array("0mm", "-5mm", "1.635mm"), _
                "End:=", Array("0mm", "-5mm", "-0.035mm")), _
            "CharImp:=", "Zpi")))
'''
with open(vbs_path, 'w') as f:
    f.write(vbs_code)
print(f"  VBS written to {vbs_path}")

try:
    oDesktop.RunScript(vbs_path)
    print("  RunScript SUCCESS")
except Exception as e:
    print(f"  RunScript failed: {e}")

try:
    oDesktop.RunScriptWithArguments(vbs_path, "")
    print("  RunScriptWithArguments SUCCESS")
except Exception as e:
    print(f"  RunScriptWithArguments failed: {e}")

# ---- Test 3: AssignTerminal with face ---
print("\n--- Test 3: AssignTerminal with face ---")
try:
    oBnd.AssignTerminal(
        ["NAME:T1",
         "Faces:=", [fid],
         "ParentBndID:=", "Port1",
         "TerminalResistance:=", "50ohm"])
    print("  AssignTerminal face = SUCCESS")
except Exception as e:
    print(f"  AssignTerminal face failed: {e}")

# Try with Objects
try:
    oBnd.AssignTerminal(
        ["NAME:T2",
         "Objects:=", ["Trace"],
         "ParentBndID:=", "Port1",
         "TerminalResistance:=", "50ohm"])
    print("  AssignTerminal objects = SUCCESS")
except Exception as e:
    print(f"  AssignTerminal objects failed: {e}")

# Try with different params
try:
    oBnd.AssignTerminal(
        ["NAME:T3",
         "Objects:=", ["PRect"],
         "ParentBndID:=", "Port1",
         "TerminalResistance:=", "50ohm",
         "TerminalType:=", "Auto"])
    print("  AssignTerminal auto = SUCCESS")
except Exception as e:
    print(f"  AssignTerminal auto failed: {e}")

# Try signal/reference
try:
    oBnd.AssignTerminal(
        ["NAME:T4",
         "SignalLine:=", "Trace",
         "Ground:=", "Gnd",
         "ParentBndID:=", "Port1",
         "TerminalResistance:=", "50ohm"])
    print("  AssignTerminal signal/gnd = SUCCESS")
except Exception as e:
    print(f"  AssignTerminal signal/gnd failed: {e}")

# Check excitations
try:
    exc = oBnd.GetExcitations()
    print(f"  Excitations: {exc}")
except: pass

# Check if any terminal was created
try:
    terms = oBnd.GetExcitationsOfType("Terminal")
    print(f"  Terminals: {terms}")
except: pass

# ---- Test 4: Try directly manipulating the AEDT file ----
# Save project first
hfss.save_project()

# Read the aedt file to understand port structure
print("\n--- Test 4: Check AEDT file structure ---")
aedt_file = os.path.join(PROJECT_DIR, "AutoIDTest.aedt")
try:
    with open(aedt_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Find port section
    port_start = content.find("Port1")
    if port_start >= 0:
        section = content[max(0,port_start-200):port_start+500]
        print(f"  Port section:\n{section[:700]}")
    else:
        print("  Port1 not found in aedt file")
        
    # Find terminal section
    term_start = content.find("Terminal")
    if term_start >= 0:
        section = content[max(0,term_start-100):term_start+300]
        print(f"\n  Terminal section:\n{section[:400]}")
    else:
        print("  Terminal not found in aedt file")
except Exception as e:
    print(f"  File read error: {e}")

print("\n--- Final Messages ---")
try:
    msgs = oDesktop.GetMessages("", "", 2)
    for m in msgs:
        print(f"  {m}")
except: pass

print("\nDone.")
try: hfss.release_desktop()
except: pass
