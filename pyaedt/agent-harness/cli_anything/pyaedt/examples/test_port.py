"""Quick test: Try different LumpedPort formats in AEDT 2019.1 DrivenModal."""
import os, sys
os.environ['ANSYSEM_ROOT193'] = r'C:\Program Files\AnsysEM\AnsysEM19.3\Win64'
PROJECT_DIR = r"D:\class_design"
LOG = os.path.join(PROJECT_DIR, "port_test.log")

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

# Create fresh test project
hfss = Hfss(
    projectname=os.path.join(PROJECT_DIR, "Port_Test"),
    designname="TestPort",
    solution_type="DrivenModal",
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
print("  Substrate OK")

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-30", "YPosition:=", "-15", "ZPosition:=", "-0.035",
     "XSize:=", "60", "YSize:=", "30", "ZSize:=", "0.035"],
    ["NAME:Attributes", "Name:=", "GND", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  Ground OK")

oEditor.CreateBox(
    ["NAME:BoxParameters",
     "XPosition:=", "-0.5", "YPosition:=", "-10", "ZPosition:=", "1.6",
     "XSize:=", "1", "YSize:=", "8", "ZSize:=", "0.035"],
    ["NAME:Attributes", "Name:=", "Trace", "Flags:=", "", "Color:=", "(255 128 0)",
     "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "UDMId:=", "",
     "MaterialValue:=", '"copper"', "SurfaceMaterialValue:=", '""',
     "SolveInside:=", False, "IsMaterialEditable:=", True,
     "UseMaterialAppearance:=", False, "IsLightweight:=", False])
print("  Trace OK")

# Create port rectangle
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
print("  PortRect OK")

oBnd = oDesign.GetModule("BoundarySetup")

# Test Method 1: With integration line, no mm suffix
print("\n--- Method 1: No mm suffix ---")
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
    print("  Method 1 = SUCCESS")
except Exception as e:
    print(f"  Method 1 FAILED: {e}")
    # Delete port if it was partially created
    try:
        oBnd.DeleteBoundary(["P1"])
    except: pass

# Delete port for next test
try:
    oBnd.DeleteBoundary(["P1"])
except: pass

# Test Method 2: With mm suffix
print("\n--- Method 2: With mm suffix ---")
try:
    oBnd.AssignLumpedPort(
        ["NAME:P2",
         "Objects:=", ["PortRect"],
         "RenormalizeAllTerminals:=", True,
         "DoDeembed:=", False,
         ["NAME:Modes",
          ["NAME:Mode1",
           "ModeNum:=", 1,
           "UseIntLine:=", True,
           ["NAME:IntLine",
            "Start:=", ["0mm", "-10mm", "-0.035mm"],
            "End:=",   ["0mm", "-10mm", "1.635mm"]],
           "CharImp:=", "Zpi"]]])
    print("  Method 2 = SUCCESS")
except Exception as e:
    print(f"  Method 2 FAILED: {e}")

try:
    oBnd.DeleteBoundary(["P2"])
except: pass

# Test Method 3: Simpler - no integration line
print("\n--- Method 3: No integration line ---")
try:
    oBnd.AssignLumpedPort(
        ["NAME:P3",
         "Objects:=", ["PortRect"],
         "RenormalizeAllTerminals:=", True,
         "DoDeembed:=", False,
         ["NAME:Modes",
          ["NAME:Mode1",
           "ModeNum:=", 1,
           "UseIntLine:=", False,
           "CharImp:=", "Zpi"]]])
    print("  Method 3 = SUCCESS")
except Exception as e:
    print(f"  Method 3 FAILED: {e}")

try:
    oBnd.DeleteBoundary(["P3"])
except: pass

# Test Method 4: Without RenormalizeAllTerminals/DoDeembed
print("\n--- Method 4: Minimal params ---")
try:
    oBnd.AssignLumpedPort(
        ["NAME:P4",
         "Objects:=", ["PortRect"],
         ["NAME:Modes",
          ["NAME:Mode1",
           "ModeNum:=", 1,
           "UseIntLine:=", True,
           ["NAME:IntLine",
            "Start:=", ["0mm", "-10mm", "-0.035mm"],
            "End:=",   ["0mm", "-10mm", "1.635mm"]],
           "CharImp:=", "Zpi"]]])
    print("  Method 4 = SUCCESS")
except Exception as e:
    print(f"  Method 4 FAILED: {e}")

try:
    oBnd.DeleteBoundary(["P4"])
except: pass

# Test Method 5: LumpedRLC instead?
print("\n--- Method 5: With NumberOfModes ---")
try:
    oBnd.AssignLumpedPort(
        ["NAME:P5",
         "Objects:=", ["PortRect"],
         "DoDeembed:=", False,
         "RenormalizeAllTerminals:=", True,
         "NumberOfModes:=", 1,
         ["NAME:Modes",
          ["NAME:Mode1",
           "ModeNum:=", 1,
           "UseIntLine:=", True,
           ["NAME:IntLine",
            "Start:=", ["0mm", "-10mm", "-0.035mm"],
            "End:=",   ["0mm", "-10mm", "1.635mm"]],
           "CharImp:=", "Zpi",
           "AlignmentGroup:=", 0,
           "RenormImp:=", "50ohm"]]])
    print("  Method 5 = SUCCESS")
except Exception as e:
    print(f"  Method 5 FAILED: {e}")

try:
    oBnd.DeleteBoundary(["P5"])
except: pass

print("\nDone.")
try: hfss.release_desktop()
except: pass
